#!/usr/bin/env python3
"""Live-demo builder: Microsoft Entra ID <-> Juniper Mist Access Assurance.

Builds, step by step with an explanation on screen for each stage:
  * the demo site (renames an existing site and sets its address/location)
  * Entra "IdP Connector" app registration (Graph permissions, admin consent, secret)
  * Entra onboarding group (+ demo users)
  * Mist OAuth/Azure identity provider + realm mapping
  * Mist internal (onboard) CA for BYOD certificates
  * a Multi-PSK SSID and an 802.1X SSID on Mist NAC, each in its own WLAN template,
    both applied to the demo site
  * Mist NAC tag + rule (EAP-TLS + Entra group -> allow)
  * Entra SAML policies (sign whole response; NameID = UPN in email format)
  * PSK self-service portal and certificate onboarding portal, each with an Entra SAML app
  * a Word report of everything done, with QR codes for both portals

Before the first change it saves a snapshot (<ini>.state.json). --teardown deletes
everything the build created and restores the site, CA and Access Assurance settings
from that snapshot, returning the org and tenant to their pre-demo state.

Usage:
    python3 mist_entra_access_assurance.py [--ini PATH] [--dry-run] [--pause]
    python3 mist_entra_access_assurance.py --teardown [--yes]
    python3 mist_entra_access_assurance.py --rotate-secret
"""
import argparse
import configparser
import json
import os
import re
import secrets
import sys
import time
from datetime import datetime, timedelta, timezone

import msal
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_INI = os.path.join(HERE, "mist_entra_access_assurance.ini")
TOKEN_CACHE = os.path.join(HERE, ".msal_token_cache.json")

GRAPH = "https://graph.microsoft.com/v1.0"
GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"
# Microsoft Graph Command Line Tools - Microsoft's public client for admin scripts
GRAPH_CLI_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
GRAPH_SCOPES = [
    "Application.ReadWrite.All",
    "AppRoleAssignment.ReadWrite.All",
    "DelegatedPermissionGrant.ReadWrite.All",
    "Group.ReadWrite.All",
    "Directory.Read.All",
    "Policy.ReadWrite.ApplicationConfiguration",
    "Policy.Read.All",
]
# Entra policies assigned to both portal SAML apps: (graph collection, name, definition, why)
SAML_POLICIES = [
    # Mist rejects SAML unless the whole Response is signed ("the response has to
    # be signed"); Entra's default signs only the assertion.
    ("tokenIssuancePolicies", "Mist SAML - sign response and assertion",
     json.dumps({"TokenIssuancePolicy": {
         "Version": 1,
         "SigningAlgorithm": "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256",
         "TokenResponseSigningPolicy": "ResponseAndToken",
         "SamlTokenVersion": "2.0"}}),
     "sign SAML response + assertion"),
    # Without this Entra sends an opaque pairwise NameID: PSKs get named after it and
    # the Marvis client fails "failed to get user profile data / requestFailed".
    # UPN as an emailAddress NameID also makes the cert CN carry the IdP realm.
    ("claimsMappingPolicies", "Mist SAML - NameID is UPN (email format)",
     json.dumps({"ClaimsMappingPolicy": {
         "Version": 1, "IncludeBasicClaimSet": "true",
         "ClaimsSchema": [{
             "Source": "user", "ID": "userprincipalname",
             "SamlClaimType": "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/nameidentifier",
             "SamlNameIdFormat": "urn:oasis:names:tc:SAML:1.1:nameid-format:emailAddress"}]}}),
     "NameID = user principal name, email format"),
]
NON_GALLERY_TEMPLATE = "8adf8e6e-67b2-4cf2-a259-e3dc5476c621"  # "Create your own application"
IDP_APP_ROLES = ["User.Read.All", "GroupMember.Read.All", "Device.Read.All"]
IDP_DELEGATED = "User.Read"
ZERO_GUID = "00000000-0000-0000-0000-000000000000"
MIST_NAME_MAX = 32  # Mist UI limit for labels / object names
DRY = "<new>"


class BuildError(Exception):
    pass


# --------------------------------------------------------------------------- screen output
COLOR = sys.stdout.isatty() and os.environ.get("NO_COLOR") is None


def _c(code, text):
    return f"\033[{code}m{text}\033[0m" if COLOR else text


# One timestamp per run, shared by the Word report and the audit log file names.
RUN_TS = datetime.now()
AUDIT_LOG = None
_ANSI = re.compile(r"\033\[[0-9;]*m")


class _Tee:
    """Copy everything printed to the screen into the audit log, without colours and
    with a time on every line."""

    def __init__(self, real, fh):
        self.real, self.fh, self.bol = real, fh, True

    def write(self, s):
        self.real.write(s)
        for part in _ANSI.sub("", s).splitlines(keepends=True):
            if self.bol:
                self.fh.write(datetime.now().strftime("%H:%M:%S  "))
            self.fh.write(part)
            self.bol = part.endswith("\n")
        self.fh.flush()
        return len(s)

    def flush(self):
        self.real.flush()
        self.fh.flush()

    def isatty(self):
        return self.real.isatty()

    def __getattr__(self, name):
        return getattr(self.real, name)


def stamped(path, ts=RUN_TS):
    """Report.docx -> Report_20261119-143005.docx"""
    base, ext = os.path.splitext(path)
    return f"{base}_{ts:%Y%m%d-%H%M%S}{ext}"


def start_audit_log(cfg, ini_path, mode, argv):
    p = cfg.get("output", "audit_log", fallback="").strip() if cfg.has_section("output") else ""
    p = os.path.expanduser(p or "Mist_AA_Audit.log")
    if not os.path.isabs(p):
        p = os.path.join(os.path.dirname(os.path.abspath(ini_path)), p)
    base, ext = os.path.splitext(p)
    path = f"{base}_{RUN_TS:%Y%m%d-%H%M%S}_{mode}{ext or '.log'}"
    fh = open(path, "w", encoding="utf-8")
    fh.write(f"Mist Access Assurance + Entra ID demo - audit log\n"
             f"Run:     {mode}\nStarted: {RUN_TS:%Y-%m-%d %H:%M:%S} {time.strftime('%Z')}\n"
             f"Ini:     {os.path.abspath(ini_path)}\nCommand: {' '.join(argv)}\n{'=' * 72}\n")
    sys.stdout = _Tee(sys.stdout, fh)
    sys.stderr = _Tee(sys.stderr, fh)
    return path


GOOD = {"created", "uploaded", "enabled", "granted", "assigned", "set", "updated", "renamed", "added",
        "deleted", "removed", "restored", "disabled", "purged", "revoked"}


class Run:
    """Stage banners on screen + event log for the Word report."""

    def __init__(self, total, pause=False):
        self.total, self.pause, self.n = total, pause, 0
        self.stages, self.cur, self.t0 = [], None, None

    def stage(self, title, explain):
        self.end()
        self.n += 1
        if self.pause and self.n > 1:
            input(_c("2", "   Press Enter for the next step..."))
        print()
        print(_c("1;36", f"━━ Step {self.n}/{self.total}  {title} ") + _c("36", "━" * max(3, 60 - len(title))))
        print(_c("2", f"   {explain}"))
        self.cur = {"title": title, "explain": explain, "events": [], "secs": 0}
        self.t0 = time.time()

    def end(self):
        if self.cur:
            self.cur["secs"] = time.time() - self.t0
            self.stages.append(self.cur)
            print(_c("2", f"   done in {self.cur['secs']:.1f}s"))
            self.cur = None

    def log(self, status, kind, name, extra=""):
        if status in GOOD:
            sym, col = "✓", "32"
        elif status == "exists":
            sym, col = "•", "2"
        elif status.startswith("would"):
            sym, col = "○", "33"
        elif status == "note":
            sym, col = "i", "33"
        else:
            sym, col = "!", "31"
        line = f"   {_c(col, sym)} {_c(col, status.ljust(9))} {kind}: {_c('1', name)}"
        print(line + (_c("2", f"  ({extra})") if extra else ""))
        if self.cur is not None:
            self.cur["events"].append((status, kind, name, extra))


RUN = Run(0)


def log(status, kind, name, extra=""):
    RUN.log(status, kind, name, extra)


# --------------------------------------------------------------------------- state (snapshot)
class State:
    """Snapshot of the org/tenant before the demo + everything the build created."""

    def __init__(self, path, dry=False):
        self.path, self.dry = path, dry
        self.data = json.load(open(path)) if os.path.exists(path) else None

    @property
    def exists(self):
        return self.data is not None

    def start(self, snapshot):
        if self.data is None:
            self.data = {"snapshot": snapshot, "created": []}
            self.save()

    def record(self, system, kind, oid, name, **extra):
        if self.dry or self.data is None:
            return
        self.data["created"].append({"system": system, "kind": kind, "id": oid, "name": name, **extra})
        self.save()

    def created(self, system, kind):
        return [c for c in (self.data or {}).get("created", []) if c["system"] == system and c["kind"] == kind]

    def save(self):
        if self.dry:
            return
        tmp = self.path + ".tmp"
        with open(tmp, "w") as f:
            json.dump(self.data, f, indent=1)
        os.replace(tmp, self.path)

    def finish(self):
        done = self.path.replace(".state.json", f".state.restored-{datetime.now():%Y%m%d-%H%M%S}.json")
        os.replace(self.path, done)
        return done


STATE = None


# --------------------------------------------------------------------------- config
def load_config(path):
    if not os.path.exists(path):
        raise BuildError(f"ini not found: {path} (copy mist_entra_access_assurance.ini.example)")
    cfg = configparser.ConfigParser(inline_comment_prefixes=(";", "#"))
    cfg.read(path)
    for sec in ("mist", "entra", "site", "wlan", "nac", "psk_portal", "cert_portal"):
        if not cfg.has_section(sec):
            raise BuildError(f"ini is missing [{sec}]")
    return cfg


def names(cfg):
    """All object names, with Mist names checked against the 32-character UI limit."""
    e = cfg["entra"]
    prefix = e.get("name_prefix", "").strip() or "Mist AA"
    n = {
        "prefix": prefix,
        "group": e["user_group"].strip(),
        "entra_idp_app": f"{prefix} IdP Connector",
        "entra_psk_app": f"{prefix} PSK Portal",
        "entra_cert_app": f"{prefix} Cert Portal",
        "mist_idp": cfg["mist"].get("idp_name", "").strip() or "Entra ID",
        "psk_ssid": cfg["wlan"]["psk_ssid"].strip(),
        "cert_ssid": cfg["wlan"]["cert_ssid"].strip(),
        "nac_tag": cfg["nac"].get("tag_name", "").strip() or f"Entra grp: {e['user_group'].strip()}",
        "nac_rule": cfg["nac"].get("rule_name", "").strip() or "EAP-TLS Entra users",
        "psk_portal": cfg["psk_portal"].get("name", "").strip() or "PSK Self-Service",
        "cert_portal": cfg["cert_portal"].get("name", "").strip() or "Certificate Onboarding",
        "site": cfg["site"]["name"].strip(),
    }
    # One WLAN template per SSID, both applied to the demo site.
    n["psk_template"] = cfg["wlan"].get("psk_template", "").strip() or f"AA - {n['psk_ssid']}"
    n["cert_template"] = cfg["wlan"].get("cert_template", "").strip() or f"AA - {n['cert_ssid']}"
    n["floorplan"] = cfg["site"].get("floorplan_name", "").strip() or "Ground Floor"
    # Roles: Access Assurance returns one (RADIUS group) for 802.1X and OpenRoaming; the PSK portal
    # stamps one on every PSK it issues.
    n["corp_role"] = cfg["nac"].get("role", "").strip() or "Corp"
    n["psk_role"] = cfg["psk_portal"].get("role", "").strip() or "BYOD"
    # Optional second PSK portal on the same SSID: its PSKs get a role that a WxLAN policy in the
    # PSK SSID's template limits to the internet (no local network).
    i = cfg["psk_portal_isolated"] if cfg.has_section("psk_portal_isolated") else {}
    n["iso_enabled"] = cfg.getboolean("psk_portal_isolated", "enabled", fallback=False)
    n["iso_portal"] = (i.get("name", "") or "").strip() or "BYoD_Isolated"
    n["iso_role"] = (i.get("role", "") or "").strip() or "BYoD_Isolated"
    n["iso_vlan"] = (i.get("vlan_id", "") or "").strip()
    n["entra_iso_app"] = f"{prefix} PSK Portal Isolated"
    n["iso_user_tag"], n["internet_tag"], n["local_tag"] = "Isolated BYOD Users", "Internet", "Local Network"
    # Sample PSKs (one per PSK role) shown as Wi-Fi QR codes in the report.
    n["samples_enabled"] = cfg.getboolean("sample_psks", "enabled", fallback=False)
    if n["iso_enabled"] and (n["iso_portal"] == n["psk_portal"] or n["iso_role"] == n["psk_role"]):
        raise BuildError("[psk_portal_isolated] name and role must differ from [psk_portal]")
    # Optional OpenRoaming SSID ([openroaming] enabled = true)
    o = cfg["openroaming"] if cfg.has_section("openroaming") else {}
    n["or_enabled"] = cfg.getboolean("openroaming", "enabled", fallback=False)
    n["or_ssid"] = (o.get("ssid", "") or "").strip() or "OpenRoaming"
    n["or_template"] = (o.get("template", "") or "").strip() or f"AA - {n['or_ssid']}"
    n["or_idp"] = (o.get("idp_name", "") or "").strip() or "OpenRoaming"
    n["or_tag"] = (o.get("tag_name", "") or "").strip() or f"SSID: {n['or_ssid']}"
    n["or_rule"] = (o.get("rule_name", "") or "").strip() or "OpenRoaming SSID"
    n["or_role"] = (o.get("role", "") or "").strip() or "Guest"
    n["corp_role_tag"], n["or_role_tag"] = f"Role: {n['corp_role']}", f"Role: {n['or_role']}"
    if n["or_enabled"] and n["or_role"] == n["corp_role"]:
        raise BuildError("[openroaming] role must differ from [nac] role")
    templates = [n["psk_template"], n["cert_template"]] + ([n["or_template"]] if n["or_enabled"] else [])
    if len(set(templates)) != len(templates):
        raise BuildError("WLAN template names must all be different ([wlan] psk_template / cert_template, "
                         "[openroaming] template)")
    if n["or_enabled"] and n["or_ssid"] in (n["psk_ssid"], n["cert_ssid"]):
        raise BuildError("[openroaming] ssid must differ from the PSK and 802.1X SSIDs")
    keys = ["mist_idp", "floorplan", "psk_template", "cert_template", "psk_ssid", "cert_ssid", "nac_tag", "nac_rule",
            "psk_portal", "cert_portal", "corp_role_tag", "psk_role"]
    if n["iso_enabled"]:
        keys += ["iso_portal", "iso_role"]
    if n["or_enabled"]:
        keys += ["or_ssid", "or_template", "or_idp", "or_tag", "or_rule", "or_role_tag"]
    too_long = [f"{k} '{n[k]}' ({len(n[k])})" for k in keys if len(n[k]) > MIST_NAME_MAX]
    if too_long:
        raise BuildError(f"Mist names must be <= {MIST_NAME_MAX} characters - shorten in the ini: "
                         + "; ".join(too_long))
    return n


def mist_credentials(cfg):
    m = cfg["mist"]
    cloud, token = m.get("cloud", "").strip(), m.get("token", "").strip()
    if not (cloud and token):
        store = os.path.expanduser(m.get("credentials", "").strip())
        label = m.get("label", "").strip()
        if not (store and label):
            raise BuildError("[mist] needs cloud + token (or credentials + label)")
        sc = configparser.ConfigParser()
        sc.read(store)
        if not sc.has_section(label):
            raise BuildError(f"label [{label}] not in {store}")
        cloud, token = sc[label].get("cloud", "").strip(), sc[label].get("token", "").strip()
    cloud = cloud.replace("https://", "").rstrip("/")
    if cloud.startswith("manage."):
        cloud = "api." + cloud[len("manage."):]
    return cloud, token


def build_session(cfg):
    s = requests.Session()
    if cfg.has_section("network"):
        net = cfg["network"]
        ca = net.get("ca_bundle", "").strip()
        if ca:
            s.verify = os.path.expanduser(ca)
        elif not net.getboolean("verify_ssl", fallback=True):
            s.verify = False
            requests.packages.urllib3.disable_warnings()
        proxies = {k: net.get(f"{k}_proxy", "").strip() for k in ("http", "https")}
        proxies = {k: v for k, v in proxies.items() if v}
        if proxies:
            s.proxies.update(proxies)
    return s


def csv_list(value):
    return [x.strip() for x in (value or "").split(",") if x.strip()]


def int_or_none(value):
    value = (value or "").strip()
    return int(value) if value else None


# --------------------------------------------------------------------------- Mist API
class Mist:
    def __init__(self, session, cloud, token, org_id=None):
        self.s, self.cloud = session, cloud
        self.base = f"https://{cloud}/api/v1"
        self.h = {"Authorization": f"Token {token}"}
        self.org_id = org_id or self._single_org()
        self.org = f"{self.base}/orgs/{self.org_id}"

    def req(self, method, url, body=None, ok404=False):
        r = self.s.request(method, url, headers=self.h, json=body, timeout=60)
        if ok404 and r.status_code == 404:
            return None
        if not r.ok:
            raise BuildError(f"Mist {method} {url.replace(self.base, '')} -> {r.status_code}: {r.text[:300]}")
        return r.json() if r.text else {}

    def _single_org(self):
        me = self.req("GET", f"{self.base}/self")
        orgs = [p["org_id"] for p in me.get("privileges", [])
                if p.get("scope") == "org" and p.get("role") in ("admin", "write")]
        if len(orgs) != 1:
            raise BuildError(f"token has write access to {len(orgs)} orgs - set [mist] org_id")
        return orgs[0]

    def list(self, ep):
        return self.req("GET", f"{self.org}/{ep}?limit=1000")

    def find(self, ep, value, key="name"):
        hits = [o for o in self.list(ep) if o.get(key) == value]
        if len(hits) > 1:
            raise BuildError(f"{len(hits)} Mist {ep} named '{value}' - remove duplicates first")
        return hits[0] if hits else None

    def create(self, ep, body):
        return self.req("POST", f"{self.org}/{ep}", body)

    def update(self, ep, oid, body):
        return self.req("PUT", f"{self.org}/{ep}/{oid}", body)

    def delete(self, ep, oid):
        return self.req("DELETE", f"{self.org}/{ep}/{oid}", ok404=True)


# --------------------------------------------------------------------------- Graph API
def token_scopes(result):
    """Scopes granted in a Graph access token (its 'scp' claim)."""
    import base64
    try:
        payload = result["access_token"].split(".")[1]
        claims = json.loads(base64.urlsafe_b64decode(payload + "=" * (-len(payload) % 4)))
        return set(claims.get("scp", "").lower().split())
    except (KeyError, IndexError, ValueError):
        return set()


class Graph:
    def __init__(self, session, tenant):
        self.s = session
        cache = msal.SerializableTokenCache()
        if os.path.exists(TOKEN_CACHE):
            cache.deserialize(open(TOKEN_CACHE).read())
        app = msal.PublicClientApplication(
            GRAPH_CLI_CLIENT_ID, authority=f"https://login.microsoftonline.com/{tenant}",
            token_cache=cache, http_client=session)
        result = None
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
            if result and not all(s.lower() in token_scopes(result) for s in GRAPH_SCOPES):
                result = None  # cached token lacks a newly added scope - sign in again
        if not result:
            flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)
            if "user_code" not in flow:
                raise BuildError(f"device flow failed: {flow.get('error_description', flow)}")
            print(_c("1;33", "\n   Entra sign-in needed (Global Admin):"))
            print("   " + flow["message"] + "\n")
            while True:  # idle keep-alive connections get dropped while you sign in
                try:
                    result = app.acquire_token_by_device_flow(flow)
                    break
                except requests.exceptions.ConnectionError:
                    if time.time() > flow.get("expires_at", 0):
                        raise BuildError("device code expired - re-run and sign in sooner")
                    time.sleep(2)
        if "access_token" not in result:
            raise BuildError(f"Entra sign-in failed: {result.get('error_description', result)}")
        if cache.has_state_changed:
            with open(TOKEN_CACHE, "w") as f:
                f.write(cache.serialize())
            os.chmod(TOKEN_CACHE, 0o600)
        self.h = {"Authorization": f"Bearer {result['access_token']}"}

    def req(self, method, path, body=None, retry=True, ok404=False):
        """Call Graph, retrying while a just-created object replicates."""
        url = path if path.startswith("http") else GRAPH + path
        for _ in range(12):
            r = self.s.request(method, url, headers=self.h, json=body, timeout=60)
            if r.ok:
                return r.json() if r.text else {}
            if ok404 and r.status_code == 404:
                return None
            transient = r.status_code in (404, 429) or (
                r.status_code == 400 and re.search(r"does not exist|not found|Request_ResourceNotFound",
                                                   r.text, re.I))
            if not (retry and transient):
                break
            time.sleep(5)
        raise BuildError(f"Graph {method} {path} -> {r.status_code}: {r.text[:400]}")

    def one(self, path, what):
        hits = self.req("GET", path, retry=False).get("value", [])
        if len(hits) > 1:
            raise BuildError(f"{len(hits)} Entra {what} - remove duplicates first")
        return hits[0] if hits else None


def odata(value):
    return value.replace("'", "''")


# --------------------------------------------------------------------------- site
def find_site(mist, cfg, n):
    sites = mist.req("GET", f"{mist.org}/sites?limit=1000")
    by_name = {s["name"]: s for s in sites}
    cur = cfg["site"].get("current_name", "").strip()
    site = by_name.get(n["site"]) or by_name.get(cur)
    if not site:
        raise BuildError(f"[site] neither '{n['site']}' nor current_name '{cur}' exists in the org")
    return site


def site_snapshot(site):
    return {k: site.get(k) for k in ("id", "name", "address", "latlng", "country_code", "timezone")}


def step_site(mist, cfg, n, site, dry):
    s = cfg["site"]
    want = {"name": n["site"], "address": s.get("address", "").strip() or None}
    if s.get("country_code", "").strip():
        want["country_code"] = s["country_code"].strip().upper()
    if s.get("timezone", "").strip():
        want["timezone"] = s["timezone"].strip()
    if s.get("latitude", "").strip() and s.get("longitude", "").strip():
        want["latlng"] = {"lat": float(s["latitude"]), "lng": float(s["longitude"])}
    diff = {k: v for k, v in want.items() if site.get(k) != v}
    if not diff:
        log("exists", "site", n["site"], want.get("address") or "")
        return site
    if dry:
        log("would set", "site", f"{site['name']} -> {n['site']}", want.get("address") or "")
        return {**site, **want}
    site = mist.req("PUT", f"{mist.base}/sites/{site['id']}", diff)
    if "name" in diff:
        log("renamed", "site", n["site"], f"was '{STATE.data['snapshot']['site']['name']}'")
    rest = [f"{k}={v if k != 'latlng' else '%.4f,%.4f' % (v['lat'], v['lng'])}"
            for k, v in diff.items() if k != "name"]
    if rest:
        log("set", "site location", n["site"], ", ".join(rest))
    return site


DEFAULT_FLOORPLAN = os.path.join(HERE, "sample_office_floorplan.png")  # AI-generated 40 x 25 m office
DEFAULT_FLOORPLAN_PPM = 50


def step_floorplan(mist, cfg, n, site, ini_path, dry):
    """Upload the floorplan image as a map on the demo site. [site] floorplan blank = the
    bundled AI-generated plan; a path = that image (needs floorplan_ppm); none = skip."""
    s = cfg["site"]
    value = s.get("floorplan", "").strip()
    ppm = s.get("floorplan_ppm", "").strip()
    if value.lower() in ("none", "off", "no"):
        log("note", "floorplan", "none", "[site] floorplan = none")
        return
    if value:  # override: your own image, its scale must be given
        path = os.path.expanduser(value)
        if not os.path.isabs(path):
            path = os.path.join(os.path.dirname(os.path.abspath(ini_path)), path)
        if not ppm:
            raise BuildError("[site] floorplan_ppm is required with your own floorplan (image pixels per metre)")
    else:  # default: the bundled AI-generated office plan
        path, ppm = DEFAULT_FLOORPLAN, ppm or DEFAULT_FLOORPLAN_PPM
    if not os.path.isfile(path):
        raise BuildError(f"[site] floorplan: {path} not found")
    try:
        ppm = float(ppm)
    except ValueError:
        raise BuildError(f"[site] floorplan_ppm must be a number, not '{ppm}'")
    if ppm <= 0:
        raise BuildError("[site] floorplan_ppm must be greater than 0")
    from PIL import Image
    with Image.open(path) as im:
        width, height = im.size
    size = f"{width / ppm:.0f} m x {height / ppm:.0f} m, {ppm:g} px = 1 m"
    if site.get("id") and any(m.get("name") == n["floorplan"]
                              for m in mist.req("GET", f"{mist.base}/sites/{site['id']}/maps")):
        log("exists", "floorplan", n["floorplan"])
        return
    if dry:
        log("would upload", "floorplan", n["floorplan"], f"{os.path.basename(path)}, {size}")
        return
    m = mist.req("POST", f"{mist.base}/sites/{site['id']}/maps",
                 {"name": n["floorplan"], "type": "image", "width": width, "height": height, "ppm": ppm})
    STATE.record("mist", "maps", m["id"], n["floorplan"], site_id=site["id"])
    with open(path, "rb") as f:
        r = mist.s.post(f"{mist.base}/sites/{site['id']}/maps/{m['id']}/image", headers=mist.h,
                        files={"file": (os.path.basename(path), f, "image/png" if path.lower().endswith(".png")
                                        else "image/jpeg")}, timeout=120)
    if not r.ok:
        raise BuildError(f"Mist floorplan image upload -> {r.status_code}: {r.text[:300]}")
    log("uploaded", "floorplan", n["floorplan"], f"{os.path.basename(path)}, {size}")


# --------------------------------------------------------------------------- Entra steps
def entra_find_app(g, name):
    return g.one(f"/applications?$filter=displayName eq '{odata(name)}'", f"apps named '{name}'")


def entra_find_sp(g, app_id):
    return g.one(f"/servicePrincipals?$filter=appId eq '{app_id}'", f"service principals for {app_id}")


def step_idp_app(g, n, dry):
    graph_sp = entra_find_sp(g, GRAPH_APP_ID)
    roles = {r["value"]: r["id"] for r in graph_sp["appRoles"]}
    scopes = {s["value"]: s["id"] for s in graph_sp["oauth2PermissionScopes"]}
    access = {"resourceAppId": GRAPH_APP_ID, "resourceAccess":
              [{"id": scopes[IDP_DELEGATED], "type": "Scope"}] +
              [{"id": roles[r], "type": "Role"} for r in IDP_APP_ROLES]}
    body = {"isFallbackPublicClient": True, "requiredResourceAccess": [access]}
    name = n["entra_idp_app"]

    app = entra_find_app(g, name)
    if app:
        log("exists", "app registration", name, f"client ID {app['appId']}")
        if not dry:
            g.req("PATCH", f"/applications/{app['id']}", body)
    elif dry:
        log("would create", "app registration", name)
        app = {"id": DRY, "appId": DRY}
    else:
        app = g.req("POST", "/applications", {"displayName": name, "signInAudience": "AzureADMyOrg", **body})
        STATE.record("entra", "application", app["id"], name, appId=app["appId"])
        log("created", "app registration", name, f"client ID {app['appId']}")
        log("set", "public client flows", name, "on - lets Mist check passwords for EAP-TTLS (ROPC)")

    sp = entra_find_sp(g, app["appId"]) if app["appId"] != DRY else None
    if not sp and not dry:
        sp = g.req("POST", "/servicePrincipals", {"appId": app["appId"]})
    sp = sp or {"id": DRY}

    have = set()
    if sp["id"] != DRY:
        have = {a["appRoleId"] for a in g.req("GET", f"/servicePrincipals/{sp['id']}/appRoleAssignments")
                .get("value", []) if a["resourceId"] == graph_sp["id"]}
    for r in IDP_APP_ROLES:
        if roles[r] in have:
            continue
        if dry:
            log("would grant", "admin consent", f"Microsoft Graph {r} (application)")
            continue
        g.req("POST", f"/servicePrincipals/{graph_sp['id']}/appRoleAssignedTo",
              {"principalId": sp["id"], "resourceId": graph_sp["id"], "appRoleId": roles[r]})
        log("granted", "admin consent", f"Microsoft Graph {r} (application)")

    grant = None
    if sp["id"] != DRY:
        grant = g.one(f"/oauth2PermissionGrants?$filter=clientId eq '{sp['id']}' and "
                      f"resourceId eq '{graph_sp['id']}'", "delegated grants")
    if not grant or IDP_DELEGATED not in (grant.get("scope") or "").split():
        if dry:
            log("would grant", "admin consent", f"Microsoft Graph {IDP_DELEGATED} (delegated)")
        else:
            if grant:
                g.req("PATCH", f"/oauth2PermissionGrants/{grant['id']}",
                      {"scope": f"{grant.get('scope', '')} {IDP_DELEGATED}".strip()})
            else:
                g.req("POST", "/oauth2PermissionGrants",
                      {"clientId": sp["id"], "consentType": "AllPrincipals",
                       "resourceId": graph_sp["id"], "scope": IDP_DELEGATED})
            log("granted", "admin consent", f"Microsoft Graph {IDP_DELEGATED} (delegated)")
    return app


def entra_new_secret(g, app, months):
    end = datetime.now(timezone.utc) + timedelta(days=round(months * 30.44))
    res = g.req("POST", f"/applications/{app['id']}/addPassword",
                {"passwordCredential": {"displayName": "Mist Access Assurance",
                                        "endDateTime": end.strftime("%Y-%m-%dT%H:%M:%SZ")}})
    log("created", "client secret", "IdP Connector", f"expires {end:%Y-%m-%d}; sent straight to Mist, never shown")
    return res["secretText"]


def step_group(g, cfg, n, dry):
    name = n["group"]
    grp = g.one(f"/groups?$filter=displayName eq '{odata(name)}'", f"groups named '{name}'")
    if grp:
        log("exists", "security group", name)
    elif dry:
        log("would create", "security group", name)
        grp = {"id": DRY, "displayName": name}
    else:
        nick = re.sub(r"[^A-Za-z0-9]", "", name)[:60] or "MistWiFiUsers"
        grp = g.req("POST", "/groups", {"displayName": name, "mailEnabled": False, "mailNickname": nick,
                                        "securityEnabled": True,
                                        "description": "Users allowed to onboard to Mist Wi-Fi"})
        STATE.record("entra", "group", grp["id"], name)
        log("created", "security group", name)

    members = set()
    if grp["id"] != DRY:
        members = {m.get("userPrincipalName", "").lower() for m in
                   g.req("GET", f"/groups/{grp['id']}/members?$select=userPrincipalName").get("value", [])}
    for upn in csv_list(cfg["entra"].get("demo_users")):
        if upn.lower() in members:
            log("exists", "group member", upn)
            continue
        if dry:
            log("would add", "group member", upn)
            continue
        user = g.req("GET", f"/users/{upn}?$select=id,userPrincipalName", retry=False, ok404=True)
        if not user:
            raise BuildError(f"[entra] demo_users: {upn} not found in the tenant")
        g.req("POST", f"/groups/{grp['id']}/members/$ref",
              {"@odata.id": f"{GRAPH}/directoryObjects/{user['id']}"})
        STATE.record("entra", "group_member", user["id"], upn, group_id=grp["id"])
        log("added", "group member", upn, f"to {name}")
    return grp


def pem(b64):
    b64 = re.sub(r"\s", "", b64)
    return ("-----BEGIN CERTIFICATE-----\n" + "\n".join(b64[i:i + 64] for i in range(0, len(b64), 64))
            + "\n-----END CERTIFICATE-----")


def metadata_cert(session, tenant_id, app_id):
    url = f"https://login.microsoftonline.com/{tenant_id}/federationmetadata/2007-06/federationmetadata.xml?appid={app_id}"
    for _ in range(12):
        m = re.search(r"<IDPSSODescriptor.*?<X509Certificate>([^<]+)</X509Certificate>",
                      session.get(url, timeout=30).text, re.S)
        if m:
            return m.group(1)
        time.sleep(5)
    raise BuildError(f"no signing cert in SAML metadata for app {app_id}")


def entra_find_policy(g, kind, name):
    return g.one(f"/policies/{kind}?$filter=displayName eq '{odata(name)}'", f"{kind} named '{name}'")


def step_saml_policies(g, dry):
    out = []
    for kind, name, definition, why in SAML_POLICIES:
        pol = entra_find_policy(g, kind, name)
        if pol:
            if json.loads(pol["definition"][0]) != json.loads(definition) and not dry:
                g.req("PATCH", f"/policies/{kind}/{pol['id']}", {"definition": [definition]})
                log("updated", "Entra policy", name, why)
            else:
                log("exists", "Entra policy", name, why)
        elif dry:
            log("would create", "Entra policy", name, why)
            pol = {"id": DRY}
        else:
            pol = g.req("POST", f"/policies/{kind}",
                        {"displayName": name, "definition": [definition], "isOrganizationDefault": False})
            STATE.record("entra", kind, pol["id"], name)
            log("created", "Entra policy", name, why)
        out.append((kind, pol, why))
    return out


def entra_saml_app(g, session, tenant_id, name, sso_url, group, cert_months, dry, policies):
    """Non-gallery SAML app for a Mist portal. Returns the signing cert (base64 DER)."""
    app = entra_find_app(g, name)
    if app:
        log("exists", "SAML enterprise app", name)
        sp = entra_find_sp(g, app["appId"])
    elif dry:
        log("would create", "SAML enterprise app", name, sso_url)
        return None
    else:
        res = g.req("POST", f"/applicationTemplates/{NON_GALLERY_TEMPLATE}/instantiate", {"displayName": name})
        app, sp = res["application"], res["servicePrincipal"]
        STATE.record("entra", "application", app["id"], name, appId=app["appId"])
        log("created", "SAML enterprise app", name)
    if dry:
        return None

    g.req("PATCH", f"/servicePrincipals/{sp['id']}",
          {"preferredSingleSignOnMode": "saml", "appRoleAssignmentRequired": True})
    current = g.req("GET", f"/applications/{app['id']}?$select=identifierUris,web")
    if current.get("identifierUris") != [sso_url]:
        g.req("PATCH", f"/applications/{app['id']}",
              {"identifierUris": [sso_url], "web": {"redirectUris": [sso_url]}})
        log("set", "SAML Entity ID + Reply URL", name, sso_url)

    for kind, policy, why in policies:
        assigned = g.req("GET", f"/servicePrincipals/{sp['id']}/{kind}").get("value", [])
        if policy["id"] in {p["id"] for p in assigned}:
            continue
        for other in assigned:  # an SP takes only one policy of each kind
            g.req("DELETE", f"/servicePrincipals/{sp['id']}/{kind}/{other['id']}/$ref")
        g.req("POST", f"/servicePrincipals/{sp['id']}/{kind}/$ref",
              {"@odata.id": f"{GRAPH}/policies/{kind}/{policy['id']}"})
        log("assigned", "Entra policy", why, f"to {name}")

    sp = g.req("GET", f"/servicePrincipals/{sp['id']}")
    if sp.get("preferredTokenSigningKeyThumbprint"):
        cert = metadata_cert(session, tenant_id, app["appId"])
    else:
        # Entra caps SAML signing certs at 3 years
        end = datetime.now(timezone.utc) + timedelta(days=min(round(cert_months * 30.44), 3 * 365 - 2))
        kc = g.req("POST", f"/servicePrincipals/{sp['id']}/addTokenSigningCertificate",
                   {"displayName": f"CN={name}", "endDateTime": end.strftime("%Y-%m-%dT%H:%M:%SZ")})
        g.req("PATCH", f"/servicePrincipals/{sp['id']}", {"preferredTokenSigningKeyThumbprint": kc["thumbprint"]})
        cert = kc["key"]
        log("created", "SAML signing certificate", name, f"expires {end:%Y-%m-%d}")

    assigned = {a["principalId"] for a in g.req("GET", f"/servicePrincipals/{sp['id']}/appRoleAssignedTo")
                .get("value", [])}
    if group["id"] not in assigned:
        role = next((r["id"] for r in sp.get("appRoles", [])
                     if r.get("isEnabled") and "User" in r.get("allowedMemberTypes", [])), ZERO_GUID)
        try:
            g.req("POST", f"/servicePrincipals/{sp['id']}/appRoleAssignedTo",
                  {"principalId": group["id"], "resourceId": sp["id"], "appRoleId": role}, retry=False)
            log("assigned", "group", group["displayName"], f"only its members can sign in to {name}")
        except BuildError as e:
            g.req("PATCH", f"/servicePrincipals/{sp['id']}", {"appRoleAssignmentRequired": False})
            log("WARNING", "group assignment failed", name,
                "needs Entra ID P1 - ANY tenant user can now sign in. " + str(e)[:100])
    return cert


# --------------------------------------------------------------------------- Mist steps
def mist_ensure(mist, ep, kind, name, body, dry, explain=""):
    obj = mist.find(ep, name)
    if obj:
        log("exists", kind, name)
        return obj
    if dry:
        log("would create", kind, name, explain)
        return {"id": DRY, **body}
    obj = mist.create(ep, body)
    STATE.record("mist", ep, obj["id"], name)
    log("created", kind, name, explain)
    return obj


def step_internal_ca(mist, dry):
    # Must use the dedicated endpoint: mist_scep inside PUT /setting is stored but
    # does NOT provision the CA (the UI keeps showing "Disable").
    scep = mist.req("GET", f"{mist.org}/setting/mist_scep")
    providers = list(scep.get("cert_providers") or [])
    if scep.get("enabled") and "byod" in providers:
        log("exists", "internal CA", "enabled for BYOD certificates")
        return
    if dry:
        log("would enable", "internal CA", "BYOD certificate provider")
        return
    if "byod" not in providers:
        providers.append("byod")
    scep = mist.req("PUT", f"{mist.org}/setting/mist_scep", {"enable": True, "cert_providers": providers})
    if not scep.get("enabled"):
        raise BuildError(f"Mist internal CA did not enable: {scep}")
    log("enabled", "internal CA", "Onboard CA", "issues client certificates to onboarded devices")


def wlan_template(mist, name, site, site_name, dry):
    """Find or create a WLAN template and make sure it applies to the demo site."""
    tmpl = mist.find("templates", name)
    if tmpl:
        log("exists", "WLAN template", name)
        ids = tmpl.get("applies", {}).get("site_ids", [])
        if not dry and site["id"] not in ids:
            tmpl = mist.update("templates", tmpl["id"], {"applies": {"site_ids": ids + [site["id"]]}})
            log("updated", "WLAN template", name, f"applied to {site_name}")
        return tmpl
    if dry:
        log("would create", "WLAN template", name, f"applied to {site_name}")
        return {"id": DRY}
    tmpl = mist.create("templates", {"name": name, "applies": {"site_ids": [site["id"]]}})
    STATE.record("mist", "templates", tmpl["id"], name)
    log("created", "WLAN template", name, f"applied to {site_name}")
    return tmpl


# Best practice for every demo SSID: 2.4/5/6 GHz with high-density data rates.
WLAN_BEST_PRACTICE = {"bands": ["24", "5", "6"],
                      "rateset": {band: {"template": "high-density", "min_rssi": 0} for band in ("24", "5", "6")}}
# Passpoint roaming consortiums the OpenRoaming SSID advertises.
OPENROAMING_OPERATORS = ["openroaming_legacy", "openroaming_settlement_free"]
# Guest / BYOD SSIDs: block client-to-client traffic on the same subnet (not just the same AP).
CLIENT_ISOLATION = {"isolation": True, "l2_isolation": True}


def step_wlans(mist, cfg, n, site, dry):
    """Every demo SSID (PSK, 802.1X and, if enabled, OpenRoaming), each in its own WLAN template
    applied to the demo site."""
    w = cfg["wlan"]
    if w.get("template_name", "").strip():
        log("note", "ini [wlan] template_name", w["template_name"].strip(),
            "no longer used - each SSID now has its own template (psk_template / cert_template)")
    wlans = mist.list("wlans")
    templates = {t["id"]: t["name"] for t in mist.list("templates")}
    pairs = [(n["psk_ssid"], n["psk_template"]), (n["cert_ssid"], n["cert_template"])]
    if n["or_enabled"]:
        pairs.append((n["or_ssid"], n["or_template"]))
    for ssid, tname in pairs:
        elsewhere = [templates.get(x["template_id"], x["template_id"]) for x in wlans
                     if x.get("ssid") == ssid and x.get("template_id") and templates.get(x["template_id"]) != tname]
        if elsewhere:
            raise BuildError(f"SSID {ssid} already exists in WLAN template '{elsewhere[0]}' - run --teardown "
                             f"first (e.g. a build from an older version), or rename the SSID in the ini")

    def wlan(ssid, tname, auth, vlan, extra, explain):
        tmpl = wlan_template(mist, tname, site, site["name"], dry)
        if any(x.get("ssid") == ssid and x.get("template_id") == tmpl["id"] for x in wlans):
            log("exists", "SSID", ssid, f"in {tname}")
            return
        if dry:
            log("would create", "SSID", ssid, explain)
            return
        body = {"ssid": ssid, "enabled": True, "template_id": tmpl["id"], "auth": auth, **extra}
        if vlan:
            body.update({"vlan_enabled": True, "vlan_id": vlan})
        obj = mist.create("wlans", body)
        STATE.record("mist", "wlans", obj["id"], ssid)
        log("created", "SSID", ssid, explain)

    # BYOD Multi-PSK: WPA2, cloud PSKs looked up by Access Assurance (RADIUS), clients isolated.
    # WPA2 only, so 2.4/5 GHz: the Mist UI's PSK portal editor hides RADIUS-PSK WLANs that use WPA3,
    # and 6 GHz needs WPA3. No 802.11r: Fast Roaming is not offered for Multi-PSK. The default PSK
    # is a random fallback nobody is told, so only portal-issued PSKs can join.
    wlan(n["psk_ssid"], n["psk_template"],
         {"type": "psk", "multi_psk_only": False, "pairwise": ["wpa2-ccmp"], "enable_mac_auth": False,
          "private_wlan": False},
         int_or_none(w.get("psk_vlan_id")),
         {**WLAN_BEST_PRACTICE, **CLIENT_ISOLATION, "bands": ["24", "5"], "roam_mode": "NONE", "disable_11be": True,
          "auth_servers": [], "mist_nac": {"enabled": True},
          "dynamic_psk": {"enabled": True, "source": "radius", "default_psk": secrets.token_urlsafe(18),
                          "private_wlan": False}},
         "WPA2 Multi-PSK via Access Assurance (2.4/5 GHz) - every user gets their own passphrase; "
         "high-density rates, clients isolated (same subnet)")
    # Corp 802.1X: WPA3-Enterprise only (no WPA3+WPA2 transition), to match the WPA3 profile the
    # certificate portal installs.
    wlan(n["cert_ssid"], n["cert_template"],
         {"type": "eap", "pairwise": ["wpa3"], "enable_mac_auth": False, "eap_reauth": False,
          "multi_psk_only": False, "private_wlan": False},
         int_or_none(w.get("cert_vlan_id")),
         {**WLAN_BEST_PRACTICE, "roam_mode": "11r", "disable_11be": True, "auth_servers": [], "mist_nac": {"enabled": True}},
         "WPA3-Enterprise 802.1X (no WPA2), authenticated by Mist Access Assurance; high-density rates, 802.11r")
    # OpenRoaming: Passpoint (Hotspot 2.0) 802.1X guest SSID, modelled on a working home-org setup.
    if n["or_enabled"]:
        wlan(n["or_ssid"], n["or_template"],
             {"type": "eap", "pairwise": ["wpa3", "wpa2-ccmp"], "enable_mac_auth": False, "eap_reauth": False,
              "multi_psk_only": False, "private_wlan": False},
             None,
             {**WLAN_BEST_PRACTICE, **CLIENT_ISOLATION, "roam_mode": "11r", "disable_11be": False,
              "auth_servers": [], "mist_nac": {"enabled": True},
              "hotspot20": {"enabled": True, "operators": OPENROAMING_OPERATORS}},
             "Passpoint / OpenRoaming 802.1X, WPA3 + WPA2, 2.4/5/6 GHz, high-density, 802.11r, isolated")


def step_mist_idp(mist, g, cfg, n, tenant_id, realms, idp_app, args):
    dry = args.dry_run
    body = {"name": n["mist_idp"], "idp_type": "oauth", "oauth_type": "azure", "oauth_tenant_id": tenant_id,
            "oauth_cc_client_id": idp_app["appId"], "oauth_ropc_client_id": idp_app["appId"],
            "oauth_ropc_client_secret": ""}
    months = cfg["entra"].getfloat("secret_months", fallback=24)
    sso = mist.find("ssos", n["mist_idp"])
    if sso:
        log("exists", "identity provider", n["mist_idp"])
        if not dry:
            if args.rotate_secret:
                body["oauth_cc_client_secret"] = entra_new_secret(g, idp_app, months)
            sso = mist.update("ssos", sso["id"], body)
            if args.rotate_secret:
                log("updated", "identity provider", n["mist_idp"], "new client secret")
    elif dry:
        log("would create", "identity provider", n["mist_idp"], "OAuth / Azure")
        sso = {"id": DRY}
    else:
        body["oauth_cc_client_secret"] = entra_new_secret(g, idp_app, months)
        sso = mist.create("ssos", body)
        STATE.record("mist", "ssos", sso["id"], n["mist_idp"])
        log("created", "identity provider", n["mist_idp"], f"OAuth / Azure, tenant {tenant_id}")

    setting = mist.req("GET", f"{mist.org}/setting")
    nac = dict(setting.get("mist_nac") or {})
    current = next((i for i in nac.get("idps", []) if i.get("id") == sso["id"]), None)
    make_default = cfg["entra"].getboolean("default_idp", fallback=True)
    change = current is None or sorted(current.get("user_realms", [])) != sorted(realms)
    set_default = make_default and nac.get("default_idp_id") != sso["id"]
    if not (change or set_default):
        log("exists", "realm mapping", ", ".join(realms), "-> Entra ID")
    elif dry:
        log("would set", "realm mapping", ", ".join(realms), "-> Entra ID" + (", default IdP" if set_default else ""))
    else:
        nac["idps"] = [i for i in nac.get("idps", []) if i.get("id") != sso["id"]] + \
                      [{"id": sso["id"], "user_realms": realms}]
        if set_default:
            nac["default_idp_id"] = sso["id"]
        mist.req("PUT", f"{mist.org}/setting", {"mist_nac": nac})
        log("set", "realm mapping", ", ".join(realms),
            "users @ these domains are looked up in Entra" + ("; default IdP" if set_default else ""))
    return sso


def step_openroaming_idp(mist, n, dry):
    """OpenRoaming identity provider for the OpenRoaming SSID, added to Access Assurance with no realms."""
    if not n["or_enabled"]:
        log("note", "OpenRoaming", "off", "[openroaming] enabled = false (or no [openroaming] section)")
        return
    ssid = n["or_ssid"]
    sso = mist_ensure(mist, "ssos", "identity provider", n["or_idp"],
                      {"name": n["or_idp"], "idp_type": "openroaming", "openroaming_ssids": [ssid]},
                      dry, f"OpenRoaming federation for {ssid}")
    if sso["id"] == DRY:
        return
    if sso.get("openroaming_ssids") != [ssid] and not dry:
        mist.update("ssos", sso["id"], {"openroaming_ssids": [ssid]})
        log("updated", "identity provider", n["or_idp"], f"SSID {ssid}")
    nac = dict(mist.req("GET", f"{mist.org}/setting").get("mist_nac") or {})
    if not any(i.get("id") == sso["id"] for i in nac.get("idps", [])):
        if dry:
            log("would add", "Access Assurance IdP", n["or_idp"], "no user realms")
        else:
            nac["idps"] = nac.get("idps", []) + [{"id": sso["id"], "user_realms": []}]
            mist.req("PUT", f"{mist.org}/setting", {"mist_nac": nac})
            log("added", "Access Assurance IdP", n["or_idp"], "no user realms")


def role_tag(mist, tag_name, role, dry):
    """A RADIUS-group NAC label: Access Assurance returns it and the Mist AP uses it as the client's role."""
    return mist_ensure(mist, "nactags", "NAC label", tag_name,
                       {"name": tag_name, "type": "radius_group", "radius_group": role},
                       dry, f"assigns role '{role}'")


def rule_with_role(mist, name, body, role_tag_obj, role, dry, explain):
    """Create the rule with the role label applied, or add the label to an existing rule."""
    rule = mist_ensure(mist, "nacrules", "auth policy rule", name,
                       {**body, "apply_tags": [role_tag_obj["id"]]}, dry, explain)
    if rule["id"] != DRY and role_tag_obj["id"] not in (rule.get("apply_tags") or []):
        if dry:
            log("would set", "auth policy rule", name, f"role '{role}'")
        else:
            rule = mist.update("nacrules", rule["id"], {"apply_tags": (rule.get("apply_tags") or []) + [role_tag_obj["id"]]})
            log("set", "auth policy rule", name, f"role '{role}'")
    return rule


def step_nac(mist, n, dry):
    """Access Assurance policy: every NAC label and rule, each rule returning its role."""
    tag = mist_ensure(mist, "nactags", "NAC label", n["nac_tag"],
                      {"name": n["nac_tag"], "type": "match", "match": "idp_role", "values": [n["group"]]},
                      dry, f"matches Entra group '{n['group']}'")
    role = role_tag(mist, n["corp_role_tag"], n["corp_role"], dry)
    rule_with_role(mist, n["nac_rule"],
                   {"name": n["nac_rule"], "action": "allow", "enabled": True, "order": 1,
                    "matching": {"auth_type": "eap-tls", "nactags": [tag["id"]]}},
                   role, n["corp_role"], dry, f"EAP-TLS certificate + Entra group member -> allow, role '{n['corp_role']}'")
    if n["or_enabled"]:
        ssid = n["or_ssid"]
        tag = mist_ensure(mist, "nactags", "NAC label", n["or_tag"],
                          {"name": n["or_tag"], "type": "match", "match": "ssid", "values": [ssid]},
                          dry, f"matches SSID {ssid}")
        role = role_tag(mist, n["or_role_tag"], n["or_role"], dry)
        rule_with_role(mist, n["or_rule"],
                       {"name": n["or_rule"], "action": "allow", "enabled": True, "order": 2,
                        "matching": {"nactags": [tag["id"]]}},
                       role, n["or_role"], dry, f"SSID {ssid} -> allow, role '{n['or_role']}'")


def step_wxlan(mist, n, dry):
    """Org WxLAN labels + a policy in the PSK SSID's WLAN template: users whose PSK carries the
    isolated role may reach the internet but not the local network."""
    if not n["iso_enabled"]:
        log("note", "isolated BYOD", "off", "[psk_portal_isolated] enabled = false (or no section)")
        return
    users = mist_ensure(mist, "wxtags", "WxLAN label", n["iso_user_tag"],
                        {"name": n["iso_user_tag"], "type": "match", "match": "radius_group", "op": "in",
                         "values": [n["iso_role"]]}, dry, f"PSK role '{n['iso_role']}'")
    internet = mist_ensure(mist, "wxtags", "WxLAN label", n["internet_tag"],
                           {"name": n["internet_tag"], "type": "match", "match": "ip_range_subnet", "op": "in",
                            "values": ["0.0.0.0/0"]}, dry, "0.0.0.0/0")
    local = mist_ensure(mist, "wxtags", "WxLAN label", n["local_tag"],
                        {"name": n["local_tag"], "type": "match", "match": "ip_range_subnet", "op": "in",
                         "values": ["192.168.0.0/16", "172.16.0.0/12", "10.0.0.0/8"]}, dry, "RFC 1918 subnets")
    what = f"{n['iso_user_tag']} -> allow {n['internet_tag']}, deny {n['local_tag']}"
    tmpl = mist.find("templates", n["psk_template"])
    if not tmpl:
        log("would create" if dry else "WARNING", "template policy", what,
            "" if dry else f"WLAN template '{n['psk_template']}' not found")
        return
    rule = next((x for x in mist.list("wxrules") if x.get("template_id") == tmpl["id"]
                 and x.get("src_wxtags") == [users["id"]]), None)
    if rule:
        log("exists", "template policy", what, f"in {n['psk_template']}")
    elif dry:
        log("would create", "template policy", what, f"in {n['psk_template']}")
    else:
        rule = mist.create("wxrules", {"template_id": tmpl["id"], "order": 1, "enabled": True, "action": "allow",
                                       "src_wxtags": [users["id"]], "dst_wxtags": [],
                                       "dst_allow_wxtags": [internet["id"]], "dst_deny_wxtags": [local["id"]]})
        STATE.record("mist", "wxrules", rule["id"], what)
        log("created", "template policy", what, f"in {n['psk_template']}")


def portal_sso(tenant_id, cert_b64):
    return {"idp_sso_url": f"https://login.microsoftonline.com/{tenant_id}/saml2",
            "issuer": f"https://sts.windows.net/{tenant_id}/",
            "idp_cert": pem(cert_b64) if cert_b64 else "",
            "idp_sign_algo": "sha256", "nameid_format": "email"}


def step_psk_portal(mist, g, session, cfg, n, tenant_id, group, policies, dry,
                    name=None, role=None, app=None, vlan_id=None, section="psk_portal"):
    """A BYOD PSK portal on the PSK SSID. Defaults: [psk_portal]; the isolated portal passes its own
    name, role, Entra app, VLAN and ini section. A blank or missing expire_days / max_usage /
    passphrase_length in that section falls back to [psk_portal]."""
    def setting(key, default):
        for sec in (section, "psk_portal"):
            v = cfg.get(sec, key, fallback="").strip() if cfg.has_section(sec) else ""
            if v:
                try:
                    return int(v)
                except ValueError:
                    raise BuildError(f"[{sec}] {key} must be a whole number, not '{v}'")
        return default
    name, role, app = name or n["psk_portal"], role or n["psk_role"], app or n["entra_psk_app"]
    body = {"name": name, "type": "byod", "auth": "sso", "ssid": n["psk_ssid"],
            "expire_time": setting("expire_days", 365) * 1440,
            "max_usage": setting("max_usage", 0),
            # demo: no e-mail is collected or sent
            "required_fields": [], "notify_expiry": False, "notify_on_create_or_edit": False,
            "passphrase_rules": {"length": setting("passphrase_length", 12),
                                 "alphabets_enabled": True, "numerics_enabled": True, "symbols_enabled": False},
            "role": role,
            "sso": {**portal_sso(tenant_id, None), "use_sso_role_for_psk_role": False}}
    vlan = int_or_none(vlan_id or cfg["wlan"].get("psk_vlan_id"))
    if vlan:
        body["vlan_id"] = vlan
    portal = mist_ensure(mist, "pskportals", "PSK portal", name, body, dry,
                         f"BYOD self-service for {n['psk_ssid']}, no e-mail, role '{role}'"
                         + (f", VLAN {vlan}" if vlan else ""))
    sso_url = portal.get("portal_sso_url") or f"https://{mist.cloud}/api/v1/pskportal/{portal['id']}/saml/login"
    cert = entra_saml_app(g, session, tenant_id, app, sso_url, group,
                          cfg["entra"].getfloat("saml_cert_months", fallback=36), dry, policies)
    if cert and not dry:
        body["sso"] = {**portal_sso(tenant_id, cert), "use_sso_role_for_psk_role": False}
        if (portal.get("sso", {}).get("idp_cert") != body["sso"]["idp_cert"]
                or any(portal.get(k) != body[k] for k in ("required_fields", "notify_expiry", "role",
                                                          "expire_time", "max_usage",
                                                          "notify_on_create_or_edit", "ssid", "passphrase_rules"))
                or portal.get("vlan_id") != body.get("vlan_id")):
            portal = mist.update("pskportals", portal["id"], body)
            log("set", "PSK portal", name, f"Entra ID SAML sign-in, role '{role}', passphrase rules")
    return portal


PASSPHRASE_CHARS = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789"  # no 0/O, 1/l/I


def step_sample_psks(mist, cfg, n, dry):
    """One ready-made PSK per PSK role on the PSK SSID, for the report's Wi-Fi QR codes.
    No device limit and no expiry unless [sample_psks] says otherwise. Teardown deletes them
    with every other PSK on the demo SSID."""
    if not n["samples_enabled"]:
        log("note", "sample PSKs", "off", "[sample_psks] enabled = false (or no section)")
        return []
    c = cfg["sample_psks"]
    max_usage = c.getint("max_usage", fallback=0)
    expire_days = c.getint("expire_days", fallback=0)
    roles = [(n["psk_role"], "Corporate BYOD access")]
    if n["iso_enabled"]:
        roles.append((n["iso_role"], "Internet only - no local network"))
    out = []
    for role, access in roles:
        name = f"Sample-{role}"
        body = {"name": name, "ssid": n["psk_ssid"], "role": role, "usage": "multi", "max_usage": max_usage,
                "passphrase": "".join(secrets.choice(PASSPHRASE_CHARS) for _ in range(12))}
        if expire_days:
            body["expire_time"] = int(time.time()) + expire_days * 86400
        limits = (f"{max_usage} devices" if max_usage else "unlimited devices") + \
                 (f", expires in {expire_days} days" if expire_days else ", no expiry")
        psk = mist.find("psks", name)
        if psk:
            log("exists", "sample PSK", name, f"role '{role}'")
        elif dry:
            log("would create", "sample PSK", name, f"role '{role}', {limits}")
            continue
        else:
            psk = mist.create("psks", body)
            log("created", "sample PSK", name, f"role '{role}', {limits}")
        out.append({"name": name, "role": role, "ssid": n["psk_ssid"], "access": access,
                    "passphrase": psk.get("passphrase", ""), "limits": limits})
    return out


def step_cert_portal(mist, g, session, cfg, n, tenant_id, group, policies, dry):
    c = cfg["cert_portal"]
    body = {"name": n["cert_portal"], "type": "marvis_client", "auth": "sso", "ssid": n["cert_ssid"],
            "access_type": c.get("access_type", "wireless").strip(),
            "eap_type": c.get("eap_type", "wpa3").strip().lower(),
            "cert_expire_time": c.getint("cert_expire_days", fallback=365),
            "notify_expiry": False,
            "sso": {**portal_sso(tenant_id, None), "use_sso_role_for_cert": False}}
    if body["eap_type"] not in ("wpa2", "wpa3"):
        raise BuildError("[cert_portal] eap_type must be wpa2 or wpa3")
    portal = mist_ensure(mist, "nacportals", "certificate portal", n["cert_portal"], body, dry,
                         f"Marvis Client onboarding for {n['cert_ssid']}, {body['eap_type'].upper()}")
    sso_url = portal.get("portal_sso_url") or f"https://{mist.cloud}/api/v1/nacportal/{portal['id']}/saml/login"
    cert = entra_saml_app(g, session, tenant_id, n["entra_cert_app"], sso_url, group,
                          cfg["entra"].getfloat("saml_cert_months", fallback=36), dry, policies)
    if cert and not dry:
        body["sso"] = {**portal_sso(tenant_id, cert), "use_sso_role_for_cert": False}
        if (portal.get("sso", {}).get("idp_cert") != body["sso"]["idp_cert"]
                or any(portal.get(k) != body[k] for k in ("ssid", "eap_type"))):
            portal = mist.update("nacportals", portal["id"], body)
            log("set", "certificate portal", n["cert_portal"], f"Entra ID SAML sign-in, {body['eap_type'].upper()} profile")
    return portal


# --------------------------------------------------------------------------- build
RADIUS_WARMUP = (
    "On an org that has never run Access Assurance, the first 802.1X request makes Mist build the "
    "org's cloud RADIUS service, which takes about 10 minutes. Until then, joins fail with "
    "'Radius Server Failed to respond (514) / RadSec Connection Failure (515)' and Access "
    "Assurance logs nothing. It happens once per org: after it works, a teardown and rebuild on the "
    "same org authenticates straight away. Tip: make one join attempt on {ssid} about 10 minutes "
    "before the demo.")


def radius_ready(mist):
    """True if Access Assurance has handled traffic for this org recently (its cloud RADIUS
    service is built), False if not, None if it can't be checked."""
    try:
        r = mist.req("GET", f"{mist.org}/nac_clients/events/search?duration=30d&limit=1")
        return bool(r.get("total") or r.get("results"))
    except (BuildError, requests.RequestException):
        return None


def report_path(cfg, ini_path):
    p = cfg.get("output", "report", fallback="").strip() if cfg.has_section("output") else ""
    if not p:
        p = "Mist_AA_Demo_Report.docx"
    p = os.path.expanduser(p)
    p = p if os.path.isabs(p) else os.path.join(os.path.dirname(os.path.abspath(ini_path)), p)
    return stamped(p)


def build(cfg, mist, g, session, args):
    global RUN
    dry = args.dry_run
    n = names(cfg)
    RUN = Run(11, pause=args.pause)
    started = datetime.now()

    RUN.stage("Connect to Entra ID and Mist",
              "Sign in to both clouds and save a snapshot of the starting state, so the demo can be undone.")
    org = g.req("GET", "/organization")["value"][0]
    tenant_id = org["id"]
    realms = csv_list(cfg["entra"].get("domains")) or [d["name"] for d in org.get("verifiedDomains", [])]
    org_name = mist.req("GET", mist.org).get("name", "")
    site = find_site(mist, cfg, n)
    log("exists", "Entra ID tenant", org.get("displayName", ""), tenant_id)
    log("exists", "Mist organisation", org_name, f"{mist.org_id} on {mist.cloud}")
    warm = radius_ready(mist)
    if warm:
        log("exists", "Access Assurance cloud RADIUS", "ready", "NAC traffic seen in the last 30 days")
    elif warm is False:
        log("note", "Access Assurance cloud RADIUS", "no NAC traffic on this org yet",
            "the first 802.1X request starts a one-off ~10 minute build")
    if STATE.exists:
        log("exists", "snapshot", os.path.basename(STATE.path), "kept from the first build")
    elif dry:
        log("would save", "snapshot", os.path.basename(STATE.path))
    else:
        setting = mist.req("GET", f"{mist.org}/setting")
        STATE.start({
            "taken": datetime.now(timezone.utc).isoformat(), "org_id": mist.org_id, "cloud": mist.cloud,
            "tenant_id": tenant_id, "site": site_snapshot(site),
            "mist_nac": setting.get("mist_nac"), "setting_had_mist_scep": "mist_scep" in setting,
            "mist_scep": mist.req("GET", f"{mist.org}/setting/mist_scep"),
            "names": n,
        })
        log("set", "snapshot", os.path.basename(STATE.path), "used by --teardown to restore everything")

    if args.site:
        RUN.stage("Prepare the demo site",
                  "Rename the site, set its address and upload its floorplan so the demo looks real.")
        site = step_site(mist, cfg, n, site, dry)
        step_floorplan(mist, cfg, n, site, args.ini, dry)
    else:
        RUN.stage("Demo site: left as it is",
                  "Networks only: the site keeps its name, address and floorplan; the SSIDs are applied to it.")
        log("note", "site", site["name"], "networks only - not renamed, no floorplan uploaded")

    RUN.stage("Entra ID: register Mist as an application",
              "Mist Access Assurance needs an Entra app to read users, groups and devices via Microsoft Graph.")
    idp_app = step_idp_app(g, n, dry)

    RUN.stage("Entra ID: onboarding group",
              "Only members of this group may use the portals and get Wi-Fi access.")
    group = step_group(g, cfg, n, dry)

    RUN.stage("Entra ID: SAML settings for Mist portals",
              "Sign the whole SAML response and send the user's UPN as their e-mail-format ID - both required by Mist.")
    policies = step_saml_policies(g, dry)

    RUN.stage("Mist: turn on the internal certificate authority",
              "Mist's onboard CA issues each onboarded device its own client certificate.")
    step_internal_ca(mist, dry)

    RUN.stage("Mist: create the SSIDs",
              f"{n['psk_ssid']} uses a personal passphrase per user; {n['cert_ssid']} uses certificates (802.1X)"
              + (f"; {n['or_ssid']} is a Passpoint SSID for OpenRoaming visitors" if n["or_enabled"] else "")
              + ". Each SSID has its own WLAN template, applied to the demo site.")
    step_wlans(mist, cfg, n, site, dry)

    RUN.stage("Mist: identity providers",
              "Add Entra ID and route the tenant's domains to it"
              + ("; add OpenRoaming so visitors with an OpenRoaming identity are recognised" if n["or_enabled"] else "")
              + ".")
    step_mist_idp(mist, g, cfg, n, tenant_id, realms, idp_app, args)
    step_openroaming_idp(mist, n, dry)

    RUN.stage("Mist: Access Assurance policy",
              f"Allow {n['cert_ssid']} only with a valid certificate AND Entra group membership (role "
              f"'{n['corp_role']}')" + (f"; allow {n['or_ssid']} by SSID (role '{n['or_role']}')" if n["or_enabled"] else "")
              + ". PSK joins don't use these rules.")
    step_nac(mist, n, dry)

    RUN.stage("Mist: WxLAN policy",
              f"Users whose PSK carries the role '{n['iso_role']}' reach the internet but not the local network.")
    step_wxlan(mist, n, dry)

    RUN.stage("Onboarding portals",
              "Users sign in with Entra ID. The PSK portals give each user their own passphrase and QR code; "
              "the certificate portal's Marvis Client app installs a certificate and Wi-Fi profile"
              + ("; sample PSKs for the report's Wi-Fi QR codes" if n["samples_enabled"] else "") + ".")
    psk = step_psk_portal(mist, g, session, cfg, n, tenant_id, group, policies, dry)
    iso = None
    if n["iso_enabled"]:
        iso = step_psk_portal(mist, g, session, cfg, n, tenant_id, group, policies, dry, name=n["iso_portal"],
                              role=n["iso_role"], app=n["entra_iso_app"], vlan_id=n["iso_vlan"],
                              section="psk_portal_isolated")
    cert_portal = step_cert_portal(mist, g, session, cfg, n, tenant_id, group, policies, dry)
    samples = step_sample_psks(mist, cfg, n, dry)
    RUN.end()

    print()
    if dry:
        print(_c("1;33", "Dry run complete - nothing was changed. Re-run without --dry-run to build."))
        return
    # The report shows what the demo actually did across every build run, so a re-run
    # (where most steps report "exists") doesn't hide the original build.
    hist = STATE.data.setdefault("history", [])
    for st in RUN.stages:
        h = next((x for x in hist if x["title"] == st["title"]), None)
        if h is None:
            h = {"title": st["title"], "explain": st["explain"], "events": []}
            hist.append(h)
        for ev in st["events"]:
            if ev[0] != "exists" and list(ev) not in h["events"]:
                h["events"].append(list(ev))
    STATE.data.setdefault("first_build", started.isoformat())
    STATE.save()
    stages = [{"title": h["title"], "explain": h["explain"], "events": [tuple(e) for e in h["events"]]}
              for h in hist]
    out = report_path(cfg, args.ini)
    from demo_report import write_report
    write_report(out, {
        "tenant_name": org.get("displayName", ""), "tenant_id": tenant_id,
        "org_name": org_name, "org_id": mist.org_id, "cloud": mist.cloud,
        "site": {"name": site.get("name"), "address": site.get("address")},
        "psk_ssid": n["psk_ssid"], "cert_ssid": n["cert_ssid"], "group": n["group"],
        "or_ssid": n["or_ssid"] if n["or_enabled"] else None,
        "demo_users": csv_list(cfg["entra"].get("demo_users")),
        "portals": [
            {"title": "Wi-Fi passphrase (PSK)", "url": psk.get("ui_url", ""), "ssid": n["psk_ssid"],
             "how": "Sign in with Entra ID, then scan the QR code shown or type the passphrase."},
        ] + ([
            {"title": "Isolated PSK (internet only)", "url": iso.get("ui_url", ""), "ssid": n["psk_ssid"],
             "how": "As the PSK portal, but devices can reach only the internet, not the local network."},
        ] if iso else []) + [
            {"title": "Certificate onboarding", "url": cert_portal.get("ui_url", ""), "ssid": n["cert_ssid"],
             "how": "Sign in with Entra ID, install the Marvis Client app, then tap Install Network Profile."},
        ],
        "sample_psks": samples,
        "stages": stages, "started": datetime.fromisoformat(STATE.data["first_build"]),
        "radius_note": None if warm else RADIUS_WARMUP.format(ssid=n["cert_ssid"]),
        "finished": datetime.now(),
    })
    print(_c("1;32", "━━ Demo environment ready " + "━" * 40))
    print(f"   PSK portal:          {psk.get('ui_url', '-')}")
    if iso:
        print(f"   Isolated PSK portal: {iso.get('ui_url', '-')}")
    print(f"   Certificate portal:  {cert_portal.get('ui_url', '-')}")
    print(f"   Report with QR codes: {out}")
    print(f"   Audit log:            {AUDIT_LOG}")
    if not warm:
        import textwrap
        print(_c("33", "\n" + textwrap.fill(RADIUS_WARMUP.format(ssid=n["cert_ssid"]), 92,
                                             initial_indent="   i ", subsequent_indent="     ")))
    print(_c("2", "   Undo everything with:  python3 mist_entra_access_assurance.py --teardown"))


# --------------------------------------------------------------------------- teardown
def purge_deleted(g, oid, what):
    """Permanently remove an object from the Entra recycle bin."""
    for _ in range(12):
        r = g.s.delete(f"{GRAPH}/directory/deletedItems/{oid}", headers=g.h, timeout=60)
        if r.ok or r.status_code == 404 and _ > 3:
            return r.ok
        time.sleep(5)
    log("WARNING", "could not purge", what, "still in the Entra recycle bin")
    return False


def teardown(cfg, mist, g, args):
    global RUN
    if not STATE.exists:
        raise BuildError(f"no snapshot {STATE.path} - nothing was built with this ini, or it was already "
                         "torn down. (Snapshots are written by the build.)")
    snap = STATE.data["snapshot"]
    if snap["org_id"] != mist.org_id:
        raise BuildError(f"snapshot is for org {snap['org_id']}, ini points at {mist.org_id}")
    created = STATE.data["created"]
    psk_ssid = snap["names"]["psk_ssid"]
    org_psks = [p for p in mist.list("psks") if p.get("ssid") == psk_ssid] if \
        any(c["kind"] == "wlans" and c["name"] == psk_ssid for c in created) else []

    print(_c("1", "\nTeardown will return the Mist org and Entra tenant to the snapshot taken "
                  f"{snap['taken'][:19].replace('T', ' ')} UTC:"))
    for c in reversed(created):
        print(f"   - delete {c['system']} {c['kind']}: {c['name']}")
    if org_psks:
        print(f"   - delete {len(org_psks)} PSK(s) issued on {psk_ssid} during the demo")
    print(f"   - restore site '{snap['site']['name']}' (name, address, location)")
    print("   - restore Access Assurance identity-provider settings")
    ca_was = snap["mist_scep"].get("enabled")
    print("   - " + ("disable the internal CA (it was off) - removes every certificate it issued"
                     if not ca_was else "keep the internal CA (it was already on); restore its providers"))
    print("   - permanently purge deleted Entra apps/groups from the recycle bin")
    if not args.yes and input("\nType 'teardown' to continue: ").strip() != "teardown":
        print("Cancelled - nothing changed.")
        return

    RUN = Run(4, pause=args.pause)
    RUN.stage("Mist: remove demo configuration", "Portals, policy, SSIDs, WLAN templates, identity provider and floorplan.")
    for p in org_psks:
        mist.delete("psks", p["id"])
        log("deleted", "PSK", p["name"])
    setting = mist.req("GET", f"{mist.org}/setting")
    nac = snap["mist_nac"] or {}
    if setting.get("mist_nac") != nac:
        mist.req("PUT", f"{mist.org}/setting", {"mist_nac": nac})
        log("restored", "Access Assurance IdP settings", "realm mapping / default IdP")
    # Newest first, but NAC labels last: a rule may have been given a label added on a later run.
    for c in sorted(reversed([c for c in created if c["system"] == "mist"]), key=lambda c: c["kind"] == "nactags"):
        if c["kind"] == "maps":
            mist.req("DELETE", f"{mist.base}/sites/{c['site_id']}/maps/{c['id']}", ok404=True)
        else:
            mist.delete(c["kind"], c["id"])
        log("deleted", {"ssos": "identity provider", "templates": "WLAN template", "wlans": "SSID",
                        "nactags": "NAC label", "nacrules": "auth policy rule", "pskportals": "PSK portal",
                        "nacportals": "certificate portal", "maps": "floorplan",
                        "wxtags": "WxLAN label", "wxrules": "template policy"}.get(c["kind"], c["kind"]), c["name"])

    RUN.stage("Mist: restore the internal CA and site", "Put the CA and the site back exactly as they were.")
    if not ca_was:
        # Always DELETE: 'enabled' can already read false (e.g. once the cert portal is
        # gone) while the CA's data and issued certificates are still there.
        mist.req("DELETE", f"{mist.org}/setting/mist_scep")
        log("disabled", "internal CA", "Onboard CA", "was off before the demo - CA and its certificates removed")
    else:
        mist.req("PUT", f"{mist.org}/setting/mist_scep",
                 {"enable": True, "cert_providers": snap["mist_scep"].get("cert_providers", [])})
        log("restored", "internal CA providers", ", ".join(snap["mist_scep"].get("cert_providers", [])))
    if not snap.get("setting_had_mist_scep") and "mist_scep" in mist.req("GET", f"{mist.org}/setting"):
        try:
            mist.req("PUT", f"{mist.org}/setting", {"mist_scep": None})
        except BuildError:
            pass  # leftover marker only; the CA itself is governed by /setting/mist_scep
    s = snap["site"]
    mist.req("PUT", f"{mist.base}/sites/{s['id']}",
             {k: s[k] for k in ("name", "address", "latlng", "country_code", "timezone")})
    log("restored", "site", s["name"], f"{s.get('country_code')}, {s.get('timezone')}")

    RUN.stage("Entra ID: remove demo apps, policies and group", "Deleted and purged from the recycle bin.")
    purge = []
    for c in reversed([c for c in created if c["system"] == "entra"]):
        k = c["kind"]
        if k == "group_member":
            if not any(x["kind"] == "group" and x["id"] == c["group_id"] for x in created):
                g.req("DELETE", f"/groups/{c['group_id']}/members/{c['id']}/$ref", retry=False, ok404=True)
                log("removed", "group member", c["name"])
        elif k == "application":
            g.req("DELETE", f"/applications/{c['id']}", retry=False, ok404=True)
            purge.append(c)
            log("deleted", "app", c["name"])
        elif k == "group":
            g.req("DELETE", f"/groups/{c['id']}", retry=False, ok404=True)
            purge.append(c)
            log("deleted", "security group", c["name"])
        else:
            g.req("DELETE", f"/policies/{k}/{c['id']}", retry=False, ok404=True)
            log("deleted", "Entra policy", c["name"])

    RUN.stage("Entra ID: empty the recycle bin", "So the tenant has no trace of the demo objects.")
    for c in purge:
        if purge_deleted(g, c["id"], c["name"]):
            log("purged", "recycle bin", c["name"])
    RUN.end()

    done = STATE.finish()
    print(_c("1;32", "\n━━ Org and tenant restored to their pre-demo state " + "━" * 15))
    print(_c("2", f"   Snapshot archived as {os.path.basename(done)}"))
    print(_c("2", f"   Audit log: {AUDIT_LOG}"))


# --------------------------------------------------------------------------- main
def build_summary(cfg, n, with_site=None):
    """Short plan of the build, shown before anything connects."""
    s, e = cfg["site"], cfg["entra"]
    plan = (s.get("floorplan", "") or "").strip()
    plan = ("no floorplan" if plan.lower() in ("none", "off", "no")
            else f"floorplan '{plan}'" if plan else "the built-in sample floorplan")
    ssids = [f"{n['psk_ssid']} (Multi-PSK)", f"{n['cert_ssid']} (802.1X)"] + \
            ([f"{n['or_ssid']} (OpenRoaming)"] if n["or_enabled"] else [])
    portals = [n["psk_portal"]] + ([n["iso_portal"]] if n["iso_enabled"] else []) + [n["cert_portal"]]
    users = csv_list(e.get("demo_users"))
    lines = [
        f"Entra ID ({e.get('tenant', '').strip()}): '{n['prefix']} ...' apps, group '{n['group']}'"
        + (f" with {len(users)} demo user{'s' if len(users) != 1 else ''}" if users else "") + ", SAML policies",
        ("Demo site: left as it is (networks only)" if with_site is False else
         ("Demo site (your choice below): " if with_site is None else "Demo site: ")
         + f"rename '{s.get('current_name', '').strip()}' to '{n['site']}', set its address and upload {plan}"),
        f"SSIDs: {', '.join(ssids)}",
        "Identity providers: " + n["mist_idp"] + (f", {n['or_idp']}" if n["or_enabled"] else ""),
        f"Access Assurance roles: {n['corp_role']}" + (f", {n['or_role']}" if n["or_enabled"] else "")
        + (f"; WxLAN policy: {n['iso_role']} = internet only" if n["iso_enabled"] else ""),
        f"Portals: {', '.join(portals)}",
    ]
    print(_c("1", "\nAbout to build:"))
    for line in lines:
        print(f"   • {line}")


def choose_scope(args, cfg, n):
    """Site + floorplan + networks, or networks only. Asks unless --with-site / --networks-only."""
    if args.with_site or args.networks_only:
        build_summary(cfg, n, args.with_site)
        choice = "networks only" if args.networks_only else "site, floorplan and networks"
        print(f"\n   Building: {choice} (from the command line)\n")
        return args.with_site
    build_summary(cfg, n)
    if not sys.stdin.isatty():
        raise BuildError("no terminal to ask on - pass --with-site or --networks-only")
    print("\n   1  Site, floorplan and networks  (rename the site and upload the plan)")
    print("   2  Networks only  (leave the site's name, address and plan as they are)")
    while True:
        a = input("\nWhat should I build? [1/2, Enter = 1, q = quit]: ").strip().lower()
        if a in ("", "1"):
            return True
        if a == "2":
            return False
        if a in ("q", "quit"):
            raise KeyboardInterrupt


def main():
    global STATE, AUDIT_LOG
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--ini", default=DEFAULT_INI)
    ap.add_argument("--dry-run", action="store_true", help="read only; show what would be built")
    ap.add_argument("--pause", action="store_true", help="wait for Enter between steps (live demo)")
    ap.add_argument("--rotate-secret", action="store_true",
                    help="new Entra client secret pushed to the existing Mist identity provider")
    ap.add_argument("--teardown", action="store_true", help="undo the demo using the saved snapshot")
    ap.add_argument("--yes", action="store_true", help="teardown without the confirmation prompt")
    scope = ap.add_mutually_exclusive_group()
    scope.add_argument("--with-site", action="store_true",
                       help="build the demo site and floorplan too, without asking")
    scope.add_argument("--networks-only", action="store_true",
                       help="build only the networks (SSIDs, policy, portals); leave the site as it is")
    args = ap.parse_args()
    try:
        cfg = load_config(args.ini)
        n = names(cfg)  # fail fast on over-long Mist names
        AUDIT_LOG = start_audit_log(cfg, args.ini, "teardown" if args.teardown else
                                    "dry-run" if args.dry_run else "build", sys.argv)
        print(_c("2", f"Audit log: {AUDIT_LOG}"))
        args.site = False if args.teardown else choose_scope(args, cfg, n)
        STATE = State(os.path.splitext(os.path.abspath(args.ini))[0] + ".state.json", dry=args.dry_run)
        session = build_session(cfg)
        cloud, token = mist_credentials(cfg)
        mist = Mist(session, cloud, token, cfg["mist"].get("org_id", "").strip() or None)
        tenant = cfg["entra"].get("tenant", "").strip()
        if not tenant:
            raise BuildError("[entra] tenant is required (e.g. contoso.onmicrosoft.com or the tenant ID)")
        g = Graph(session, tenant)
        if args.teardown:
            teardown(cfg, mist, g, args)
        else:
            build(cfg, mist, g, session, args)
    except BuildError as e:
        RUN.end() if RUN.cur else None
        print(_c("1;31", f"\nERROR: {e}"), file=sys.stderr)
        print("Fix it and re-run: finished steps are detected and skipped, and --teardown still "
              "undoes whatever was built.", file=sys.stderr)
        sys.exit(1)
    except requests.exceptions.SSLError:
        print(_c("1;31", "\nERROR: SSL error - set [network] ca_bundle (Zscaler/proxy)"), file=sys.stderr)
        sys.exit(1)
    except KeyboardInterrupt:
        sys.exit(130)


if __name__ == "__main__":
    main()
