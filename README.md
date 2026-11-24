# Mist Entra Access Assurance demo

This is a live-demo builder for **Juniper Mist Access Assurance with Microsoft Entra ID**. One command sets up Entra ID single sign-on and Wi-Fi onboarding for both PSK and certificate users. A second command puts the Mist org and the Entra tenant back exactly as they were.

## What the audience sees

The build prints one banner per step, with a one-line explanation and a ✓ for each change:

| Step | What happens |
|---|---|
| 1. Connect | Signs in to Entra ID and Mist, and saves a snapshot of the starting state |
| 2. Demo site | Renames the site, sets its address, country, timezone and location from the ini, and uploads its floorplan |
| 3. Entra app | **IdP Connector** app registration with Graph `User.Read` (delegated) and `User.Read.All`, `GroupMember.Read.All`, `Device.Read.All` (application). Grants admin consent and turns on public client flows (ROPC / EAP-TTLS) |
| 4. Entra group | Onboarding security group, with the demo users added |
| 5. Mist IdP | OAuth/Azure identity provider using a fresh client secret (never shown), plus realm mapping and the default IdP |
| 6. Internal CA | Turns on the Mist onboard CA for BYOD certificates |
| 7. SSIDs | A Multi-PSK SSID and an 802.1X SSID on Mist NAC, each in its own WLAN template. Both templates are applied to the demo site |
| 8. Auth policy | NAC label (Entra group) and the rule *EAP-TLS + group → allow* |
| 9. OpenRoaming | A Passpoint (Hotspot 2.0) 802.1X SSID in its own WLAN template, an OpenRoaming identity provider for it, and the rule *SSID → allow* (an SSID-match NAC label). Turn off with `[openroaming] enabled = false` |
| 10. Entra SAML | Policies that sign the whole SAML response and send the UPN as the NameID in email format. Mist requires both |
| 11. PSK portal | BYOD self-service portal with Entra SSO and no email collection or notifications, plus its SAML app |
| 12. Cert portal | Marvis Client certificate onboarding portal with Entra SSO, plus its SAML app |

When the build finishes, it writes a **Word report** with both portal QR codes and every action taken.

Every run (build, dry-run or teardown) also writes an **audit log**: the same output as the screen, without colours and with a time on each line. Both file names carry the run's timestamp, so each run keeps its own copy:

- `Mist_Access_Assurance_Demo_Report_20261119-143005.docx`
- `Mist_Access_Assurance_Audit_20261119-143005_build.log`

Set the base names under `[output]` in the ini.

## Setup

```bash
pip3 install -r requirements.txt
cp mist_entra_access_assurance.ini.example mist_entra_access_assurance.ini
```

In the ini, set:
- the Mist `cloud`, `token` and (optionally) `org_id`
- the Entra `tenant` and `demo_users`
- the SSID names
- the site's `current_name`, plus the new name and address
- the site's floorplan image. Leave `floorplan` blank to use the included AI-generated `sample_office_floorplan.png` (a 40 × 25 m office at 50 px = 1 m, with a 24-seat meeting room and 10 two-desk offices). To use your own plan, set `floorplan` to its path and `floorplan_ppm` to its scale in pixels per metre. Set `floorplan = none` for no floorplan

Mist object names must be **32 characters or fewer** (Mist UI limit). The script checks this before it starts.

## Run the demo

```bash
python3 mist_entra_access_assurance.py --dry-run
python3 mist_entra_access_assurance.py --pause
python3 mist_entra_access_assurance.py --teardown
```

- `--dry-run` reads only and shows every step it would take.
- `--pause` waits for Enter between steps, so you can talk the audience through them.
- `--teardown` restores everything from the snapshot. Add `--yes` to skip the prompt.

Entra sign-in uses a device code. Sign in as a Global Administrator. The token is cached in `.msal_token_cache.json` (chmod 600), so later runs usually don't need it. Re-running the build is safe: anything that already exists is skipped, and the report keeps the full history.

## Teardown restores

The snapshot, `<ini>.state.json`, is written before the first change. Teardown then does the following:

- Deletes every Mist object the build created, plus any PSKs issued on the demo PSK SSID.
- Deletes the uploaded floorplan.
- Restores the org's Access Assurance IdP settings and the site's original name, address, country, timezone and location.
- Turns the internal CA back off if it was off before. This removes every certificate it issued. If the CA was already on, teardown keeps it and restores its providers.
- Deletes the Entra app registrations, SAML apps, policies and group (if the build created it), and purges them from the Entra recycle bin. Demo users added to a group that already existed are removed from it.

Objects the build didn't create are never touched. The snapshot is archived as `<ini>.state.restored-<time>.json`.

## Demo tips

- **On an org that has never run Access Assurance, allow 10 minutes after the first join attempt.** The first 802.1X request makes Mist build the org's cloud RADIUS service, which takes about 10 minutes. Until then, joins to the 802.1X SSID fail with "Radius Server Failed to respond (514) / RadSec Connection Failure (515)" and Access Assurance logs nothing. This happens once per org: after it works, a teardown and rebuild on the same org authenticates straight away. Step 1 of the build tells you if the org has no Access Assurance traffic yet. Before the demo, make one join attempt on the 802.1X SSID to start the build.
- If a phone or browser was used in an earlier demo, open the portal links in a **new private tab**. The portals reuse a cached sign-in that belongs to the previous (deleted) portal, which makes the page blank or makes the Marvis app fail with "failed to get user profile data".
- The PSK portal fills the PSK's email field from the Entra sign-in, but no email is sent: notifications are off and no email is asked for. The portal's "Email Passphrase" button only sends if someone presses it.
- EAP-TTLS (username/password) uses ROPC, which skips MFA. Exclude the Mist Access Assurance source IPs from Conditional Access MFA, or use certificates only.
- Group assignment to the SAML apps needs Entra ID P1/P2. Without it, the script warns you and lets any tenant user sign in.
- The site needs an AP for the SSIDs to broadcast.
