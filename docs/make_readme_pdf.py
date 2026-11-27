"""Build docs/README.pdf: the README plus simulated screenshots of the script.

The screenshots are drawn from example data (Contoso tenant, example IDs) in the same
colours and layout the script prints, so no real tenant, org or passphrase appears.

    python3 docs/make_readme_pdf.py
"""
import io
import os
import re

import qrcode
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (Image as RLImage, KeepTogether, ListFlowable, ListItem, Paragraph,
                                Preformatted, SimpleDocTemplate, Spacer, Table, TableStyle)

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SHOTS = os.path.join(HERE, "screenshots")
MONO = "/System/Library/Fonts/Menlo.ttc"
SANS = "/System/Library/Fonts/Supplemental/Arial.ttf"
SANS_B = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
SANS_I = "/System/Library/Fonts/Supplemental/Arial Italic.ttf"
UNI = "/Library/Fonts/Arial Unicode.ttf"
ACCENT = colors.HexColor("#1F4E79")

# ----------------------------------------------------------------- terminal screenshots
TERM = {"bg": (30, 32, 36), "fg": (220, 220, 220), "dim": (130, 135, 140), "bold": (255, 255, 255),
        "cyan": (80, 200, 220), "green": (110, 200, 110), "yellow": (230, 190, 80), "red": (235, 95, 95)}


def banner(n, total, title):
    return [("cyan", "━━ "), ("cyan+bold", f"Step {n}/{total}  {title} "), ("cyan", "━" * max(3, 60 - len(title)))]


def explain(text):
    return [("dim", f"   {text}")]


def ev(status, kind, name, extra=""):
    sym, col = {"exists": ("•", "dim"), "note": ("i", "yellow")}.get(
        status, ("○", "yellow") if status.startswith("would") else ("✓", "green"))
    line = [("", "   "), (col, sym), ("", " "), (col, status.ljust(9)), ("", f" {kind}: "), ("bold", name)]
    return line + ([("dim", f"  ({extra})")] if extra else [])


def text(t, style=""):
    return [(style, t)]


def render_terminal(lines, path, title, cols=104):
    font = ImageFont.truetype(MONO, 26)
    bold = ImageFont.truetype(MONO, 26, index=1)
    cw = font.getlength("M")
    lh = 36
    w, h = int(cols * cw) + 60, len(lines) * lh + 110
    img = Image.new("RGB", (w, h), TERM["bg"])
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, w, 52], fill=(52, 54, 60))
    for i, c in enumerate([(237, 106, 94), (245, 191, 79), (98, 197, 84)]):
        d.ellipse([24 + i * 36, 16, 44 + i * 36, 36], fill=c)
    tf = ImageFont.truetype(SANS, 22)
    d.text((w / 2 - tf.getlength(title) / 2, 14), title, fill=(200, 200, 200), font=tf)
    y = 76
    for segs in lines:
        x = 30
        for style, t in segs:
            parts = style.split("+")
            col = TERM.get(parts[0], TERM["fg"]) if parts[0] else TERM["fg"]
            if parts[0] == "bold":
                col = TERM["bold"]
            f = bold if "bold" in parts else font
            for ch in t:  # bold Menlo lacks box-drawing glyphs: draw those in the regular face
                d.text((x, y), ch, fill=col, font=font if ch in "━•✓○" else f)
                x += cw
        y += lh
    img.save(path)
    return path


def shot_summary():
    L = [text("$ python3 mist_entra_access_assurance.py --pause", "bold"),
         text("Audit log: /demo/Mist_AA_Audit_20261127-163249_build.log", "dim"),
         text(""),
         text("About to build:", "bold"),
         text("   • Entra ID (contoso.onmicrosoft.com): 'Mist AA ...' apps, group 'Mist WiFi Users' with 3 demo"),
         text("     users, SAML policies"),
         text("   • Demo site (your choice below): rename 'Primary Site' to 'AA Demo', set its address and upload"),
         text("     the built-in sample floorplan"),
         text("   • SSIDs: Corp-PSK (Multi-PSK), Corp-Secure (802.1X), OpenRoaming (OpenRoaming)"),
         text("   • Identity providers: Entra ID, OpenRoaming"),
         text("   • Access Assurance roles: Corp, Guest; WxLAN policy: BYoD_Isolated = internet only"),
         text("   • Portals: PSK Self-Service, BYoD_Isolated, Certificate Onboarding"),
         text(""),
         text("   1  Site, floorplan and networks  (rename the site and upload the plan)"),
         text("   2  Networks only  (leave the site's name, address and plan as they are)"),
         text(""),
         [("", "What should I build? [1/2, Enter = 1, q = quit]: "), ("bold", "1")]]
    return render_terminal(L, os.path.join(SHOTS, "01_summary_prompt.png"), "Start-up summary and choice")


def shot_build():
    L = [banner(7, 11, "Mist: create the SSIDs"),
         explain("Corp-PSK uses a personal passphrase per user; Corp-Secure uses certificates (802.1X); ..."),
         ev("created", "WLAN template", "AA - Corp-PSK", "applied to AA Demo"),
         ev("created", "SSID", "Corp-PSK", "WPA2 Multi-PSK via Access Assurance (2.4/5 GHz)"),
         ev("created", "WLAN template", "AA - Corp-Secure", "applied to AA Demo"),
         ev("created", "SSID", "Corp-Secure", "WPA3-Enterprise 802.1X (no WPA2), 802.11r"),
         ev("created", "WLAN template", "AA - OpenRoaming", "applied to AA Demo"),
         ev("created", "SSID", "OpenRoaming", "Passpoint / OpenRoaming, WPA3 + WPA2, isolated"),
         text("   done in 1.2s", "dim"),
         text(""),
         banner(8, 11, "Mist: identity providers"),
         explain("Add Entra ID and route the tenant's domains to it; add OpenRoaming ..."),
         ev("created", "identity provider", "Entra ID", "OAuth / Azure"),
         ev("set", "realm mapping", "contoso.com", "users @ these domains are looked up in Entra; default IdP"),
         ev("created", "identity provider", "OpenRoaming", "OpenRoaming federation for OpenRoaming"),
         text("   done in 1.9s", "dim"),
         text(""),
         banner(9, 11, "Mist: Access Assurance policy"),
         explain("Allow Corp-Secure only with a valid certificate AND Entra group membership (role 'Corp'); ..."),
         ev("created", "NAC label", "Entra grp: Mist WiFi Users", "matches Entra group 'Mist WiFi Users'"),
         ev("created", "NAC label", "Role: Corp", "assigns role 'Corp'"),
         ev("created", "auth policy rule", "EAP-TLS Entra users", "EAP-TLS certificate + Entra group member -> allow"),
         ev("created", "NAC label", "SSID: OpenRoaming", "matches SSID OpenRoaming"),
         ev("created", "NAC label", "Role: Guest", "assigns role 'Guest'"),
         ev("created", "auth policy rule", "OpenRoaming SSID", "SSID OpenRoaming -> allow, role 'Guest'"),
         text("   done in 1.2s", "dim")]
    return render_terminal(L, os.path.join(SHOTS, "02_build_steps.png"), "Build steps 7-9")


def shot_finish():
    L = [banner(11, 11, "Onboarding portals"),
         ev("created", "PSK portal", "PSK Self-Service", "BYOD self-service for Corp-PSK, role 'BYOD'"),
         ev("created", "SAML enterprise app", "Mist AA PSK Portal"),
         ev("assigned", "group", "Mist WiFi Users", "only its members can sign in to Mist AA PSK Portal"),
         ev("created", "PSK portal", "BYoD_Isolated", "BYOD self-service for Corp-PSK, role 'BYoD_Isolated'"),
         ev("created", "certificate portal", "Certificate Onboarding", "Marvis Client onboarding, WPA3"),
         ev("created", "sample PSK", "Sample-BYOD", "role 'BYOD', unlimited devices, no expiry"),
         ev("created", "sample PSK", "Sample-BYoD_Isolated", "role 'BYoD_Isolated', unlimited devices"),
         text("   done in 39.8s", "dim"),
         text(""),
         [("green", "━━ "), ("green+bold", "Demo environment ready "), ("green", "━" * 40)],
         text("   PSK portal:          https://pskportal.gc1.mist.com/#!byod/11111111-2222-3333-4444-555555555555"),
         text("   Isolated PSK portal: https://pskportal.gc1.mist.com/#!byod/66666666-7777-8888-9999-000000000000"),
         text("   Certificate portal:  https://nacportal.gc1.mist.com/#/portal/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"),
         text("   Report with QR codes: /demo/Mist_AA_Demo_Report_20261127-163249.docx"),
         text("   Audit log:            /demo/Mist_AA_Audit_20261127-163249_build.log"),
         text("   Undo everything with:  python3 mist_entra_access_assurance.py --teardown", "dim")]
    return render_terminal(L, os.path.join(SHOTS, "03_build_finished.png"), "Build finished")


def shot_teardown():
    L = [text("$ python3 mist_entra_access_assurance.py --teardown --yes", "bold"),
         text(""),
         banner(1, 4, "Mist: remove demo configuration"),
         ev("deleted", "PSK", "Sample-BYOD"),
         ev("deleted", "PSK", "Sample-BYoD_Isolated"),
         ev("restored", "Access Assurance IdP settings", "realm mapping / default IdP"),
         ev("deleted", "PSK portal", "BYoD_Isolated"),
         ev("deleted", "template policy", "Isolated BYOD Users -> allow Internet, deny Local Network"),
         ev("deleted", "SSID", "Corp-Secure"),
         ev("deleted", "WLAN template", "AA - Corp-Secure"),
         text("   ...", "dim"),
         banner(2, 4, "Mist: restore the internal CA and site"),
         ev("disabled", "internal CA", "Onboard CA", "was off before the demo - CA and its certificates removed"),
         ev("restored", "site", "Primary Site", "GB, Europe/London"),
         banner(3, 4, "Entra ID: remove demo apps, policies and group"),
         ev("deleted", "app", "Mist AA PSK Portal"),
         ev("deleted", "security group", "Mist WiFi Users"),
         banner(4, 4, "Entra ID: empty the recycle bin"),
         ev("purged", "recycle bin", "Mist AA PSK Portal"),
         text(""),
         [("green", "━━ "), ("green+bold", "Org and tenant restored to their pre-demo state "), ("green", "━" * 15)]]
    return render_terminal(L, os.path.join(SHOTS, "04_teardown.png"), "Teardown")


def _qr(payload, size):
    q = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=10, border=2)
    q.add_data(payload)
    q.make(fit=True)
    return q.make_image(fill_color="black", back_color="white").convert("RGB").resize((size, size))


def shot_report_psk_page():
    """Mock of the report's sample-PSK page with example (non-working) passphrases."""
    w, h = 1240, 900
    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)
    hb, b, r, m = (ImageFont.truetype(SANS_B, s) for s in (40,)), ImageFont.truetype(SANS_B, 28), \
        ImageFont.truetype(SANS, 24), ImageFont.truetype(MONO, 28)
    hb = next(hb)
    d.text((70, 50), "Join the Wi-Fi with a sample passphrase", fill=(31, 78, 121), font=hb)
    d.text((70, 112), "Point a phone camera at a code and tap the prompt to join - no portal or sign-in needed.",
           fill=(40, 40, 40), font=r)
    for i, (role, access, pw) in enumerate((("BYOD", "Corporate BYOD access", "ExamplePass1"),
                                            ("BYoD_Isolated", "Internet only - no local network", "ExamplePass2"))):
        x0 = 70 + i * 560
        d.rectangle([x0, 170, x0 + 540, 850], outline=(190, 190, 190), width=2)
        cx = x0 + 270
        for y, t, f, c in ((195, f"Role: {role}", b, (31, 78, 121)), (240, access, r, (40, 40, 40))):
            d.text((cx - d.textlength(t, font=f) / 2, y), t, fill=c, font=f)
        img.paste(_qr(f"WIFI:T:WPA;S:Corp-PSK;P:{pw};;", 380), (cx - 190, 290))
        for y, (lab, val, f) in zip((690, 730, 772), (("SSID: ", "Corp-PSK", r), ("Passphrase: ", pw, m),
                                                    ("PSK name: ", f"Sample-{role}", r))):
            tw = d.textlength(lab, font=b) + d.textlength(val, font=f)
            d.text((cx - tw / 2, y), lab, fill=(20, 20, 20), font=b)
            d.text((cx - tw / 2 + d.textlength(lab, font=b), y), val, fill=(20, 20, 20), font=f)
    path = os.path.join(SHOTS, "05_report_sample_psks.png")
    img.save(path)
    return path


# ----------------------------------------------------------------- README -> PDF
def setup_fonts():
    pdfmetrics.registerFont(TTFont("Body", SANS))
    pdfmetrics.registerFont(TTFont("Body-Bold", SANS_B))
    pdfmetrics.registerFont(TTFont("Body-Italic", SANS_I))
    pdfmetrics.registerFont(TTFont("Uni", UNI))
    pdfmetrics.registerFont(TTFont("Mono", MONO, subfontIndex=0))
    from reportlab.pdfbase.pdfmetrics import registerFontFamily
    registerFontFamily("Body", normal="Body", bold="Body-Bold", italic="Body-Italic", boldItalic="Body-Bold")


def inline(t):
    t = t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    t = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r"\1", t)  # [text](link) -> text
    t = re.sub(r"`([^`]+)`", r'<font face="Mono" size="8.5" color="#8B1A1A">\1</font>', t)
    t = re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", t)
    t = re.sub(r"(?<![\w*])\*([^*]+)\*(?![\w*])", r"<i>\1</i>", t)
    for ch in "✓→×":
        t = t.replace(ch, f'<font face="Uni">{ch}</font>')
    return t


STY = {
    "h1": ParagraphStyle("h1", fontName="Body-Bold", fontSize=20, leading=24, textColor=ACCENT, spaceAfter=8),
    "h2": ParagraphStyle("h2", fontName="Body-Bold", fontSize=14, leading=18, textColor=ACCENT, spaceBefore=12,
                         spaceAfter=6),
    "p": ParagraphStyle("p", fontName="Body", fontSize=10, leading=14, spaceAfter=6),
    "cell": ParagraphStyle("cell", fontName="Body", fontSize=8.5, leading=11),
    "cellh": ParagraphStyle("cellh", fontName="Body-Bold", fontSize=9, leading=11, textColor=colors.white),
    "cap": ParagraphStyle("cap", fontName="Body-Italic", fontSize=8.5, leading=11, alignment=TA_CENTER,
                          textColor=colors.HexColor("#555555"), spaceAfter=10),
    "code": ParagraphStyle("code", fontName="Mono", fontSize=8.5, leading=11, backColor=colors.HexColor("#F3F3F3"),
                           borderPadding=6, spaceBefore=4, spaceAfter=10),
}


def figure(path, caption, width=170 * mm):
    iw, ih = Image.open(path).size
    return KeepTogether([Spacer(1, 4), RLImage(path, width=width, height=width * ih / iw),
                         Paragraph(inline(caption), STY["cap"])])


def md_to_flowables(md, figures):
    out, lines, i = [], md.splitlines(), 0
    while i < len(lines):
        ln = lines[i]
        if ln.startswith("```"):
            j = i + 1
            while not lines[j].startswith("```"):
                j += 1
            out.append(Preformatted("\n".join(lines[i + 1:j]), STY["code"]))
            i = j + 1
            continue
        if ln.startswith("# "):
            out.append(Paragraph(inline(ln[2:]), STY["h1"]))
        elif ln.startswith("## "):
            out.append(Paragraph(inline(ln[3:]), STY["h2"]))
        elif ln.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r"-+", c) for c in cells):
                    rows.append(cells)
                i += 1
            data = [[Paragraph(inline(c), STY["cellh" if r == 0 else "cell"]) for c in row]
                    for r, row in enumerate(rows)]
            t = Table(data, colWidths=[38 * mm, 132 * mm], repeatRows=1)
            t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), ACCENT),
                                   ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BBBBBB")),
                                   ("VALIGN", (0, 0), (-1, -1), "TOP"),
                                   ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F8FB")])]))
            out += [t, Spacer(1, 8)]
            continue
        elif ln.startswith("- "):
            items = []
            while i < len(lines) and (lines[i].startswith("- ") or lines[i].startswith("  ")):
                if lines[i].startswith("- "):
                    items.append(lines[i][2:])
                else:
                    items[-1] += " " + lines[i].strip()
                i += 1
            out.append(ListFlowable([ListItem(Paragraph(inline(x), STY["p"]), leftIndent=12) for x in items],
                                    bulletType="bullet", start="•", leftIndent=12))
            continue
        elif ln.strip():
            out.append(Paragraph(inline(ln), STY["p"]))
        for key, figs in figures.items():
            if ln.startswith(key):
                out += figs
        i += 1
    return out


def main():
    os.makedirs(SHOTS, exist_ok=True)
    s1, s2, s3, s4, s5 = shot_summary(), shot_build(), shot_finish(), shot_teardown(), shot_report_psk_page()
    setup_fonts()
    figures = {
        "The build prints one banner per step": [
            figure(s1, "Before it connects, the build summarises what it will create and asks what to build."),
            figure(s2, "Each step prints a banner, a one-line explanation and a ✓ for every change.")],
        "When the build finishes": [
            figure(s3, "The finish banner lists the portal links, the report and the audit log."),
            figure(s5, "Report page with one Wi-Fi QR code per sample PSK role (example passphrases).",
                   width=140 * mm)],
        "Objects the build didn't create are never touched": [
            figure(s4, "Teardown removes what the build created and restores the org and tenant.")],
    }
    md = open(os.path.join(ROOT, "README.md"), encoding="utf-8").read()
    pdf = os.path.join(HERE, "README.pdf")
    doc = SimpleDocTemplate(pdf, pagesize=A4, leftMargin=20 * mm, rightMargin=20 * mm, topMargin=18 * mm,
                            bottomMargin=18 * mm, title="Mist Entra Access Assurance demo",
                            author="mist-entra-access-assurance")

    def footer(c, d):
        c.setFont("Body", 8)
        c.setFillColor(colors.HexColor("#888888"))
        c.drawRightString(A4[0] - 20 * mm, 10 * mm, f"Page {d.page}")
        c.drawString(20 * mm, 10 * mm, "Mist Entra Access Assurance demo - screenshots use example data")

    doc.build(md_to_flowables(md, figures), onFirstPage=footer, onLaterPages=footer)
    print(pdf)


if __name__ == "__main__":
    main()
