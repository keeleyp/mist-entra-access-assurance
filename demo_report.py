"""Word (.docx) report of a Mist Access Assurance + Entra ID demo build."""
import io
from datetime import datetime

import qrcode
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

ACCENT = "1F4E79"
FONT = "Calibri"


def _qr_png(url):
    q = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=10, border=2)
    q.add_data(url)
    q.make(fit=True)
    buf = io.BytesIO()
    q.make_image(fill_color="black", back_color="white").save(buf, format="PNG")
    buf.seek(0)
    return buf


def _shade(cell, fill):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcpr.append(shd)


def _fixed_width(t, widths_mm):
    """Fixed layout + explicit grid so every renderer honours the column widths."""
    t.autofit = False
    tblpr = t._tbl.tblPr
    tblw = tblpr.find(qn("w:tblW"))
    if tblw is None:
        tblw = OxmlElement("w:tblW")
        tblpr.append(tblw)
    tblw.set(qn("w:type"), "dxa")
    tblw.set(qn("w:w"), str(int(sum(widths_mm) * 56.7)))
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblpr.append(layout)
    for col, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), widths_mm):
        col.set(qn("w:w"), str(int(w * 56.7)))


def _table(doc, headers, rows, widths_mm):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = ""
        r = c.paragraphs[0].add_run(h)
        r.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        _shade(c, ACCENT)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = str(v)
    _fixed_width(t, widths_mm)
    for row in t.rows:
        for i, w in enumerate(widths_mm):
            row.cells[i].width = Mm(w)
            for p in row.cells[i].paragraphs:
                p.paragraph_format.space_after = Pt(0)
                for r in p.runs:
                    r.font.size = Pt(9)
    doc.add_paragraph()
    return t


def _setup(doc):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Mm(18))
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(10.5)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    for name, size in (("Title", 24), ("Heading 1", 15), ("Heading 2", 12)):
        st = doc.styles[name]
        st.font.name = FONT
        st.font.size = Pt(size)
        st.font.color.rgb = RGBColor.from_string(ACCENT)
        st.paragraph_format.space_before = Pt(12 if name != "Title" else 0)
        st.paragraph_format.space_after = Pt(6)
        st.paragraph_format.keep_with_next = True


def write_report(path, ctx):
    """ctx keys: tenant_name, tenant_id, org_name, org_id, cloud, site{name,address},
    psk_ssid, cert_ssid, group, demo_users, portals[{title,url,ssid,how}],
    stages[{title,explain,secs,events[(status,kind,name,extra)]}], started, finished,
    radius_note (optional)"""
    doc = Document()
    _setup(doc)

    doc.add_paragraph("Mist Access Assurance with Microsoft Entra ID", style="Title")
    p = doc.add_paragraph()
    p.add_run(f"Demo build report  ·  {ctx['finished']:%d %B %Y, %H:%M}").italic = True

    doc.add_paragraph(
        "This document records exactly what the build script configured in Microsoft Entra ID "
        "and in Juniper Mist for this demo: an Entra ID identity provider for Mist Access "
        "Assurance, a self-service PSK portal and a certificate onboarding portal, both signed "
        "in with Entra ID. Secrets are never written to this report.")

    doc.add_heading("Environment", level=1)
    _table(doc, ["Item", "Value"], [
        ("Entra ID tenant", f"{ctx['tenant_name']}  ({ctx['tenant_id']})"),
        ("Mist organisation", f"{ctx['org_name']}  ({ctx['org_id']})"),
        ("Mist cloud", ctx["cloud"]),
        ("Demo site", f"{ctx['site']['name']} — {ctx['site'].get('address') or ''}"),
        ("SSIDs", f"{ctx['psk_ssid']} (Multi-PSK)   ·   {ctx['cert_ssid']} (802.1X EAP-TLS)"
                  + (f"   ·   {ctx['or_ssid']} (OpenRoaming / Passpoint)" if ctx.get("or_ssid") else "")),
        ("Onboarding group", f"{ctx['group']}  (members: {', '.join(ctx['demo_users']) or 'none added'})"),
        ("Build time", f"{ctx['started']:%H:%M:%S} – {ctx['finished']:%H:%M:%S}"),
    ], [45, 129])

    doc.add_heading("Onboarding portals", level=1)
    doc.add_paragraph("Scan a code with a phone, sign in with an Entra ID account that is a "
                      f"member of '{ctx['group']}', and follow the portal.")
    portals = ctx["portals"]
    t = doc.add_table(rows=2, cols=len(portals))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, prt in enumerate(portals):
        top = t.rows[0].cells[i]
        top.text = ""
        hp = top.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = hp.add_run(prt["title"])
        r.bold = True
        r.font.size = Pt(13)
        r.font.color.rgb = RGBColor.from_string(ACCENT)
        ip = top.add_paragraph()
        ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
        ip.add_run().add_picture(_qr_png(prt["url"]), width=Mm(58))
        sp = top.add_paragraph()
        sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sp.add_run(f"SSID: {prt['ssid']}").bold = True
        bot = t.rows[1].cells[i]
        bot.text = ""
        up = bot.paragraphs[0]
        ur = up.add_run(prt["url"])
        ur.font.size = Pt(8)
        hw = bot.add_paragraph()
        hr = hw.add_run(prt["how"])
        hr.font.size = Pt(9)
        for row in t.rows:
            row.cells[i].width = Mm(87)
    _fixed_width(t, [87] * len(portals))
    doc.add_paragraph()

    doc.add_heading("What was built, step by step", level=1)
    for n, st in enumerate(ctx["stages"], 1):
        doc.add_heading(f"Step {n}. {st['title']}", level=2)
        doc.add_paragraph(st["explain"])
        rows = [(s.capitalize(), f"{k}: {nm}", ex) for s, k, nm, ex in st["events"]]
        if rows:
            _table(doc, ["Result", "Object", "Detail"], rows, [24, 82, 68])
        else:
            doc.add_paragraph("No changes were needed.")

    doc.add_heading("Notes", level=1)
    if ctx.get("radius_note"):
        p = doc.add_paragraph(style="List Bullet")
        p.add_run("First use of Access Assurance on this org: ").bold = True
        p.add_run(ctx["radius_note"])
    for line in (
        "Username/password (EAP-TTLS) uses the Entra ROPC flow, which skips MFA. Exclude the "
        "Mist Access Assurance source IPs from Conditional Access MFA, or use certificates only.",
        "Entra signs the whole SAML response and sends the user principal name as the NameID "
        "(emailAddress format); Mist requires both.",
        "Certificates are issued by the Mist organisation's internal (onboard) CA with the "
        "user's UPN as the common name, so Access Assurance finds the user in Entra ID by realm.",
        "Demo tip: if a phone or browser was used in an earlier demo, open the portal links in a "
        "new private/incognito tab. The portals reuse a cached sign-in, which belongs to the "
        "previous (deleted) portal and makes the page blank or the Marvis app fail.",
        "To return the Entra tenant and the Mist organisation to their state before the demo, "
        "run the script with --teardown. It uses the snapshot taken at the start of the build.",
    ):
        doc.add_paragraph(line, style="List Bullet")

    doc.save(path)
    return path
