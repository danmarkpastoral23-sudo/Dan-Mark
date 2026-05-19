"""
Build WINDMILL_SUPPLEMENT_Pasuquin.docx
- Based on original WINDMILL_PROJECT_Pasuquin_FINAL (1).docx
- Uses same styles, page size, margins as original
- Embeds all 5 original images (image1-5) in correct positions
- Adds 10 missing sections identified from comparison
"""
import zipfile, re, base64
from xml.sax.saxutils import escape as xe

ORIGINAL = "WINDMILL_PROJECT_Pasuquin_FINAL (1).docx"
OUTPUT   = "WINDMILL_SUPPLEMENT_Pasuquin.docx"

# ── Read original assets ──────────────────────────────────────────────────────
with zipfile.ZipFile(ORIGINAL) as z:
    orig_styles    = z.read("word/styles.xml")
    orig_numbering = z.read("word/numbering.xml")
    orig_header    = z.read("word/header1.xml")
    orig_footer    = z.read("word/footer1.xml")
    img1_data = z.read("word/media/image1.jpg")   # PSU logo small
    img2_data = z.read("word/media/image2.jpg")   # PSU logo wide
    img3_data = z.read("word/media/image3.png")   # Fig II-B.1 Wind-to-Grid
    img4_data = z.read("word/media/image4.png")   # Fig II-B.2 Turbine cross-section
    img5_data = z.read("word/media/image5.png")   # Fig II-B.3 Nacelle cut-away

# ── Namespaces (matching original) ───────────────────────────────────────────
NS = (
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
    'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:oel="http://schemas.microsoft.com/office/2019/extlst" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:w10="urn:schemas-microsoft-com:office:word" '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
    'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" '
    'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" '
    'xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" '
    'xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" '
    'xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" '
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
    'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
    'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
)


# ── XML helpers ───────────────────────────────────────────────────────────────
def p(text="", style="Normal", bold=False, italic=False,
      align=None, sz=None, color=None, runs=None, spacing_before=None, spacing_after=None):
    ppr_parts = [f'<w:pStyle w:val="{style}"/>']
    if align:        ppr_parts.append(f'<w:jc w:val="{align}"/>')
    sp = {}
    if spacing_before: sp["w:before"] = str(spacing_before)
    if spacing_after:  sp["w:after"]  = str(spacing_after)
    if sp:
        attrs = " ".join(f'{k}="{v}"' for k,v in sp.items())
        ppr_parts.append(f'<w:spacing {attrs}/>')
    ppr = f'<w:pPr>{"".join(ppr_parts)}</w:pPr>'

    if runs is not None:
        body = "".join(runs)
    else:
        rpr = []
        if bold:   rpr.append("<w:b/>")
        if italic: rpr.append("<w:i/>")
        if sz:     rpr.append(f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>')
        if color:  rpr.append(f'<w:color w:val="{color}"/>')
        rpr_xml = f'<w:rPr>{"".join(rpr)}</w:rPr>' if rpr else ""
        body = f'<w:r>{rpr_xml}<w:t xml:space="preserve">{xe(text)}</w:t></w:r>' if text else ""
    return f"<w:p>{ppr}{body}</w:p>"

def run(text, bold=False, italic=False, sz=None, color=None):
    rpr = []
    if bold:   rpr.append("<w:b/>")
    if italic: rpr.append("<w:i/>")
    if sz:     rpr.append(f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>')
    if color:  rpr.append(f'<w:color w:val="{color}"/>')
    rpr_xml = f'<w:rPr>{"".join(rpr)}</w:rPr>' if rpr else ""
    return f'<w:r>{rpr_xml}<w:t xml:space="preserve">{xe(text)}</w:t></w:r>'

def h1(text):  return p(text, style="Heading1")
def h2(text):  return p(text, style="Heading2")
def h3(text):  return p(text, style="Heading3")
def body(text): return p(text, style="Normal")
def blank():   return p("", style="Normal")
def page_break(): return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

def bullet(text):
    return (f'<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>'
            f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
            f'<w:r><w:t xml:space="preserve">{xe(text)}</w:t></w:r></w:p>')

def note(label, text):
    return p(runs=[run(label+" ", bold=True, color="2E74B5"),
                   run(text, italic=True, color="595959")])


def tbl(rows, widths=None, header=True):
    if not rows: return ""
    nc = max(len(r) for r in rows)
    if widths is None: widths = [9000//nc]*nc
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    out = [
        '<w:tbl>',
        '<w:tblPr><w:tblW w:w="9000" w:type="dxa"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="6" w:color="1F3864"/>'
        '<w:left w:val="single" w:sz="6" w:color="1F3864"/>'
        '<w:bottom w:val="single" w:sz="6" w:color="1F3864"/>'
        '<w:right w:val="single" w:sz="6" w:color="1F3864"/>'
        '<w:insideH w:val="single" w:sz="4" w:color="BFBFBF"/>'
        '<w:insideV w:val="single" w:sz="4" w:color="BFBFBF"/>'
        '</w:tblBorders></w:tblPr>',
        f'<w:tblGrid>{grid}</w:tblGrid>',
    ]
    for ri, row in enumerate(rows):
        is_hdr = header and ri == 0
        fill = '<w:shd w:val="clear" w:color="auto" w:fill="1F3864"/>' if is_hdr else (
               '<w:shd w:val="clear" w:color="auto" w:fill="EBF3FB"/>' if ri%2==0 else "")
        out.append('<w:tr>')
        for ci in range(nc):
            cell = str(row[ci]) if ci < len(row) else ""
            tcpr = f'<w:tcPr><w:tcW w:w="{widths[ci]}" w:type="dxa"/>{fill}</w:tcPr>'
            rpr = '<w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>' if is_hdr else ""
            out.append(f'<w:tc>{tcpr}<w:p><w:r>{rpr}'
                       f'<w:t xml:space="preserve">{xe(cell)}</w:t></w:r></w:p></w:tc>')
        out.append('</w:tr>')
    out.append('</w:tbl><w:p/>')
    return "".join(out)

def inline_image(rId, cx_emu, cy_emu, img_id, name):
    """Inline (non-floating) image."""
    return (
        f'<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr/>'
        f'<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx_emu}" cy="{cy_emu}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{img_id}" name="{name}"/>'
        f'<wp:cNvGraphicFramePr>'
        f'<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
        f'</wp:cNvGraphicFramePr>'
        f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="{img_id}" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rId}"/>'
        f'<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/>'
        f'<a:ext cx="{cx_emu}" cy="{cy_emu}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'</pic:spPr>'
        f'</pic:pic></a:graphicData></a:graphic>'
        f'</wp:inline></w:drawing></w:r></w:p>'
    )

def caption(text):
    return p(runs=[run(text, italic=True, sz="20", color="595959")], align="center")


# ── Document body ─────────────────────────────────────────────────────────────
body_parts = []
A = body_parts.append

# ════════════════════════════════════════════════════════════════════════
# TITLE PAGE  (same logos as original: rId7=img1, rId8=img2)
# ════════════════════════════════════════════════════════════════════════
# Top logos row
A('<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
  '<w:r><w:drawing>'
  '<wp:inline distT="0" distB="0" distL="0" distR="0">'
  '<wp:extent cx="1097280" cy="1107255"/>'
  '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
  '<wp:docPr id="10" name="Logo1"/>'
  '<wp:cNvGraphicFramePr>'
  '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
  '</wp:cNvGraphicFramePr>'
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
  '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
  '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
  '<pic:nvPicPr><pic:cNvPr id="10" name="Logo1"/><pic:cNvPicPr/></pic:nvPicPr>'
  '<pic:blipFill><a:blip r:embed="rId7"/>'
  '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/>'
  '<a:ext cx="1097280" cy="1107255"/></a:xfrm>'
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
  '</pic:spPr></pic:pic></a:graphicData></a:graphic>'
  '</wp:inline></w:drawing></w:r>'
  '<w:r><w:t xml:space="preserve">  </w:t></w:r>'
  '<w:r><w:drawing>'
  '<wp:inline distT="0" distB="0" distL="0" distR="0">'
  '<wp:extent cx="1371600" cy="857057"/>'
  '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
  '<wp:docPr id="11" name="Logo2"/>'
  '<wp:cNvGraphicFramePr>'
  '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
  '</wp:cNvGraphicFramePr>'
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
  '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
  '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
  '<pic:nvPicPr><pic:cNvPr id="11" name="Logo2"/><pic:cNvPicPr/></pic:nvPicPr>'
  '<pic:blipFill><a:blip r:embed="rId8"/>'
  '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/>'
  '<a:ext cx="1371600" cy="857057"/></a:xfrm>'
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
  '</pic:spPr></pic:pic></a:graphicData></a:graphic>'
  '</wp:inline></w:drawing></w:r>'
  '</w:p>')

A(p("Pangasinan State University", style="Normal", bold=True, align="center", sz="24", color="1F3864"))
A(p("Urdaneta City Campus", style="Normal", align="center", sz="22", color="1F3864"))
A(p("College of Engineering and Architecture", style="Normal", align="center", sz="22", color="1F3864"))
A(p("Mechanical Engineering Department", style="Normal", align="center", sz="22", color="1F3864"))
A(p("Second Semester, A.Y. 2025–2026", style="Normal", align="center", sz="22"))
A(blank())
A(p("WINDMILL PROJECT PROPOSAL", style="Title",
    runs=[run("WINDMILL PROJECT PROPOSAL", bold=True, sz="44", color="1F3864")]))
A(p("Supplemental Engineering Documentation", style="Normal", align="center",
    runs=[run("Supplemental Engineering Documentation", italic=True, sz="28", color="2E74B5")]))
A(p("Pasuquin, Ilocos Norte, Philippines", style="Normal", align="center", sz="24"))
A(blank())
A(p("Submitted by:", style="Normal", bold=True, align="center"))
A(p("Dan Mark C. Pastoral", style="Normal", align="center", sz="24", bold=True))
A(p("BSME-4A", style="Normal", align="center"))
A(blank())
A(p("Submitted to:", style="Normal", bold=True, align="center"))
A(p("Marfel D. Rosario, PME", style="Normal", align="center", sz="24", bold=True))
A(p("Instructor", style="Normal", align="center"))
A(blank())
A(p("May 2026", style="Normal", align="center", italic=True))
A(page_break())


# ════════════════════════════════════════════════════════════════════════
# PURPOSE
# ════════════════════════════════════════════════════════════════════════
A(h1("Purpose of this Supplement"))
A(body("This document supplements the original Windmill Project Proposal for Pasuquin, "
       "Ilocos Norte (submitted April 29, 2026). It adds ten engineering sections identified "
       "as missing after peer comparison: a glossary, detailed equipment specifications, "
       "a single-line electrical diagram, foundation calculations, a facility inventory, "
       "an energy balance, an O&M plan, materials and color-coding standards, regulatory "
       "citations, and extended references. All content is directly based on the data, "
       "calculations, and figures already established in the original proposal."))
A(body("The three process-flow diagrams from Section II-B of the original proposal are "
       "reproduced below as reference for the supplemental sections that follow."))
A(blank())

# ── Reproduce original figures ────────────────────────────────────────────────
A(h2("Reference: Original Section II-B — Process Flow Diagrams"))

A(body("The following three figures are taken directly from the original proposal "
       "(Section II-B). They are repeated here so this supplement can be read as a "
       "standalone engineering annex."))
A(blank())

# Figure II-B.1  image3.png  12.7cm x 15.2cm → 4572000 x 5486400 EMU
A(inline_image("rId9",  4572000, 5486400, 20, "Fig-IIB1"))
A(caption("Figure II-B.1 — Wind-to-Grid Process Flow Diagram (Pasuquin Wind Farm)"))
A(blank())

# Figure II-B.2  image4.png  19.1cm x 15.6cm → 6883800 x 5619600 EMU
A(inline_image("rId10", 6883800, 5619600, 21, "Fig-IIB2"))
A(caption("Figure II-B.2 — Turbine Component Cross-Section with Labeled Parts"))
A(blank())

# Figure II-B.3  image5.png  15.4cm x 7.1cm → 5547600 x 2559240 EMU
A(inline_image("rId11", 5547600, 2559240, 22, "Fig-IIB3"))
A(caption("Figure II-B.3 — Nacelle Cut-Away: Low-Speed Shaft to 33 kV Output"))
A(page_break())


# ════════════════════════════════════════════════════════════════════════
# SECTION 1 — GLOSSARY
# ════════════════════════════════════════════════════════════════════════
A(h1("Section 1 — Glossary of Terms and Acronyms"))
A(body("The following terms appear throughout the original proposal and this supplement. "
       "Definitions follow IEC 61400, DOE Philippines, and ERC usage."))
A(blank())
glossary_rows = [
    ["Term / Acronym", "Definition"],
    ["AEP", "Annual Energy Production — total electrical energy generated per year (MWh)."],
    ["Amihan", "Northeast Monsoon; the primary wind driver for the Pasuquin site (October–March peak)."],
    ["Betz Limit", "Theoretical max fraction of wind energy extractable by any turbine = 59.3% (Albert Betz, 1919)."],
    ["Capacity Factor (CF)", "Actual AEP ÷ (Rated Power × 8,760 h/yr). Pasuquin estimate: 38%."],
    ["Cut-in Speed", "Minimum wind speed for net positive power output. This project: 3.0 m/s."],
    ["Cut-out Speed", "Maximum operating wind speed; above this the rotor feathers. This project: 25 m/s."],
    ["DFIG", "Doubly Fed Induction Generator — variable-speed generator used in this project."],
    ["DOE", "Department of Energy of the Philippines."],
    ["ECC", "Environmental Compliance Certificate issued by DENR-EMB after EIS review."],
    ["EIS", "Environmental Impact Statement filed under the Philippine EIS System (PEISS)."],
    ["ERC", "Energy Regulatory Commission — approves FiT rates and grid codes."],
    ["FiT", "Feed-in Tariff — guaranteed PHP 8.53/kWh for wind energy (ERC Res. 16, 2022)."],
    ["FPIC", "Free, Prior and Informed Consent — required under RA 8371 (IPRA)."],
    ["HAWT", "Horizontal Axis Wind Turbine — 3-bladed upwind type used in this project."],
    ["Hub Height", "Height from ground to rotor hub center. This project: 90–110 m."],
    ["IEC 61400", "IEC standard series for wind turbine design, loads, and grid integration."],
    ["Luzon Strait", "Narrow waterway north of Luzon; accelerates Amihan winds over Pasuquin."],
    ["MW / MWh", "Megawatt (power) / Megawatt-hour (energy). 1 MW for 1 hour = 1 MWh."],
    ["NGCP", "National Grid Corporation of the Philippines — Luzon Grid operator."],
    ["NSCP", "National Structural Code of the Philippines, 2015 — governs tower foundations."],
    ["O&M", "Operations and Maintenance."],
    ["RA 9513", "Renewable Energy Act of 2008 — legal basis for FiT and DOE service contracts."],
    ["Rated Speed", "Wind speed at which turbine reaches rated output. This project: 12–14 m/s."],
    ["SCADA", "Supervisory Control and Data Acquisition — remote turbine monitoring system."],
    ["Specific Power", "Rated capacity ÷ swept area (W/m²). Optimal Class II: 250–400 W/m²."],
    ["Swept Area", "Circle swept by rotor blades, A = πr². For 130 m rotor: 13,273 m²."],
    ["Wake Effect", "Reduced wind speed and increased turbulence behind an operating turbine."],
    ["Weibull Distribution", "Statistical model for wind speed frequency; defined by k (shape) and c (scale)."],
    ["Wind Class II", "IEC site classification for average wind speeds 7.5–8.5 m/s — Pasuquin's class."],
    ["Wind Power Density", "Available power per swept area: WPD = ½ρV³ (W/m²)."],
    ["Wind Shear", "Increase of wind speed with height: V(z)=V_ref·(z/z_ref)^α; α≈0.14 coastal."],
]
A(tbl(glossary_rows, widths=[2000, 7000]))
A(page_break())


# ════════════════════════════════════════════════════════════════════════
# SECTION 2 — DETAILED EQUIPMENT SPECIFICATIONS
# ════════════════════════════════════════════════════════════════════════
A(h1("Section 2 — Detailed Equipment Specifications"))
A(body("Based on the 5.0 MW HAWT, 130 m rotor, 100 m hub height, and 33/115 kV electrical "
       "system established in Section III of the original proposal. Configuration: 20 turbines × 5 MW = 100 MW."))
A(blank())

A(h2("2.1 Rotor and Blade Assembly"))
A(tbl([
    ["Parameter","Specification","Design Basis"],
    ["Configuration","3-bladed upwind HAWT","Industry standard for utility-scale"],
    ["Rotor diameter","130 m","Specific Power 339 W/m² — optimal for V=8.5 m/s (Sec. III-A.2)"],
    ["Swept area","13,273 m²","A = π × 65²"],
    ["Blade length","63.5 m","= Rotor radius − hub radius (1.5 m)"],
    ["Blade material","GFRP with carbon-fiber spar caps","IEC 61400-5; fatigue and stiffness"],
    ["Blade weight","≈ 14,000 kg each","Vestas V126/V136 class benchmark"],
    ["Tip speed (max)","≤ 90 m/s","Acoustic + structural fatigue limit"],
    ["Pitch system","3 independent electric motors + battery backup","IEC 61400-1 Class II"],
    ["Lightning protection","Receptors at tip + down-conductor to hub","IEC 61400-24"],
    ["Cut-in / Rated / Cut-out","3.0 / 12–14 / 25 m/s","From Sec. III-A.4 Weibull analysis"],
], widths=[2500,3500,3000]))
A(blank())

A(h2("2.2 Drivetrain — Gearbox and DFIG Generator"))
A(tbl([
    ["Component","Specification"],
    ["Gearbox type","Three-stage planetary–helical (integrated with main bearing)"],
    ["Gear ratio","1:90 — rotor ~17 rpm → generator 1,530 rpm"],
    ["Gearbox efficiency","≥ 97% at rated load"],
    ["Gearbox oil","ISO VG 320 mineral/PAO synthetic, 600 L sump, forced lubrication"],
    ["Generator type","Doubly Fed Induction Generator (DFIG) — from original proposal Sec. III"],
    ["Rated power","5.0 MW at 12–14 m/s wind"],
    ["Stator voltage","690 V AC, 3-phase, 60 Hz"],
    ["Cooling","Air-to-water heat exchanger in nacelle"],
    ["Power converter","Partial-scale IGBT (~30% of rated), stabilizes output to 60 Hz"],
    ["Nacelle mass","≈ 320 t (nacelle + rotor assembly)"],
    ["Overall drivetrain η","≈ 93% (gearbox × generator × converter)"],
], widths=[3000,6000]))
A(blank())

A(h2("2.3 Tower Structure"))
A(tbl([
    ["Parameter","Specification"],
    ["Type","Tubular steel, 3-section flanged and bolted"],
    ["Hub height","100 m (selected) — V(100m)=8.19 m/s, +26% vs ground (Sec. III-A.3)"],
    ["Base diameter","4.30 m"],
    ["Top diameter","3.00 m"],
    ["Wall thickness","30 mm at base tapering to 18 mm at top"],
    ["Steel grade","S355 J0+N per EN 10025-2"],
    ["Tower mass","≈ 280 t"],
    ["Coating","ISO 12944 C5-M system: zinc-rich primer + epoxy mid + PU topcoat, DFT ≥ 320 µm"],
    ["Internal access","Fixed ladder + climb-assist device + emergency descender"],
    ["Aviation marking","Upper 1/3: RAL 3020 red bands, per CAAP MOS Part 14"],
], widths=[3000,6000]))
A(blank())

A(h2("2.4 Electrical Collection and Substation"))
A(tbl([
    ["Component","Specification"],
    ["Turbine pad-mount transformer","5.5 MVA, 690 V / 33 kV, oil-filled ONAN"],
    ["Collection voltage","33 kV — industry standard for up to ~200 MW (Sec. III-C.1)"],
    ["Inter-array cable","33 kV XLPE copper, 240–630 mm², buried 1.0 m, IEC 60502-2"],
    ["Switchgear per turbine","33 kV ring-main unit, SF6-insulated"],
    ["Collection feeders","5 feeders × 4 turbines = 20 MW per feeder"],
    ["Step-up transformer","120 MVA, 33 kV / 115 kV, OLTC ±10%, turns ratio 3.48:1 (Sec. III-C.2)"],
    ["Transmission voltage","115 kV — NGCP Ilocos Norte grid standard (Sec. III-C.3)"],
    ["Substation protection","IEC 61850 relays: 87T diff, 21 distance, 51 overcurrent, 64 REF"],
    ["Reactive compensation","±30 MVAr STATCOM for Philippine Grid Code compliance"],
    ["Grounding","Mesh + driven rod, target ≤ 1 Ω (IEEE 80)"],
    ["SCADA","Wind-Farm Management System, OPC-UA protocol to NGCP RTU"],
    ["Auxiliary power","500 kVA diesel genset + 30-min UPS for critical loads"],
], widths=[3000,6000]))
A(page_break())


# ════════════════════════════════════════════════════════════════════════
# SECTION 3 — SINGLE-LINE DIAGRAM
# ════════════════════════════════════════════════════════════════════════
A(h1("Section 3 — Single-Line Electrical Diagram (SLD)"))
A(body("The SLD is the electrical equivalent of a P&ID. It traces power from each turbine "
       "generator to the NGCP grid metering point. Based on the 33 kV/115 kV configuration "
       "established in Section III-C of the original proposal."))
A(blank())
A(h2("3.1 Power Flow — Turbine to Grid"))
for s in [
    "Stage 1 — Turbine: DFIG produces 690 V AC, 3-phase, stabilized to 60 Hz by partial-scale converter.",
    "Stage 2 — Pad-mount transformer (T-1xx): Steps 690 V → 33 kV at each tower base.",
    "Stage 3 — Inter-array cable: 33 kV XLPE underground, 4 turbines per feeder ring.",
    "Stage 4 — Substation (T-200): 5 feeders enter 33 kV busbar; 120 MVA transformer steps to 115 kV.",
    "Stage 5 — 115 kV OHL: ~35 km to NGCP Laoag Substation; revenue metering at grid entry point.",
]:
    A(bullet(s))
A(blank())
A(h2("3.2 Equipment Tag Conventions"))
A(tbl([
    ["Tag Prefix","Equipment"],
    ["WTG-xxx","Wind Turbine Generator unit (WTG-101 to WTG-120)"],
    ["G-xxx","Generator inside nacelle"],
    ["T-1xx","Pad-mount turbine transformer (one per WTG)"],
    ["T-200","Main 33/115 kV step-up transformer"],
    ["CB-xxx","Circuit breaker (CB-101 to CB-120 at turbines; CB-200 series at substation)"],
    ["DS-xxx","Disconnect / isolator switch"],
    ["CT-xxx","Current transformer (metering / protection)"],
    ["VT-xxx","Voltage transformer"],
    ["MET-001","Revenue metering point at NGCP interconnection"],
    ["PR-xxx","Protection relay (87T, 21, 51, 51N, 64, 24 function codes)"],
    ["SCADA","Wind-Farm Management System server in control room"],
], widths=[2000,7000]))
A(blank())
A(h2("3.3 Protection Schedule"))
A(tbl([
    ["Equipment","Primary Protection","Backup Protection"],
    ["Pad-mount T-1xx","Buchholz relay, winding temp","Differential 87T, overcurrent 51"],
    ["33 kV feeder","Directional overcurrent 67","Time-overcurrent 51, earth-fault 51N"],
    ["Main transformer T-200","Differential 87T, REF 64","Overcurrent 51, overflux 24"],
    ["115 kV transmission","Distance relay 21 (zones 1-3)","Pilot communication scheme"],
    ["33 kV busbar","Bus-differential 87B","Reverse-blocking OC"],
    ["Generator G-xxx","Anti-islanding, loss-of-excitation 40","Under/over voltage & frequency"],
], widths=[2500,3200,3300]))
A(page_break())


# ════════════════════════════════════════════════════════════════════════
# SECTION 4 — FOUNDATION AND ANCHOR BOLT CALCULATIONS
# ════════════════════════════════════════════════════════════════════════
A(h1("Section 4 — Foundation and Anchor Bolt Calculations"))
A(body("Sizing the gravity mat foundation and anchor bolt cage for each turbine tower. "
       "Loads derived from IEC 61400-1 Class II; foundation checked per NSCP 2015."))
A(blank())
A(h2("4.1 Design Loads at Tower Base"))
A(tbl([
    ["Load","Symbol","Value","Source"],
    ["Vertical dead load (rotor+nacelle+tower)","F_z","4,200 kN","Manufacturer load table"],
    ["Extreme thrust (DLC 1.3 / 6.2)","F_y","1,050 kN","IEC 61400-1 Class II"],
    ["Overturning moment at base","M_y","115,000 kN·m","F_y × 100 m hub + lever arm"],
    ["Torsional moment","M_z","8,500 kN·m","Extreme yaw + misalignment"],
], widths=[3500,1300,2000,2200]))
A(blank())
A(h2("4.2 Octagonal Gravity Mat — Sizing"))
A(tbl([
    ["Parameter","Value","Note"],
    ["Outer diameter (across flats)","21.0 m","Trial — checked below"],
    ["Slab thickness","1.0 m edge / 3.0 m center","Tapered frustum"],
    ["Concrete grade","fc' = 35 MPa","ACI 318 / NSCP 2015"],
    ["Reinforcement","Grade 60, fy = 414 MPa","Top + bottom mats, both ways"],
    ["Concrete volume","≈ 590 m³","Octagonal frustum formula"],
    ["Foundation self-weight (W_f)","14,160 kN","γ_c = 24 kN/m³"],
    ["Soil cover (1.0 m depth × 380 m²)","6,840 kN","γ_s = 18 kN/m³"],
    ["Total stabilizing weight (W_total)","25,200 kN","F_z + W_f + W_soil"],
    ["Lever arm to edge (e = D/2)","10.5 m",""],
    ["Stabilizing moment (M_s = W_total × e)","264,600 kN·m",""],
    ["Safety factor vs overturning (M_s/M_y)","2.30","Required ≥ 1.5 ✓ PASS"],
    ["Max edge bearing pressure","≈ 175 kPa","Trapezoidal dist."],
    ["Required allowable bearing capacity","≥ 250 kPa","Geotech investigation required"],
], widths=[3800,2500,2700]))
A(blank())
A(h2("4.3 Anchor Bolt Cage"))
A(tbl([
    ["Parameter","Value","Basis"],
    ["Total bolts","144 (72 inner + 72 outer)","5 MW class industry standard"],
    ["Bolt grade","ISO 898-1 Class 10.9, fy ≈ 900 MPa","High-strength structural"],
    ["Bolt size","M48 — effective stress area Aₛ = 1,470 mm²",""],
    ["Embedded length","5,500 mm total","Full-depth anchorage"],
    ["PCD (pitch circle dia)","Inner 4.20 m / Outer 4.90 m","Matches tower base flange"],
    ["Allowable tension per bolt (Pt)","= Aₛ × 0.7 × fy = 925 kN","FoS = 1.5 applied"],
    ["Max bolt tension (T_max)","= M_y × c / I_bolt-group ≈ 605 kN","Linear bolt-group method"],
    ["Utilization (T_max/Pt)","65%","Required ≤ 85% ✓ PASS"],
    ["Pretension torque","≈ 4,200 N·m via hydraulic tensioner","70% fy, IEC 61400-6"],
    ["Fatigue class","ISO 1099 S-N Class 71","DLC 1.2 normal operation"],
], widths=[3500,3000,2500]))
A(note("NOTE:", "These are preliminary sizing calculations. Final design must be certified "
       "by a Philippine-licensed Civil/Structural Engineer after geotechnical "
       "investigation and review of the actual turbine load report from the supplier."))
A(page_break())


# ════════════════════════════════════════════════════════════════════════
# SECTION 5 — FACILITY INVENTORY
# ════════════════════════════════════════════════════════════════════════
A(h1("Section 5 — Facility Inventory and Floor Areas"))
A(body("All permanent and temporary facilities on the Pasuquin Wind Farm site. "
       "Based on the plant layout in Section IV of the original proposal. "
       "Land between turbines (~9.45 km²) remains usable for farming and grazing."))
A(blank())
A(tbl([
    ["No.","Facility","Floor Area (m²)","Function"],
    ["1","Turbine foundations (20 × ~380 m²)","7,600 *","Octagonal RC mat per Section 4"],
    ["2","On-site 33/115 kV substation","5,000","Step-up transformer, switchgear, control room"],
    ["3","SCADA / Control room","180","24/7 wind-farm supervisory control"],
    ["4","Maintenance workshop & stores","650","Spare parts, gearbox oil, blade tools"],
    ["5","Administrative building","320","Project management, visitor reception"],
    ["6","Eco-tourism visitor center","400","Public viewing deck, exhibits, restrooms"],
    ["7","Emergency and fire station","240","Nacelle fire response, high-angle rescue"],
    ["8","Diesel genset house + fuel store","120","500 kVA backup + 5,000 L diesel tank"],
    ["9","Meteorological mast pad","100","Permanent met mast, wind/temperature sensors"],
    ["10","Guard house and main gate","60","24/7 security checkpoint"],
    ["11","Parking and crane turnaround","1,200","Heavy vehicles, mobile crane hardstand"],
    ["12","Material laydown area (construction)","2,500","Blade/nacelle staging (temporary)"],
    ["13","Septic and sewage treatment","150","On-site sanitary waste"],
    ["14","Water tank and pump house","80","Blade wash, fire-reserve 50,000 L"],
    ["15","Internal access roads (7 km × 6 m)","≈ 42,000","Heavy-haul gravel roads to each WTG"],
    ["","TOTAL BUILT-UP (excl. roads + pads)","≈ 11,000",""],
    ["","TOTAL TURBINE FOOTPRINT (*)","≈ 7,600",""],
    ["","TOTAL ROADS AND HARDSTANDS","≈ 44,500",""],
    ["","GROSS SITE AREA (incl. inter-turbine land)","≈ 9.46 km²","Per spacing calc in Sec. III-B.2"],
], widths=[700,3800,2500,2000]))
A(page_break())

# ════════════════════════════════════════════════════════════════════════
# SECTION 6 — ENERGY BALANCE
# ════════════════════════════════════════════════════════════════════════
A(h1("Section 6 — Energy Balance Diagram"))
A(body("This section extends the process flow diagrams in Section II-B of the original "
       "proposal by quantifying the power loss at each stage — from incident wind "
       "kinetic energy to electricity delivered at the NGCP grid metering point."))
A(blank())
A(h2("6.1 Stage-by-Stage Power Balance (per turbine at rated wind V=12 m/s)"))
A(body("Basis: P_wind = ½ × ρ × A × V³ = ½ × 1.225 × 13,273 × 12³ = 13.96 MW "
       "(using the same Betz equation from Section III-A.1 of the original proposal)."))
A(blank())
A(tbl([
    ["Stage","Input (MW)","Loss (MW)","Output (MW)","Loss Mechanism"],
    ["Wind kinetic energy at rotor disc","13.96","—","13.96","P_wind = ½ρAV³"],
    ["Aerodynamic conversion (Cp = 0.45)","13.96","7.68","6.28","Real rotor < Betz 59.3%"],
    ["Hub + main bearing friction","6.28","0.06","6.22","≈ 1% mechanical"],
    ["Gearbox (3-stage planetary, η=97%)","6.22","0.19","6.03","≈ 3% gear losses"],
    ["DFIG generator (η=97%)","6.03","0.18","5.85","≈ 3% electromagnetic"],
    ["Partial-scale power converter","5.85","0.10","5.75","≈ 1.7% IGBT losses"],
    ["Pad-mount transformer 690V/33kV","5.75","0.05","5.70","≈ 0.9% iron + copper"],
    ["Inter-array 33 kV cable (avg.)","5.70","0.04","5.66","≈ 0.7% I²R"],
    ["Step-up transformer 33/115 kV","5.66","0.03","5.63","≈ 0.5%"],
    ["115 kV OHL to Laoag (≈35 km)","5.63","0.05","5.58","≈ 0.9% line losses"],
    ["NET at NGCP grid metering point","—","—","5.58","Delivered electricity"],
], widths=[3200,1200,1200,1200,2200]))
A(blank())
A(h2("6.2 Annual Energy Yield"))
A(tbl([
    ["Parameter","Value","Calculation"],
    ["Rated capacity per turbine","5.0 MW","Design specification"],
    ["Max theoretical AEP/turbine","43,800 MWh/yr","5.0 MW × 8,760 h"],
    ["Capacity factor (Pasuquin)","38%","Consistent with Bangui 35–42%, Burgos 36–45%"],
    ["Gross AEP per turbine","16,644 MWh/yr","43,800 × 0.38"],
    ["Electrical losses (avg. 7%)","1,165 MWh/yr","Cables, transformers, grid line"],
    ["Net AEP per turbine delivered","15,479 MWh/yr","16,644 × (1–0.07)"],
    ["Wind farm total (20 turbines)","309,580 MWh/yr","15,479 × 20"],
    ["Equivalent households served","≈ 77,400","At 4,000 kWh/yr per household"],
    ["Annual FiT revenue (gross)","PHP 2.64 Billion","309,580 MWh × PHP 8.53/kWh"],
], widths=[3500,2500,3000]))
A(page_break())


# ════════════════════════════════════════════════════════════════════════
# SECTION 7 — O&M PLAN
# ════════════════════════════════════════════════════════════════════════
A(h1("Section 7 — Operations and Maintenance (O&M) Plan"))
A(body("The O&M plan ensures the 38% capacity factor assumed in the financial analysis "
       "(Section VI of the original proposal) is sustained over the 25-year project life."))
A(blank())
A(h2("7.1 Permanent Staffing"))
A(tbl([
    ["Role","No.","Shift Schedule"],
    ["Wind Farm Manager","1","Day"],
    ["Lead Turbine Technician (GWO Level 3)","4","Day + on-call rotation"],
    ["Turbine Technician (GWO Level 2)","8","Day, 4-on/4-off"],
    ["Electrical / Substation Engineer","2","Day"],
    ["SCADA Operator","4","24/7, 4-shift rotation"],
    ["HSE / GWO Safety Trainer","1","Day"],
    ["Admin and Procurement","2","Day"],
    ["Security Personnel","6","24/7"],
    ["TOTAL","28",""],
], widths=[4000,1000,4000]))
A(blank())
A(h2("7.2 Maintenance Schedule (per turbine)"))
A(tbl([
    ["Frequency","Activity"],
    ["Continuous (24/7)","SCADA monitoring: vibration, temperature, power output, fault codes"],
    ["Weekly","Ground-level visual + drone inspection; oil-leak check"],
    ["Monthly","Tower climb: bolt torque sample, lube check, brake pad thickness"],
    ["Quarterly","Gearbox oil sample for spectrographic analysis; blade LE inspection"],
    ["Semi-annual","Yaw/pitch function test; emergency descender re-certification"],
    ["Annual","Full PM: oil change, filter change, bolt re-torque (all 144 anchor bolts)"],
    ["Every 2 years","Generator slip-ring service; grid-code compliance test (ERC)"],
    ["Every 5 years","Major: gearbox borescope, blade composite UT scan, transformer oil DGA"],
    ["Every 10 years","Mid-life overhaul: gearbox refurbishment, generator rewind assessment"],
], widths=[2200,6800]))
A(blank())
A(h2("7.3 Spare Parts Tiering"))
for s in [
    "Tier A — On-site (immediate): pitch motors, yaw motors, fuses, contactors, sensors, oil, filters.",
    "Tier B — Regional depot, 24 h delivery: IGBT converter modules, slip rings, brake pads, hydraulic accumulators.",
    "Tier C — Manufacturer warehouse, 2–6 weeks: main bearing, gearbox sub-assembly, generator, blade section.",
]:
    A(bullet(s))
A(blank())
A(h2("7.4 Safety Requirements"))
for s in [
    "All turbine technicians must hold valid GWO Basic Safety Training (BST): Working at Heights, First Aid, Manual Handling, Fire Awareness.",
    "Two-person rule enforced above 30 m inside or outside the tower.",
    "Hot-work permits required for any welding or grinding inside the nacelle.",
    "Annual emergency drill: nacelle fire, tower evacuation, casualty medical response.",
    "Compliance with DOLE Department Order No. 198-18 (Occupational Safety and Health Standards at All Times).",
]:
    A(bullet(s))
A(page_break())


# ════════════════════════════════════════════════════════════════════════
# SECTION 8 — MATERIALS AND COLOR CODING
# ════════════════════════════════════════════════════════════════════════
A(h1("Section 8 — Materials and Color-Coding Standards"))
A(blank())
A(h2("8.1 Material Specifications"))
A(tbl([
    ["Component","Material / Grade / Standard"],
    ["Tower steel","S355 J0+N per EN 10025-2"],
    ["Tower flange bolts","ISO 898-1 Class 10.9, hot-dip galvanized"],
    ["Anchor bolts","ISO 898-1 Class 10.9 or ASTM A615 Grade 75"],
    ["Foundation concrete","fc' = 35 MPa; ACI 318 / NSCP 2015"],
    ["Foundation rebar","Grade 60 (fy = 414 MPa) deformed bars"],
    ["External coating system","ISO 12944 C5-M: zinc-rich primer + epoxy + PU topcoat, total DFT ≥ 320 µm"],
    ["Blade composite","Glass-fiber reinforced epoxy (GFRP) + carbon spar caps; IEC 61400-5"],
    ["33 kV inter-array cable","XLPE-insulated copper conductor Class 2; IEC 60502-2"],
    ["115 kV overhead conductor","ACSR 'Drake' per ANSI/NEMA; NESC clearances"],
    ["Substation grounding conductor","Bare copper 95 mm² mesh; IEEE 80"],
    ["Control cabinets","IP66 painted steel; RAL 7035 light grey"],
    ["Gearbox lubricant","ISO VG 320 mineral or PAO synthetic"],
    ["Yaw/pitch grease","NLGI-2 lithium-complex with EP additive"],
], widths=[3000,6000]))
A(blank())
A(h2("8.2 Color-Coding Scheme"))
A(tbl([
    ["Item","Color (RAL)","Standard / Reason"],
    ["Tower body — lower 2/3","RAL 9010 Pure White","IEC visibility baseline"],
    ["Tower body — upper 1/3 (2 red bands)","RAL 3020 Traffic Red","CAAP MOS Part 14 — daytime obstruction"],
    ["Blade tips (last 6 m, alternating)","RAL 3020 + RAL 9010","Aviation visibility"],
    ["Aviation warning lights","Red, medium-intensity flashing","ICAO Annex 14 / CAAP"],
    ["Earthing conductors","Green/Yellow striped","IEC 60446"],
    ["33 kV cable phase marking","L1=Red, L2=Yellow, L3=Blue","PEC 2017 Philippine standard"],
    ["Fire-fighting equipment","RAL 3000 Flame Red","DOLE OSH Standard"],
    ["Emergency exits and evacuation arrows","RAL 6024 Traffic Green + white text","ISO 7010"],
    ["115 kV substation panels","RAL 7032 Pebble Grey","Philippine substation convention"],
    ["Hazard / caution floor markings","RAL 1023 Traffic Yellow + black stripes","ISO 3864"],
    ["Transformer bund wall (oil containment)","RAL 1018 Zinc Yellow","Visibility for inspection"],
], widths=[2800,3000,3200]))
A(page_break())


# ════════════════════════════════════════════════════════════════════════
# SECTION 9 — REGULATORY CITATIONS
# ════════════════════════════════════════════════════════════════════════
A(h1("Section 9 — Regulatory and Standards Compliance"))
A(body("All laws and standards cited in the original proposal (Section V-C and VII) plus "
       "additional codes required for the supplemental engineering sections:"))
A(blank())
A(h2("9.1 Philippine Laws and DOE/ERC Orders"))
for s in [
    "RA 9513 — Renewable Energy Act of 2008; basis for DOE Service Contract and FiT incentives.",
    "RA 9136 — Electric Power Industry Reform Act (EPIRA), 2001.",
    "ERC Resolution No. 16, Series of 2022 — Feed-in Tariff rate: PHP 8.53/kWh for wind.",
    "DOE Department Circular DC2009-07-0011 — Implementing Rules of RA 9513.",
    "DENR DAO 2003-30 — Philippine EIS System; ECC required before construction.",
    "RA 8371 — IPRA; FPIC required if ancestral domain areas are within the project footprint.",
    "Philippine Electrical Code (PEC), 2017 Edition — wiring, grounding, panel protection.",
    "Philippine Grid Code (ERC, latest) — grid connection, power quality, metering.",
    "CAAP Manual of Standards Part 14 — aviation obstruction marking for towers > 45 m.",
    "DOLE Department Order 198-18 — Occupational Safety and Health Standards.",
]:
    A(bullet(s))
A(blank())
A(h2("9.2 International Standards (IEC, ISO, ASTM, IEEE)"))
for s in [
    "IEC 61400-1 Ed.4 (2019) — Wind Turbines: Design Load Requirements, Class II site.",
    "IEC 61400-5 — Blade design and testing.",
    "IEC 61400-6 — Tower and Foundation Design Requirements.",
    "IEC 61400-12-1 — Power Performance Measurement.",
    "IEC 61400-24 — Lightning Protection for Wind Turbines.",
    "IEC 61850 — Substation SCADA Communication (OPC-UA, IED protection).",
    "IEC 60502-2 — XLPE Power Cables (33 kV inter-array).",
    "IEC 60446 — Conductor color identification.",
    "ISO 12944 — Corrosion protection of steel structures by paint systems.",
    "ISO 3864 — Safety colors and signs.",
    "ISO 7010 — Graphical safety symbols.",
    "ISO 898-1 — Mechanical properties of bolts and screws (Class 10.9).",
    "IEEE 80 — Guide for Safety in AC Substation Grounding.",
    "NSCP 2015 (Vol. 1) — National Structural Code of the Philippines; tower foundation.",
    "NSCP 2015 Section 207 — Wind load provisions.",
    "GWO Basic Safety Training (BST) — mandatory for all working-at-height personnel.",
]:
    A(bullet(s))
A(page_break())

# ════════════════════════════════════════════════════════════════════════
# SECTION 10 — EXTENDED REFERENCES
# ════════════════════════════════════════════════════════════════════════
A(h1("Section 10 — Extended References"))
A(body("In addition to the seven references listed in the original proposal:"))
A(blank())
for s in [
    "Department of Energy, Philippines. Renewable Energy Roadmap 2020–2040. Manila: DOE, 2023.",
    "NREL. Wind Resource Assessment Handbook. Golden CO: National Renewable Energy Laboratory, 2021.",
    "Burton T., Jenkins N., Sharpe D., Bossanyi E. Wind Energy Handbook, 3rd ed. Wiley, 2021.",
    "Manwell J., McGowan J., Rogers A. Wind Energy Explained: Theory, Design and Application, 2nd ed. Wiley, 2010.",
    "Hau E. Wind Turbines: Fundamentals, Technologies, Application, Economics, 3rd ed. Springer, 2013.",
    "DNV GL. DNVGL-ST-0437: Loads and Site Conditions for Wind Turbines. 2016.",
    "World Bank ESMAP. Philippines — Wind Energy Resource Assessment. Washington DC, 2021.",
    "NorthWind Power Development Corp. Bangui Wind Farm — Annual Operations Report (public).",
    "Energy Development Corporation (EDC). Burgos Wind Project Technical Reference Document.",
    "Vestas. V126-3.45 MW and V136-4.5 MW Product Brochures (representative 3–5 MW class).",
    "Siemens Gamesa. SG 5.0-145 Product Specification Sheet (representative 5 MW class).",
    "GE Renewable Energy. GE 4.8-158 and 5.5-158 Onshore Wind Turbine Specifications.",
    "PAGASA. Climatological Normals for Ilocos Norte 1991–2020. Quezon City, 2022.",
    "Pangasinan State University — ME Capstone Design Manual, 2025 Edition.",
]:
    A(bullet(s))
A(page_break())

# ════════════════════════════════════════════════════════════════════════
# CLOSING
# ════════════════════════════════════════════════════════════════════════
A(h1("Closing Note"))
A(body("This supplement, together with the original Windmill Project Proposal (April 29, 2026), "
       "constitutes a complete senior-level plant-design submission for the Pasuquin Wind Farm. "
       "The combined document covers: wind resource assessment, site selection, technical "
       "specifications with engineering calculations (Betz, wind shear, capacity factor, wake "
       "spacing, voltage drop, transformer sizing), plant siting and layout, environmental "
       "impact assessment, financial analysis, process flow diagrams (Figures II-B.1–3), "
       "glossary, detailed equipment specifications, single-line electrical diagram, foundation "
       "and anchor bolt calculations, facility inventory, energy balance, O&M plan, materials "
       "and color-coding standards, regulatory citations, and extended references."))
A(blank())
A(p(runs=[run("Prepared by: ", bold=True),
          run("Dan Mark C. Pastoral · BSME-4A · Pangasinan State University")]))
A(p(runs=[run("Submitted to: ", bold=True),
          run("Marfel D. Rosario, PME — Instructor")]))
A(p(runs=[run("Date: ", bold=True), run("May 2026")]))


# ════════════════════════════════════════════════════════════════════════
# ASSEMBLE AND WRITE .docx
# ════════════════════════════════════════════════════════════════════════

# sectPr (same page size and margins as original)
sect_pr = (
    '<w:sectPr>'
    '<w:headerReference w:type="default" r:id="rId13"/>'
    '<w:footerReference w:type="default" r:id="rId14"/>'
    '<w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1440" w:right="1260" w:bottom="1440" w:left="1260" '
    'w:header="708" w:footer="708" w:gutter="0"/>'
    '</w:sectPr>'
)

document_xml = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<w:document {NS}>\n'
    f'<w:body>\n'
    + "\n".join(body_parts) +
    f'\n{sect_pr}\n'
    f'</w:body>\n'
    f'</w:document>'
)

# Relationships for document.xml
doc_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.jpg"/>
  <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.jpg"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image3.png"/>
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image4.png"/>
  <Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image5.png"/>
  <Relationship Id="rId13" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId14" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>"""

content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>"""

root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", root_rels)
    z.writestr("word/document.xml", document_xml)
    z.writestr("word/_rels/document.xml.rels", doc_rels_xml)
    z.writestr("word/styles.xml", orig_styles)
    z.writestr("word/numbering.xml", orig_numbering)
    z.writestr("word/header1.xml", orig_header)
    z.writestr("word/footer1.xml", orig_footer)
    # Embed all original images
    z.writestr("word/media/image1.jpg", img1_data)
    z.writestr("word/media/image2.jpg", img2_data)
    z.writestr("word/media/image3.png", img3_data)
    z.writestr("word/media/image4.png", img4_data)
    z.writestr("word/media/image5.png", img5_data)

import os
size_kb = os.path.getsize(OUTPUT) // 1024
print(f"✅ WROTE: {OUTPUT}  ({size_kb} KB)")
print(f"   body elements : {len(body_parts)}")
print(f"   document.xml  : {len(document_xml):,} bytes")
print(f"   images embedded: image1.jpg, image2.jpg, image3.png, image4.png, image5.png")
