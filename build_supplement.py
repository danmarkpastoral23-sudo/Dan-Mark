"""
Build a supplemental .docx for the Pasuquin Windmill Project containing
all the missing sections identified after comparing with John Carlo Rebalbusa's
15 MLD Wastewater Treatment Plant project.

Pure-Python: zipfile + XML only. No external deps.

Output: WINDMILL_SUPPLEMENT_Pasuquin.docx
"""

import zipfile
from xml.sax.saxutils import escape as xesc

# =====================================================================
# DOCX BUILDER (heading, paragraphs, bullets, tables)
# =====================================================================

W_NS = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""

# ---------- Static parts of the package ----------

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="en-US"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:pPr><w:jc w:val="center"/><w:spacing w:before="240" w:after="240"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:sz w:val="44"/><w:color w:val="1F3864"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:spacing w:before="360" w:after="120"/><w:keepNext/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:sz w:val="32"/><w:color w:val="1F3864"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:pPr><w:spacing w:before="280" w:after="100"/><w:keepNext/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:sz w:val="26"/><w:color w:val="2E74B5"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:pPr><w:spacing w:before="200" w:after="80"/><w:keepNext/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:sz w:val="22"/><w:color w:val="2E74B5"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet">
    <w:name w:val="List Bullet"/>
    <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:spacing w:after="60"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Note">
    <w:name w:val="Note"/>
    <w:pPr><w:ind w:left="360"/><w:spacing w:after="120"/></w:pPr>
    <w:rPr><w:i/><w:color w:val="595959"/></w:rPr>
  </w:style>
  <w:style w:type="table" w:styleId="GridTable">
    <w:name w:val="Grid Table"/>
    <w:tblPr>
      <w:tblBorders>
        <w:top w:val="single" w:sz="4" w:color="808080"/>
        <w:left w:val="single" w:sz="4" w:color="808080"/>
        <w:bottom w:val="single" w:sz="4" w:color="808080"/>
        <w:right w:val="single" w:sz="4" w:color="808080"/>
        <w:insideH w:val="single" w:sz="4" w:color="808080"/>
        <w:insideV w:val="single" w:sz="4" w:color="808080"/>
      </w:tblBorders>
    </w:tblPr>
  </w:style>
</w:styles>"""

NUMBERING = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#8226;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>"""


# ---------- Element builders ----------

def run(text, bold=False, italic=False, color=None, size=None):
    rpr = []
    if bold: rpr.append("<w:b/>")
    if italic: rpr.append("<w:i/>")
    if color: rpr.append(f'<w:color w:val="{color}"/>')
    if size: rpr.append(f'<w:sz w:val="{size}"/>')
    rpr_xml = f"<w:rPr>{''.join(rpr)}</w:rPr>" if rpr else ""
    return f'<w:r>{rpr_xml}<w:t xml:space="preserve">{xesc(text)}</w:t></w:r>'


def para(text="", style=None, runs=None, align=None):
    """Make a paragraph. Either pass plain text or a list of run XML strings."""
    ppr = []
    if style: ppr.append(f'<w:pStyle w:val="{style}"/>')
    if align: ppr.append(f'<w:jc w:val="{align}"/>')
    ppr_xml = f"<w:pPr>{''.join(ppr)}</w:pPr>" if ppr else ""
    if runs is None:
        body = run(text) if text else ""
    else:
        body = "".join(runs)
    return f"<w:p>{ppr_xml}{body}</w:p>"


def title(text):     return para(text, style="Title")
def h1(text):        return para(text, style="Heading1")
def h2(text):        return para(text, style="Heading2")
def h3(text):        return para(text, style="Heading3")


def p(text):
    return para(text)


def p_rich(*runs_):
    return para(runs=list(runs_))


def bullet(text):
    return para(text, style="ListBullet")


def note(text):
    return para(runs=[run("Note: ", bold=True, italic=True, color="595959"),
                      run(text, italic=True, color="595959")], style="Note")


def page_break():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def table(rows, widths=None, header=True):
    """rows: list of list of cells (str). First row treated as header if header=True."""
    if not rows: return ""
    n_cols = max(len(r) for r in rows)
    if widths is None:
        widths = [9000 // n_cols] * n_cols
    # Table grid
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    tbl_pr = """<w:tblPr>
      <w:tblW w:w="9000" w:type="dxa"/>
      <w:tblBorders>
        <w:top w:val="single" w:sz="6" w:color="404040"/>
        <w:left w:val="single" w:sz="6" w:color="404040"/>
        <w:bottom w:val="single" w:sz="6" w:color="404040"/>
        <w:right w:val="single" w:sz="6" w:color="404040"/>
        <w:insideH w:val="single" w:sz="4" w:color="808080"/>
        <w:insideV w:val="single" w:sz="4" w:color="808080"/>
      </w:tblBorders>
    </w:tblPr>"""

    out = ['<w:tbl>', tbl_pr, f'<w:tblGrid>{grid}</w:tblGrid>']
    for i, row in enumerate(rows):
        is_header = header and i == 0
        out.append('<w:tr>')
        for j in range(n_cols):
            cell = row[j] if j < len(row) else ""
            shading = '<w:shd w:val="clear" w:color="auto" w:fill="1F3864"/>' if is_header else ""
            tcpr = f'<w:tcPr><w:tcW w:w="{widths[j]}" w:type="dxa"/>{shading}</w:tcPr>'
            run_xml = run(str(cell), bold=is_header, color=("FFFFFF" if is_header else None))
            out.append(f'<w:tc>{tcpr}<w:p>{run_xml}</w:p></w:tc>')
        out.append('</w:tr>')
    out.append('</w:tbl>')
    # An empty paragraph is recommended after a table
    out.append("<w:p/>")
    return "".join(out)


# =====================================================================
# DOCUMENT CONTENT
# =====================================================================

body_parts = []
add = body_parts.append

# --------- Title page ---------
add(title("WINDMILL PROJECT PROPOSAL — SUPPLEMENT"))
add(para("Pasuquin, Ilocos Norte, Philippines", align="center",
         runs=[run("Pasuquin, Ilocos Norte, Philippines", bold=True, size="28")]))
add(para("Supplemental Engineering Documentation", align="center",
         runs=[run("Supplemental Engineering Documentation", italic=True, size="24", color="595959")]))
add(p(""))
add(para("Submitted by:", align="center",
         runs=[run("Submitted by:", bold=True)]))
add(para("Dan Mark C. Pastoral", align="center"))
add(para("BSME-4A", align="center"))
add(para("Pangasinan State University", align="center"))
add(p(""))
add(para("Submitted to:", align="center",
         runs=[run("Submitted to:", bold=True)]))
add(para("Marfel D. Rosario, PME", align="center"))
add(p(""))
add(para("May 2026", align="center", runs=[run("May 2026", italic=True)]))

add(page_break())

# --------- Foreword ---------
add(h1("Purpose of this Supplement"))
add(p("This supplement extends the original Windmill Project Proposal for Pasuquin, "
      "Ilocos Norte by adding sections that are conventional in industrial-plant "
      "design submissions: a glossary of acronyms, detailed equipment specification "
      "tables, a single-line electrical diagram description, foundation and anchor "
      "bolt calculations, a facility inventory with floor areas, an energy balance "
      "diagram, an Operations and Maintenance plan, and a piping/cable color-coding "
      "and materials standards section."))
add(p("The intent is to bring the document to the same level of engineering detail "
      "expected of a senior Mechanical Engineering capstone in plant design — "
      "complementing the wind-resource and financial analysis already presented in "
      "the main proposal."))

add(page_break())

# =====================================================================
# SECTION 1 — GLOSSARY
# =====================================================================
add(h1("Section 1 — Glossary of Terms and Acronyms"))
add(p("The following acronyms and technical terms are used throughout this document. "
      "Definitions are aligned with IEC 61400, the Philippine Department of Energy "
      "(DOE), and the Energy Regulatory Commission (ERC) usage."))

glossary = [
    ["Term", "Definition"],
    ["AEP", "Annual Energy Production. The total electrical energy a turbine or wind farm generates in one year, expressed in MWh."],
    ["Betz Limit", "The theoretical maximum fraction of kinetic energy in the wind that any turbine can extract, equal to 16/27 ≈ 59.3%. Derived by Albert Betz, 1919."],
    ["Capacity Factor (CF)", "Ratio of actual annual energy produced to the maximum possible if the turbine ran at rated power for all 8,760 hours of the year."],
    ["Cut-in Speed", "Lowest wind speed at which the turbine begins to produce net positive electrical power. For this project: 3.0 m/s."],
    ["Cut-out Speed", "Highest wind speed at which the turbine is allowed to operate. Above this, the rotor is feathered and braked. For this project: 25 m/s."],
    ["DFIG", "Doubly Fed Induction Generator. The variable-speed generator topology selected for this project."],
    ["DOE", "Department of Energy of the Philippines."],
    ["ECC", "Environmental Compliance Certificate, issued by DENR–EMB after EIS review."],
    ["EIS", "Environmental Impact Statement, the document prepared under the Philippine Environmental Impact Statement System (PEISS)."],
    ["ERC", "Energy Regulatory Commission, the agency that approves tariffs, including the Feed-in Tariff (FiT)."],
    ["FiT", "Feed-in Tariff. A guaranteed price per kWh paid to qualified renewable energy generators."],
    ["FPIC", "Free, Prior and Informed Consent. Required from indigenous cultural communities under the IPRA Law (RA 8371)."],
    ["HAWT", "Horizontal Axis Wind Turbine. The 3-bladed upwind configuration used in this project."],
    ["Hub Height", "Vertical distance from ground (or sea level) to the rotor hub center. For this project: 90–110 m."],
    ["IEC 61400", "International Electrotechnical Commission standard series for wind turbine design, structural loads, and grid integration."],
    ["MW / MWh", "Megawatt / Megawatt-hour. 1 MW = 1,000 kW. 1 MWh of energy = 1 MW delivered for 1 hour."],
    ["NGCP", "National Grid Corporation of the Philippines. The transmission grid operator."],
    ["NSCP", "National Structural Code of the Philippines. Used here for tower foundation design (NSCP 2015)."],
    ["O&M", "Operations and Maintenance — the on-going running of the wind farm after commissioning."],
    ["PME", "Professional Mechanical Engineer. Philippine licensure title under RA 8495 (Philippine Mechanical Engineering Act of 1998)."],
    ["P&ID", "Piping and Instrumentation Diagram. For wind farms, the equivalent is the Single-Line Diagram (SLD) — see Section 3."],
    ["RA 9513", "Renewable Energy Act of 2008. The legal basis for renewable energy incentives in the Philippines."],
    ["Rated Speed", "Wind speed at which the turbine first reaches its rated electrical output. For this project: 12–14 m/s."],
    ["Rotor Diameter", "Diameter of the circle swept by the rotor blades. For this project: 130–145 m."],
    ["SCADA", "Supervisory Control and Data Acquisition. The remote monitoring and control system."],
    ["Specific Power", "Rated capacity divided by rotor swept area, in W/m². A sizing indicator. Optimum for Class II sites: 250–400 W/m²."],
    ["Swept Area", "Area of the circle traced by the rotor blades, A = π·r². For a 130 m rotor: 13,273 m²."],
    ["Wake Effect", "Reduction in wind speed and increase in turbulence behind a turbine, which affects downstream turbines."],
    ["Weibull Distribution", "Probability distribution used to model wind speed frequency at a site. Defined by shape parameter k and scale parameter c."],
    ["Wind Power Density (WPD)", "Available wind power per unit area, ½·ρ·V³, in W/m². Determines whether a site is viable."],
    ["Wind Shear", "Increase of wind speed with altitude, modeled by the Power Law: V(z) = V_ref · (z/z_ref)^α. For coastal terrain: α ≈ 0.14."],
]
add(table(glossary, widths=[2200, 6800]))

add(page_break())

# =====================================================================
# SECTION 2 — DETAILED EQUIPMENT SPECIFICATIONS
# =====================================================================
add(h1("Section 2 — Detailed Equipment Specifications"))
add(p("Each major component of the wind farm is specified below with parameters, "
      "materials, and design basis. Quantities are based on a 20-turbine, 100 MW "
      "configuration (5.0 MW per turbine)."))

add(h2("2.1 Rotor and Blade Assembly"))
add(table([
    ["Parameter", "Specification", "Basis / Reference"],
    ["Configuration", "3-bladed upwind HAWT", "Industry standard"],
    ["Rotor diameter", "130 m (nominal)", "Class II site optimization"],
    ["Swept area", "13,273 m²", "A = π·(65)²"],
    ["Blade length", "63.5 m", "Diameter/2 minus hub radius"],
    ["Blade material", "Glass-fiber–reinforced epoxy (GFRP) with carbon spar caps", "IEC 61400-5"],
    ["Blade weight", "≈ 14,000 kg per blade", "Manufacturer benchmark (Vestas V126/V136 class)"],
    ["Tip speed (max)", "≤ 90 m/s", "Acoustic limit and structural fatigue"],
    ["Pitch system", "Individual electric pitch, 3 motors", "Class II load mitigation"],
    ["Lightning protection", "Receptors at tip + down-conductor to ground", "IEC 61400-24"],
    ["Coning angle", "2.5°", "Tower clearance under thrust"],
], widths=[2400, 3800, 2800]))

add(h2("2.2 Drivetrain (Gearbox + Generator)"))
add(table([
    ["Component", "Specification"],
    ["Gearbox type", "Three-stage planetary–helical, integrated"],
    ["Gear ratio", "1 : 90 (rotor 17 rpm → generator 1,530 rpm)"],
    ["Gearbox efficiency", "≥ 97% at rated load"],
    ["Lubrication", "Forced ISO VG 320 mineral oil, 600 L sump"],
    ["Generator type", "Doubly Fed Induction Generator (DFIG)"],
    ["Rated power", "5.0 MW"],
    ["Voltage", "690 V AC, 3-phase"],
    ["Synchronous speed", "1,500 rpm at 60 Hz"],
    ["Cooling", "Air-to-water heat exchanger"],
    ["Power converter", "Partial-scale (≈30% of rated), IGBT-based"],
    ["Total drivetrain efficiency", "≈ 93% (gearbox × generator × converter)"],
], widths=[3000, 6000]))

add(h2("2.3 Tower Structure"))
add(table([
    ["Parameter", "Specification"],
    ["Type", "Tubular steel, three-section bolted"],
    ["Hub height", "100 m (selected)"],
    ["Base diameter", "4.30 m"],
    ["Top diameter", "3.00 m"],
    ["Wall thickness", "30 mm at base, 18 mm at top"],
    ["Material", "Structural steel S355 J0+N (EN 10025-2)"],
    ["Tower mass", "≈ 280 t per turbine"],
    ["Internal access", "Service ladder + climb-assist + emergency descender"],
    ["Surface protection", "C5-M offshore coating system, total DFT ≥ 320 µm"],
    ["Aviation marking", "Red bands per CAAP MOS Part 14, top 1/3 painted"],
], widths=[3000, 6000]))

add(h2("2.4 Yaw and Pitch Systems"))
add(table([
    ["Parameter", "Specification"],
    ["Yaw drive", "6 × electric planetary gearmotors, 11 kW each"],
    ["Yaw bearing", "Internal-toothed slewing ring, 2.8 m PCD"],
    ["Yaw rate", "0.5°/s nominal"],
    ["Pitch drive", "3 × electric pitch motors, 22 kW each, with battery backup"],
    ["Pitch range", "0° (run) to 90° (full feather)"],
    ["Emergency feather time", "≤ 6 s from rated to 88°"],
], widths=[3000, 6000]))

add(h2("2.5 Electrical Collection and Substation"))
add(table([
    ["Component", "Specification"],
    ["Turbine output transformer", "5.5 MVA, 690 V / 33 kV, oil-filled, ONAN"],
    ["Inter-array cabling", "33 kV XLPE, copper, 240–630 mm² cross-section, buried 1.0 m"],
    ["Switchgear at turbine", "33 kV ring-main unit, SF6-insulated"],
    ["Collection feeders", "5 feeders of 4 turbines each (≈ 20 MW per feeder)"],
    ["Step-up transformer", "120 MVA, 33 kV / 115 kV, on-load tap-changer ±10%"],
    ["Substation protection", "Differential, distance, overcurrent, earth-fault relays per IEC 61850"],
    ["Reactive compensation", "±30 MVAr STATCOM for grid code compliance"],
    ["Earthing", "Mesh + rod, ≤ 1 Ω resistance"],
    ["SCADA", "Wind-Farm Management System with OPC-UA to NGCP RTU"],
    ["Auxiliary supply", "Diesel generator 500 kVA + UPS 30 minutes"],
], widths=[3000, 6000]))

add(page_break())

# =====================================================================
# SECTION 3 — SLD
# =====================================================================
add(h1("Section 3 — Single-Line Electrical Diagram (SLD)"))
add(p("The Single-Line Diagram (SLD) is the wind-farm equivalent of a P&ID. "
      "It traces the path of electrical power from each turbine generator to "
      "the NGCP grid connection point, identifying every transformer, breaker, "
      "metering point, and protection relay along the way."))

add(h2("3.1 Power Flow — Turbine to Grid"))
add(p("The diagram below describes the power flow in narrative form (a graphic "
      "version is to be included as Figure 3.1)."))
add(bullet("Stage 1 — Generation: 5.0 MW DFIG produces 690 V, 3-phase AC at variable frequency, "
           "stabilized to 60 Hz by the partial-scale power converter inside the nacelle."))
add(bullet("Stage 2 — Turbine step-up: 5.5 MVA pad-mounted transformer steps the voltage from "
           "690 V to 33 kV. Located either at the tower base (preferred) or inside the tower."))
add(bullet("Stage 3 — Inter-array collection: 33 kV XLPE underground cables connect 4 turbines "
           "in series along each feeder, with a ring-main switchgear at every turbine for "
           "isolation."))
add(bullet("Stage 4 — Substation: 5 feeders converge at the on-site 33/115 kV substation. "
           "Each feeder has its own circuit breaker, current and voltage transformers, "
           "and protection relays. The 120 MVA step-up transformer raises the voltage "
           "to 115 kV for grid export."))
add(bullet("Stage 5 — Grid connection: A 115 kV overhead transmission line of approximately "
           "30–40 km routes the energy to the NGCP Laoag Substation, where it is metered "
           "by an ERC-approved revenue meter and injected into the Luzon Grid."))

add(h2("3.2 Tag Conventions"))
add(p("The following ISA-style tag prefix convention is used throughout the SLD:"))
add(table([
    ["Tag", "Meaning"],
    ["G-xxx", "Generator (e.g., G-101 = Generator of Turbine T-101)"],
    ["T-xxx", "Transformer (T-101 = pad-mount; T-200 = main step-up)"],
    ["CB-xxx", "Circuit breaker"],
    ["DS-xxx", "Disconnect switch / isolator"],
    ["CT-xxx, VT-xxx", "Current and voltage instrument transformers"],
    ["WTG-xxx", "Wind Turbine Generator unit number"],
    ["MET-xxx", "Revenue or check metering point"],
    ["PR-xxx", "Protection relay (87T differential, 21 distance, 50/51 overcurrent, etc.)"],
], widths=[2000, 7000]))

add(h2("3.3 Protection Schedule"))
add(table([
    ["Equipment", "Primary Protection", "Backup Protection"],
    ["Pad-mount transformer", "Internal pressure / Buchholz relay", "Differential 87T, overcurrent 51"],
    ["Inter-array feeder", "Distance 21, directional 67", "Time-overcurrent 51, earth-fault 51N"],
    ["Step-up 33/115 kV", "Differential 87T, REF 64", "Overcurrent 51, overflux 24"],
    ["115 kV transmission", "Distance 21 (zone 1, 2, 3)", "Pilot communication scheme"],
    ["Busbars", "Bus differential 87B", "Reverse-blocking overcurrent"],
], widths=[2400, 3300, 3300]))

add(page_break())

# =====================================================================
# SECTION 4 — FOUNDATION AND ANCHOR BOLT CALCULATIONS
# =====================================================================
add(h1("Section 4 — Foundation and Anchor Bolt Calculations"))
add(p("This section sizes the gravity foundation and anchor bolt cage that fixes "
      "each tower to its concrete base. Loads are taken from IEC 61400-1 Class II "
      "and verified against NSCP 2015."))

add(h2("4.1 Design Loads (per turbine, at tower base)"))
add(table([
    ["Load Case", "Symbol", "Value", "Source"],
    ["Vertical dead load (rotor + nacelle + tower)", "F_z", "4,200 kN", "Manufacturer load document"],
    ["Maximum thrust (extreme operating)", "F_y", "1,050 kN", "IEC 61400-1 DLC 1.3 / 6.2"],
    ["Resulting overturning moment at base", "M_y", "115,000 kN·m", "F_y × hub height (≈ 100 m + lever arm)"],
    ["Torsional moment", "M_z", "8,500 kN·m", "Yaw misalignment + extreme yaw"],
], widths=[3500, 1200, 2300, 2000]))

add(h2("4.2 Gravity Foundation — Octagonal Mat"))
add(p("A cast-in-place reinforced concrete octagonal slab is selected. Sizing follows "
      "the standard formula for resisting overturning by self-weight."))
add(table([
    ["Parameter", "Value", "Note"],
    ["Outer diameter (across flats)", "21.0 m", "Trial size"],
    ["Slab thickness (edge / center)", "1.0 m / 3.0 m", "Tapered for economy"],
    ["Concrete grade", "fc' = 35 MPa", "ACI 318 / NSCP"],
    ["Reinforcement", "Grade 60 (fy = 414 MPa)", "Top + bottom mats, both directions"],
    ["Concrete volume", "≈ 590 m³", "Octagonal frustum"],
    ["Foundation self-weight (W_f)", "14,160 kN", "γ_c = 24 kN/m³"],
    ["Soil cover (1.0 m × 380 m²)", "6,840 kN", "γ_s = 18 kN/m³"],
    ["Total stabilizing weight (W_total)", "25,200 kN", "F_z + W_f + W_soil"],
    ["Lever arm to edge (e)", "10.5 m", "= D/2"],
    ["Stabilizing moment (M_s)", "264,600 kN·m", "= W_total × e"],
    ["Safety factor against overturning", "M_s / M_y ≈ 2.30", "Required ≥ 1.5 → OK"],
    ["Bearing pressure (max edge)", "≈ 175 kPa", "Trapezoidal distribution"],
    ["Allowable soil bearing", "≥ 250 kPa", "Geotech recommendation, to be confirmed"],
], widths=[3500, 2700, 2800]))

add(h2("4.3 Anchor Bolt Cage"))
add(p("Anchor bolts transfer the bending moment between the steel tower flange and the "
      "concrete foundation. They are arranged in a double-ring (inner and outer) cage."))
add(table([
    ["Parameter", "Value", "Basis"],
    ["Number of bolts", "144 (72 inner + 72 outer)", "Industry standard for 5 MW class"],
    ["Bolt grade", "ISO 898-1, Class 10.9 (fy ≈ 900 MPa)", ""],
    ["Bolt diameter", "M48 (effective area Aₛ = 1,470 mm²)", ""],
    ["Bolt length (embedded + flange)", "5,500 mm", "Full anchorage depth"],
    ["Pitch circle diameter", "Inner 4.20 m, Outer 4.90 m", "Matches tower base flange"],
    ["Tensile capacity per bolt (Pt)", "Aₛ × 0.7 × fyb = 925 kN", "Allowable, with FoS 1.5"],
    ["Maximum bolt tension (T_max)", "= M_y × c / I_b ≈ 605 kN", "Linear bolt-group analysis"],
    ["Utilization", "T_max / Pt ≈ 0.65 (65%)", "Required ≤ 0.85 → OK"],
    ["Pretension torque", "≈ 4,200 N·m using hydraulic tensioner", "70% of yield, IEC 61400-6"],
    ["Bolt fatigue check", "S–N curve ISO 1099, Class 71", "DLC 1.2 (normal operation)"],
], widths=[3500, 3200, 2300]))
add(note("All values above are preliminary, intended to demonstrate the calculation "
         "method. Final foundation and anchor design must be confirmed by a "
         "Philippine-licensed Civil/Structural Engineer based on geotechnical "
         "investigation and the actual turbine load document from the supplier."))

add(page_break())

# =====================================================================
# SECTION 5 — FACILITY INVENTORY
# =====================================================================
add(h1("Section 5 — Facility Inventory and Floor Area"))
add(p("This inventory lists every facility on the wind-farm site, its purpose, and "
      "the indicative floor area required. Areas marked with an asterisk (*) are "
      "footprint only; usable land between turbines remains available for grazing "
      "or low crops, consistent with co-located land use."))

add(table([
    ["No.", "Facility", "Floor Area (m²)", "Function"],
    ["1", "Turbine pads (20 × 380 m²)", "7,600 *", "Foundation footprints"],
    ["2", "On-site 33/115 kV substation", "5,000", "Step-up, switchgear, control building"],
    ["3", "Control room / SCADA center", "180", "24/7 supervisory operations"],
    ["4", "Maintenance workshop and stores", "650", "Spare parts, tools, oil storage"],
    ["5", "Administrative building", "320", "Project office and visitor reception"],
    ["6", "Eco-tourism visitor center", "400", "Public viewing, exhibits, restrooms"],
    ["7", "Emergency and fire station", "240", "Nacelle fire response, rescue gear"],
    ["8", "Diesel genset house and fuel store", "120", "500 kVA backup + 5,000 L diesel tank"],
    ["9", "Meteorological mast pad", "100", "Permanent met mast and instrumentation"],
    ["10", "Guard house and main gate", "60", "Security checkpoint"],
    ["11", "Parking and turnaround", "1,200", "Cars, service vans, crane setup"],
    ["12", "Internal access roads", "≈ 7 km × 6 m wide ≈ 42,000", "Heavy-haul roads to each turbine"],
    ["13", "Material laydown / crane hard-stand", "2,500", "Temporary during construction"],
    ["14", "Septic and sewage treatment", "150", "On-site sanitary"],
    ["15", "Water tank and pump house", "80", "Blade washing, fire reserve"],
    ["", "TOTAL BUILT-UP FACILITIES (excl. roads/turbines)", "≈ 11,000", ""],
    ["", "TOTAL ROADS AND HARDSTANDS", "≈ 44,500", ""],
    ["", "TOTAL TURBINE FOOTPRINT", "≈ 7,600", ""],
    ["", "GROSS WIND-FARM AREA (between turbines)", "≈ 9.46 km² (most usable)", ""],
], widths=[700, 3800, 2500, 2000]))

add(page_break())

# =====================================================================
# SECTION 6 — ENERGY BALANCE DIAGRAM
# =====================================================================
add(h1("Section 6 — Energy Balance Diagram"))
add(p("The energy balance traces 1.0 MW of incident wind power through the conversion "
      "chain to delivered grid energy, identifying every loss mechanism."))

add(h2("6.1 Stage-by-Stage Energy Balance (per turbine, at rated wind)"))
add(table([
    ["Stage", "Input (MW)", "Loss (MW)", "Output (MW)", "Mechanism"],
    ["Wind kinetic energy in rotor disc", "13.5", "—", "13.5", "P_wind = ½·ρ·A·V³ at V = 12 m/s"],
    ["Aerodynamic conversion (Cp = 0.45)", "13.5", "7.4", "6.1", "Below Betz limit; real-rotor losses"],
    ["Mechanical (hub + main bearing)", "6.1", "0.06", "6.04", "≈ 1% friction"],
    ["Gearbox (3-stage planetary)", "6.04", "0.18", "5.86", "≈ 3% (97% efficient)"],
    ["Generator (DFIG)", "5.86", "0.18", "5.68", "≈ 3% (97% efficient)"],
    ["Power converter (partial-scale)", "5.68", "0.10", "5.58", "≈ 1.7% on partial load through converter"],
    ["Pad-mount transformer 690 V / 33 kV", "5.58", "0.05", "5.53", "≈ 0.9% no-load + load losses"],
    ["Inter-array cable (avg.)", "5.53", "0.04", "5.49", "≈ 0.7% I²R losses"],
    ["Step-up 33 / 115 kV transformer", "5.49", "0.03", "5.46", "≈ 0.5%"],
    ["115 kV line to Laoag (≈ 35 km)", "5.46", "0.05", "5.41", "≈ 0.9% line losses"],
    ["Net at grid metering point", "—", "—", "5.41", "Useful electrical energy delivered"],
], widths=[3200, 1200, 1200, 1200, 2200]))

add(h2("6.2 Capacity-Factor-Adjusted Annual Yield"))
add(p("Translating the rated power balance into annual energy:"))
add(bullet("Theoretical: 5.0 MW × 8,760 h = 43,800 MWh/year per turbine"))
add(bullet("With 38% capacity factor: 16,644 MWh/year per turbine"))
add(bullet("With grid-side losses (cumulative ≈ 7%): 15,479 MWh/year delivered"))
add(bullet("Wind-farm total (20 turbines): ≈ 309,580 MWh/year delivered to NGCP"))

add(page_break())

# =====================================================================
# SECTION 7 — O&M PLAN
# =====================================================================
add(h1("Section 7 — Operations and Maintenance (O&M) Plan"))
add(p("A rigorous O&M plan is essential to achieve the assumed 38% capacity factor "
      "across the 25-year design life. The plan combines preventive maintenance, "
      "condition-based monitoring, and corrective response."))

add(h2("7.1 Staffing"))
add(table([
    ["Role", "Headcount", "Shift"],
    ["Plant manager", "1", "Day"],
    ["Lead turbine technicians (Level 3)", "4", "Day, on-call"],
    ["Turbine technicians (Level 2)", "8", "Day shift, 4-on/4-off rotation"],
    ["Electrical / substation engineer", "2", "Day"],
    ["SCADA operators", "4", "24/7, 4-shift rotation"],
    ["Safety / GWO trainer", "1", "Day"],
    ["Admin and procurement", "2", "Day"],
    ["Security personnel", "6", "24/7"],
    ["TOTAL permanent staff", "28", ""],
], widths=[3500, 1500, 4000]))

add(h2("7.2 Maintenance Schedule (per turbine)"))
add(table([
    ["Frequency", "Activity"],
    ["Continuous (24/7)", "SCADA monitoring of vibration, temperature, power, alarms"],
    ["Weekly", "Visual inspection from ground; oil-leak check via drone or tower-top camera"],
    ["Monthly", "Tower internal climb: bolt torque sample, lubrication check, brake pad"],
    ["Quarterly", "Gearbox oil sample for spectroscopy; blade leading-edge inspection"],
    ["Semi-annual", "Yaw and pitch system function test; emergency descender certification"],
    ["Annual", "Full preventive maintenance: oil change, filter change, bolt re-torque, blade root bolts"],
    ["Every 2 years", "Generator slip-ring servicing; grid-code compliance test"],
    ["Every 5 years", "Major inspection: gearbox borescope, blade composite scan, transformer oil DGA"],
    ["Every 10 years", "Mid-life overhaul: planned gearbox refurbishment, generator rewind option"],
], widths=[2200, 6800]))

add(h2("7.3 Spare Parts Strategy"))
add(p("A tiered spare parts inventory is maintained on-site to minimize downtime:"))
add(bullet("Tier A (held on site, immediate use): pitch motors, yaw motors, fuses, contactors, sensors, lubricants, filters."))
add(bullet("Tier B (held in regional depot, 24-hour delivery): converter modules, slip rings, hydraulic accumulators, brake pads."))
add(bullet("Tier C (manufacturer warehouse, 2–6 weeks): main bearing, gearbox sub-assembly, generator, blade segments."))

add(h2("7.4 Safety Compliance"))
add(bullet("All technicians must hold a valid GWO Basic Safety Training (BST) certificate."))
add(bullet("Working-at-height permits issued daily; two-person rule above 30 m."))
add(bullet("Hot-work permits required inside the nacelle."))
add(bullet("Annual emergency drill: tower rescue, fire suppression, medical evacuation."))
add(bullet("All work done in compliance with DOLE D.O. 198-18 (OSH Standards)."))

add(page_break())

# =====================================================================
# SECTION 8 — MATERIALS AND COLOR CODING
# =====================================================================
add(h1("Section 8 — Materials and Color-Coding Standards"))
add(p("Standardizing materials and visual identification across the wind farm "
      "improves inspection efficiency, safety, and long-term maintainability."))

add(h2("8.1 Material Standards"))
add(table([
    ["Item", "Material / Standard"],
    ["Tower steel", "S355 J0+N per EN 10025-2"],
    ["Tower flange bolts", "ISO 898-1 Class 10.9 with hot-dip galvanized finish"],
    ["Foundation concrete", "fc' = 35 MPa, ACI 318 / NSCP 2015"],
    ["Foundation rebar", "Grade 60 (fy = 414 MPa) deformed bars"],
    ["External coating", "C5-M per ISO 12944, total DFT ≥ 320 µm (zinc primer + epoxy + PU top)"],
    ["Anchor bolts", "ASTM A615 Gr 75 or ISO 898-1 Class 10.9"],
    ["33 kV cables", "XLPE-insulated, copper conductor, Class 2, IEC 60502-2"],
    ["115 kV overhead conductor", "ACSR Drake, NESC compliant"],
    ["Substation grounding mat", "Bare copper 95 mm² per IEEE 80"],
    ["Control cabinet enclosures", "IP66 painted steel, RAL 7035 (light grey)"],
    ["Lubricants — gearbox", "ISO VG 320 mineral or PAO synthetic"],
    ["Lubricants — yaw / pitch grease", "NLGI-2 lithium-complex, EP additive"],
], widths=[3000, 6000]))

add(h2("8.2 Color-Coding Scheme"))
add(table([
    ["Item", "Color", "Standard"],
    ["Tower exterior, lower 2/3", "RAL 9010 (pure white)", "IEC / aviation standard base"],
    ["Tower exterior, upper 1/3", "RAL 3020 (traffic red), 2 bands", "CAAP MOS Part 14 — daytime obstruction marking"],
    ["Blade tip (last 6 m)", "RAL 3020 alternating with RAL 9010", "Aviation visibility"],
    ["Aviation warning lights", "Red, medium-intensity flashing", "ICAO Annex 14, CAAP rules"],
    ["Earthing conductors", "Green/yellow striped insulation", "IEC 60446"],
    ["33 kV cable phase identification", "Red (L1), Yellow (L2), Blue (L3)", "Philippine standard"],
    ["Fire-fighting equipment", "RAL 3000 (flame red)", "OSH"],
    ["Emergency exits and pathways", "RAL 6024 (traffic green) signage with white text", "ISO 7010"],
    ["Control panels (115 kV)", "RAL 7032 (pebble grey)", "Substation convention"],
    ["Caution / hazard markings", "RAL 1023 (traffic yellow) + black diagonal stripes", "ISO 3864"],
], widths=[2700, 3300, 3000]))

add(page_break())

# =====================================================================
# SECTION 9 — REGULATORY CITATIONS
# =====================================================================
add(h1("Section 9 — Regulatory and Standards Citations"))
add(p("Compliance documents that govern the design, construction, and operation "
      "of the Pasuquin Wind Farm:"))
add(bullet("RA 9513 — Renewable Energy Act of 2008, with its Implementing Rules and Regulations (DOE Department Circular DC2009-07-0011)."))
add(bullet("RA 9136 — Electric Power Industry Reform Act (EPIRA), 2001."))
add(bullet("ERC Resolution No. 16, Series of 2022 — Feed-in Tariff Rates for renewable energy resources."))
add(bullet("DENR Administrative Order DAO 2003-30 — Implementing Rules of the Philippine EIS System."))
add(bullet("RA 8371 — Indigenous Peoples' Rights Act (IPRA); FPIC required for ancestral domain areas."))
add(bullet("IEC 61400-1 (Ed. 4, 2019) — Wind Turbines, Part 1: Design Requirements."))
add(bullet("IEC 61400-3-1 — Design Requirements for Fixed Offshore Wind Turbines (referenced for coastal exposure)."))
add(bullet("IEC 61400-6 — Tower and Foundation Design Requirements."))
add(bullet("IEC 61400-12-1 — Power Performance Measurements."))
add(bullet("IEC 61400-24 — Lightning Protection."))
add(bullet("IEC 61850 — Communication Networks and Systems for Power Utility Automation (substation SCADA)."))
add(bullet("Philippine Electrical Code (PEC), 2017 Edition."))
add(bullet("Philippine Grid Code, latest edition issued by ERC."))
add(bullet("National Structural Code of the Philippines (NSCP), 2015 Edition — for tower foundation design loads."))
add(bullet("NSCP 2015 Section 207 — Wind Loads."))
add(bullet("DOLE Department Order No. 198-18 — Occupational Safety and Health Standards."))
add(bullet("CAAP Manual of Standards Part 14 — Aerodromes and Obstruction Marking."))
add(bullet("ISO 12944 — Corrosion protection by paint systems."))
add(bullet("ISO 7010 — Graphical symbols, safety signs."))
add(bullet("GWO Basic Safety Training (BST) — Working at Heights, First Aid, Manual Handling, Fire Awareness, Sea Survival."))

add(page_break())

# =====================================================================
# SECTION 10 — EXTENDED REFERENCES (URL-style)
# =====================================================================
add(h1("Section 10 — Extended References"))
add(p("In addition to the references listed in the main proposal:"))
add(bullet("Department of Energy, Philippines. Renewable Energy Roadmap 2020–2040. Manila: DOE, 2023."))
add(bullet("National Renewable Energy Laboratory (NREL). Wind Resource Assessment Handbook. Golden, CO: NREL, 2021."))
add(bullet("Burton, T., Jenkins, N., Sharpe, D., and Bossanyi, E. Wind Energy Handbook, 3rd Edition. Wiley, 2021."))
add(bullet("Manwell, J., McGowan, J., and Rogers, A. Wind Energy Explained: Theory, Design and Application, 2nd Edition. Wiley, 2010."))
add(bullet("Hau, E. Wind Turbines: Fundamentals, Technologies, Application, Economics, 3rd Edition. Springer, 2013."))
add(bullet("World Bank ESMAP. Philippines — Wind Energy Resource Assessment. Washington, DC: ESMAP Technical Report, 2021."))
add(bullet("DNV GL. Recommended Practice DNVGL-RP-0286 — Coupled Analysis of Floating Wind Turbines (selected sections for tower modal analysis)."))
add(bullet("NorthWind Power Development Corp. Bangui Wind Farm Operations Report (annual)."))
add(bullet("EDC. Burgos Wind Project Technical Reference Document."))
add(bullet("Vestas V126-3.45 MW and V136-4.2 MW Technical Brochures (representative HAWT class)."))
add(bullet("Siemens Gamesa SG 5.0-145 Product Specification Sheet (representative 5 MW class)."))
add(bullet("PAGASA. Climatological Normals for Ilocos Norte 1991–2020. Quezon City, 2022."))
add(bullet("Pangasinan State University, College of Engineering and Architecture, Mechanical Engineering Department — ME Capstone Manual, 2025."))

add(page_break())

# =====================================================================
# CLOSING
# =====================================================================
add(h1("Closing Note"))
add(p("This supplement is intended to complement, not replace, the primary Windmill "
      "Project Proposal previously submitted on 29 April 2026. Together, the two "
      "documents constitute a complete plant-design submission covering: site "
      "selection, technical specifications with engineering calculations, plant "
      "siting and layout, environmental impact, financial analysis, glossary, "
      "detailed equipment specifications, single-line diagram, foundation and anchor "
      "bolt calculations, facility inventory, energy balance, operations and "
      "maintenance plan, materials and color-coding standards, and regulatory "
      "citations."))
add(p("Prepared by: Dan Mark C. Pastoral · BSME-4A · Pangasinan State University"))
add(p("For: Engr. Marfel D. Rosario, PME · May 2026"))


# =====================================================================
# ASSEMBLE document.xml
# =====================================================================

document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {W_NS}>
  <w:body>
    {''.join(body_parts)}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
      <w:cols w:space="708"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>"""

# =====================================================================
# WRITE THE .docx (a zip archive)
# =====================================================================

OUTPUT = "WINDMILL_SUPPLEMENT_Pasuquin.docx"
with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CONTENT_TYPES)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("word/_rels/document.xml.rels", DOC_RELS)
    z.writestr("word/styles.xml", STYLES)
    z.writestr("word/numbering.xml", NUMBERING)
    z.writestr("word/document.xml", document_xml)

print(f"WROTE: {OUTPUT}")
print(f"  body paragraphs: {len(body_parts)}")
print(f"  document.xml size: {len(document_xml):,} bytes")
