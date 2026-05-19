"""
Build WINDMILL_COMPLETE_Pasuquin.docx
- Takes EXACT content from original docx (all sections I-VII + figures + tables)
- Appends the supplemental sections (VIII-A through VIII-H) right after
- One seamless complete document
"""
import zipfile, re, os
from xml.sax.saxutils import escape as xe

ORIGINAL = "WINDMILL_PROJECT_Pasuquin_FINAL (1).docx"
OUTPUT   = "WINDMILL_COMPLETE_Pasuquin.docx"

# ── Read ALL original assets ──────────────────────────────────────────
with zipfile.ZipFile(ORIGINAL) as z:
    orig_doc_xml   = z.read("word/document.xml").decode("utf-8")
    orig_rels_xml  = z.read("word/_rels/document.xml.rels").decode("utf-8")
    orig_styles    = z.read("word/styles.xml")
    orig_numbering = z.read("word/numbering.xml")
    orig_header    = z.read("word/header1.xml")
    orig_footer    = z.read("word/footer1.xml")
    orig_settings  = z.read("word/settings.xml")
    # All media files
    media = {}
    for name in z.namelist():
        if "media/" in name:
            media[name] = z.read(name)


# ── Extract original body content (everything BEFORE sectPr) ─────────
# Get the namespace declarations from original document
ns_match = re.search(r'<w:document([^>]+)>', orig_doc_xml)
ORIG_NS = ns_match.group(1) if ns_match else ''

# Get full body content
body_match = re.search(r'<w:body>(.*)</w:body>', orig_doc_xml, re.DOTALL)
original_body = body_match.group(1)

# Remove the trailing sectPr — we will add it back at the very end
# (after our new sections)
original_body_no_sect = re.sub(r'\s*<w:sectPr[^>]*>.*?</w:sectPr>\s*$',
                                '', original_body, flags=re.DOTALL)

# The original sectPr (we'll use it at the end)
ORIG_SECT_PR = '<w:sectPr w:rsidR="002047EA"><w:headerReference w:type="default" r:id="rId13"/><w:footerReference w:type="default" r:id="rId14"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1260" w:bottom="1440" w:left="1260" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>'

# ── XML helpers for new supplemental sections ─────────────────────────
def xe_s(text): return xe(str(text))

def p(text="", style="Normal", bold=False, italic=False,
      align=None, sz=None, color=None, runs=None):
    ppr = f'<w:pStyle w:val="{style}"/>'
    if align: ppr += f'<w:jc w:val="{align}"/>'
    if runs is not None:
        body = "".join(runs)
    else:
        rp = ""
        if bold:   rp += "<w:b/><w:bCs/>"
        if italic: rp += "<w:i/>"
        if sz:     rp += f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
        if color:  rp += f'<w:color w:val="{color}"/>'
        rpr = f"<w:rPr>{rp}</w:rPr>" if rp else ""
        body = f'<w:r>{rpr}<w:t xml:space="preserve">{xe_s(text)}</w:t></w:r>' if text else ""
    return f"<w:p><w:pPr>{ppr}</w:pPr>{body}</w:p>"

def run(text, bold=False, italic=False, sz=None, color=None):
    rp = ""
    if bold:   rp += "<w:b/><w:bCs/>"
    if italic: rp += "<w:i/>"
    if sz:     rp += f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
    if color:  rp += f'<w:color w:val="{color}"/>'
    rpr = f"<w:rPr>{rp}</w:rPr>" if rp else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{xe_s(text)}</w:t></w:r>'

def h1(t): return p(t, style="Heading1")
def h2(t): return p(t, style="Heading2")
def body_p(t, sz="20"): return p(t, sz=sz)
def blank(): return p("")
def pb(): return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

def bullet(text):
    return (
        '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>'
        '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
        f'<w:r><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
        f'<w:t xml:space="preserve">{xe_s(text)}</w:t></w:r></w:p>'
    )

def note(text):
    return (
        f'<w:p><w:r><w:rPr><w:i/><w:sz w:val="18"/><w:szCs w:val="18"/>'
        f'<w:color w:val="595959"/></w:rPr>'
        f'<w:t xml:space="preserve">{xe_s(text)}</w:t></w:r></w:p>'
    )

def tbl(rows, widths=None, header=True):
    if not rows: return ""
    nc = max(len(r) for r in rows)
    total_w = 9360
    if widths is None: widths = [total_w // nc] * nc
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    tbl_pr = (
        f'<w:tblPr><w:tblW w:w="{total_w}" w:type="dxa"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '</w:tblBorders>'
        '<w:tblCellMar><w:left w:w="10" w:type="dxa"/><w:right w:w="10" w:type="dxa"/></w:tblCellMar>'
        '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" '
        'w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/>'
        '</w:tblPr>'
    )
    out = [f"<w:tbl>{tbl_pr}<w:tblGrid>{grid}</w:tblGrid>"]
    for ri, row in enumerate(rows):
        is_hdr = header and ri == 0
        even = (ri % 2 == 0) and not is_hdr
        bg = "1B6B3A" if is_hdr else ("F1F8E9" if even else "FFFFFF")
        bc = "1B6B3A" if is_hdr else "CCCCCC"
        sz_str = "4" if is_hdr else "1"
        mar = ('<w:tcMar><w:top w:w="100" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>'
               '<w:bottom w:w="100" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>'
               if is_hdr else
               '<w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>'
               '<w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>')
        out.append("<w:tr>")
        for ci in range(nc):
            cell = str(row[ci]) if ci < len(row) else ""
            tcpr = (f'<w:tcPr><w:tcW w:w="{widths[ci]}" w:type="dxa"/>'
                    f'<w:tcBorders>'
                    f'<w:top w:val="single" w:sz="{sz_str}" w:space="0" w:color="{bc}"/>'
                    f'<w:left w:val="single" w:sz="{sz_str}" w:space="0" w:color="{bc}"/>'
                    f'<w:bottom w:val="single" w:sz="{sz_str}" w:space="0" w:color="{bc}"/>'
                    f'<w:right w:val="single" w:sz="{sz_str}" w:space="0" w:color="{bc}"/>'
                    f'</w:tcBorders>'
                    f'<w:shd w:val="clear" w:color="auto" w:fill="{bg}"/>{mar}</w:tcPr>')
            if is_hdr:
                cc = (f'<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
                      f'<w:r><w:rPr><w:b/><w:bCs/><w:color w:val="FFFFFF"/>'
                      f'<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
                      f'<w:t>{xe_s(cell)}</w:t></w:r></w:p>')
            else:
                cc = (f'<w:p><w:r><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
                      f'<w:t xml:space="preserve">{xe_s(cell)}</w:t></w:r></w:p>')
            out.append(f"<w:tc>{tcpr}{cc}</w:tc>")
        out.append("</w:tr>")
    out.append("</w:tbl><w:p/>")
    return "".join(out)

def img_placeholder(label, width_cm=16, height_cm=11):
    return (
        '<w:p><w:pPr><w:jc w:val="center"/>'
        '<w:pBdr>'
        '<w:top w:val="single" w:sz="6" w:space="1" w:color="888888"/>'
        '<w:left w:val="single" w:sz="6" w:space="4" w:color="888888"/>'
        '<w:bottom w:val="single" w:sz="6" w:space="1" w:color="888888"/>'
        '<w:right w:val="single" w:sz="6" w:space="4" w:color="888888"/>'
        '</w:pBdr>'
        '<w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/></w:pPr>'
        f'<w:r><w:rPr><w:color w:val="888888"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
        f'<w:t>[ {xe_s(label)} ]</w:t></w:r></w:p>'
        f'<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
        f'<w:r><w:rPr><w:i/><w:color w:val="888888"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
        f'<w:t>Insert image here — approx. {width_cm} cm × {height_cm} cm</w:t></w:r></w:p>'
    )

def caption(text):
    return (
        f'<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
        f'<w:r><w:rPr><w:i/><w:sz w:val="18"/><w:szCs w:val="18"/>'
        f'<w:color w:val="595959"/></w:rPr>'
        f'<w:t>{xe_s(text)}</w:t></w:r></w:p>'
    )


# ════════════════════════════════════════════════════════════════════
# NEW SUPPLEMENTAL SECTIONS (appended after original content)
# ════════════════════════════════════════════════════════════════════
S = []  # supplemental body parts
A = S.append

# Page break before new sections
A(pb())

# ── Section VIII-A: Glossary ─────────────────────────────────────────
A(h1("Section VIII-A — Glossary of Terms and Acronyms"))
A(body_p("The following technical terms and acronyms appear throughout this proposal. "
         "Definitions are aligned with IEC 61400, Philippine DOE, and ERC usage."))
A(blank())
A(tbl([
    ["Term / Acronym", "Definition"],
    ["AEP", "Annual Energy Production — total electrical energy generated per year (MWh/yr)."],
    ["Amihan", "Northeast Monsoon (Oct–Mar); primary wind driver for Pasuquin. Luzon Strait accelerates these winds."],
    ["Betz Limit", "Theoretical maximum wind energy extractable by any turbine = 59.3% (Albert Betz, 1919)."],
    ["Capacity Factor (CF)", "Actual AEP ÷ (Rated Power × 8,760 h/yr). Pasuquin estimate: 38%."],
    ["CAAP", "Civil Aviation Authority of the Philippines — governs aviation obstruction marking on towers."],
    ["Cut-in Speed", "Minimum wind speed for net positive power output. This project: 3.0 m/s."],
    ["Cut-out Speed", "Maximum operating wind speed; above this the rotor feathers. This project: 25 m/s."],
    ["DENR", "Department of Environment and Natural Resources — issues Environmental Compliance Certificate (ECC)."],
    ["DFIG", "Doubly Fed Induction Generator — variable-speed generator topology used in this project."],
    ["DOE", "Department of Energy of the Philippines."],
    ["ECC", "Environmental Compliance Certificate issued by DENR-EMB after EIS review."],
    ["EIS", "Environmental Impact Statement — required under the Philippine EIS System (PEISS)."],
    ["EPC", "Engineering, Procurement and Construction — main project delivery contract mode."],
    ["ERC", "Energy Regulatory Commission — approves FiT rates, grid codes, and metering standards."],
    ["FiT", "Feed-in Tariff — guaranteed PHP 8.53/kWh for wind energy generators (ERC Res. 16, 2022)."],
    ["FPIC", "Free, Prior and Informed Consent — required under RA 8371 (IPRA)."],
    ["HAWT", "Horizontal Axis Wind Turbine — 3-bladed upwind configuration; industry standard."],
    ["Hub Height", "Vertical distance from ground to rotor hub center. This project: 90–110 m."],
    ["IEC 61400", "IEC standard series for wind turbine design, structural loads, and grid integration."],
    ["MW / MWh", "Megawatt (power) / Megawatt-hour (energy). 1 MWh = 1 MW delivered for 1 hour."],
    ["NGCP", "National Grid Corporation of the Philippines — operates the Luzon Grid."],
    ["NSCP", "National Structural Code of the Philippines (2015) — governs tower foundation design."],
    ["O&M", "Operations and Maintenance — ongoing running of the wind farm after commissioning."],
    ["RA 9513", "Renewable Energy Act of 2008 — legal basis for FiT incentives and DOE service contracts."],
    ["Rated Speed", "Wind speed at which turbine first reaches rated output. This project: 12–14 m/s."],
    ["SCADA", "Supervisory Control and Data Acquisition — remote turbine monitoring and control system."],
    ["Specific Power", "Rated capacity ÷ swept area (W/m²). Optimum for IEC Class II: 250–400 W/m²."],
    ["Swept Area", "Circle traced by rotor blades, A = πr². For 130 m rotor: 13,273 m²."],
    ["Wake Effect", "Reduced wind speed and increased turbulence behind an operating turbine."],
    ["Weibull Dist.", "Statistical model for wind speed frequency; defined by shape k and scale c."],
    ["Wind Class II", "IEC classification for sites with average wind 7.5–8.5 m/s — Pasuquin's class."],
    ["Wind Power Density", "Available wind power per unit area: WPD = ½ρV³ (W/m²)."],
    ["Wind Shear", "Wind speed increase with altitude: V(z)=V_ref·(z/z_ref)^α; α≈0.14 coastal."],
], widths=[2600, 6760]))
A(pb())


# ── Section VIII-B: Equipment Specs ──────────────────────────────────
A(h1("Section VIII-B — Detailed Equipment Specifications"))
A(body_p("Specifications are based on the 5.0 MW HAWT, 130 m rotor, 100 m hub height, "
         "and 33/115 kV electrical system established in Section III "
         "(20-turbine × 5 MW = 100 MW configuration)."))
A(blank())
A(h2("B.1 Rotor and Blade Assembly"))
A(tbl([
    ["Parameter", "Specification", "Design Basis"],
    ["Configuration", "3-bladed upwind HAWT", "Industry standard for utility-scale"],
    ["Rotor diameter", "130 m", "Specific Power 339 W/m² — optimal for V=8.5 m/s (Sec. III-A.2)"],
    ["Swept area", "13,273 m²", "A = π × (65)²"],
    ["Blade length", "63.5 m", "= Rotor radius − hub radius (1.5 m)"],
    ["Blade material", "GFRP with carbon-fiber spar caps", "IEC 61400-5; fatigue and stiffness"],
    ["Blade weight", "≈ 14,000 kg per blade", "Vestas V126/V136 class benchmark"],
    ["Tip speed (max)", "≤ 90 m/s", "Acoustic limit and structural fatigue"],
    ["Pitch system", "3 independent electric motors + battery backup", "IEC 61400-1 Class II"],
    ["Lightning protection", "Receptors at blade tip + down-conductor to ground", "IEC 61400-24"],
    ["Cut-in / Rated / Cut-out", "3.0 / 12–14 / 25 m/s", "Sec. III-A.4 Weibull analysis"],
], widths=[2600, 3800, 2960]))
A(blank())
A(h2("B.2 Drivetrain — Gearbox and DFIG Generator"))
A(tbl([
    ["Component", "Specification"],
    ["Gearbox type", "Three-stage planetary–helical, integrated with main bearing"],
    ["Gear ratio", "1:90 (rotor ≈17 rpm → generator 1,530 rpm at 60 Hz)"],
    ["Gearbox efficiency", "≥ 97% at rated load"],
    ["Gearbox lubrication", "ISO VG 320 mineral/PAO synthetic, 600 L sump, forced"],
    ["Generator type", "Doubly Fed Induction Generator (DFIG) — Sec. III-A"],
    ["Rated power", "5.0 MW at 12–14 m/s wind speed"],
    ["Stator voltage", "690 V AC, 3-phase, 60 Hz"],
    ["Generator cooling", "Air-to-water heat exchanger inside nacelle"],
    ["Power converter", "Partial-scale IGBT (~30% rated) — stabilizes output to 60 Hz"],
    ["Nacelle total mass", "≈ 320 t (nacelle + rotor hub assembly)"],
    ["Overall drivetrain η", "≈ 93% (gearbox × generator × converter combined)"],
], widths=[2600, 6760]))
A(blank())
A(h2("B.3 Tower Structure"))
A(tbl([
    ["Parameter", "Specification"],
    ["Type", "Tubular steel, three-section bolted with flanged joints"],
    ["Hub height selected", "100 m — V(100m) = 8.19 m/s, +26% gain vs. ground (Sec. III-A.3)"],
    ["Base diameter", "4.30 m"],
    ["Top diameter", "3.00 m"],
    ["Wall thickness", "30 mm at base, tapers to 18 mm at top"],
    ["Steel grade", "S355 J0+N per EN 10025-2"],
    ["Tower mass", "≈ 280 t per turbine"],
    ["Coating", "ISO 12944 C5-M: zinc primer + epoxy midcoat + PU topcoat, DFT ≥ 320 µm"],
    ["Internal access", "Fixed ladder + climb-assist device + emergency descender"],
    ["Aviation marking", "Upper 1/3: RAL 3020 red bands per CAAP MOS Part 14"],
], widths=[2600, 6760]))
A(blank())
A(h2("B.4 Electrical Collection System and Substation"))
A(tbl([
    ["Component", "Specification"],
    ["Turbine pad-mount transformer", "5.5 MVA, 690 V / 33 kV, oil-filled ONAN cooling"],
    ["Collection voltage", "33 kV — optimal for up to ~200 MW (Sec. III-C.1)"],
    ["Inter-array cable", "33 kV XLPE, copper, 240–630 mm², buried 1.0 m deep, IEC 60502-2"],
    ["Switchgear per turbine", "33 kV ring-main unit, SF6-insulated"],
    ["Collection feeders", "5 feeders × 4 turbines each ≈ 20 MW per feeder"],
    ["Step-up transformer", "120 MVA, 33 kV / 115 kV, OLTC ±10%, turns ratio 3.48:1 (Sec. III-C.2)"],
    ["Transmission voltage", "115 kV — NGCP Ilocos Norte grid standard (Sec. III-C.3)"],
    ["Substation protection", "IEC 61850 relays: 87T diff, 21 distance, 51 overcurrent, 64 REF"],
    ["Reactive compensation", "±30 MVAr STATCOM for Philippine Grid Code compliance"],
    ["Grounding system", "Mesh + driven rod earthing, target ≤ 1 Ω resistance (IEEE 80)"],
    ["SCADA system", "Wind-Farm Management System, OPC-UA protocol to NGCP RTU"],
    ["Auxiliary power", "500 kVA diesel genset + 30-minute UPS for critical loads"],
], widths=[2600, 6760]))
A(pb())


# ── Section VIII-C: Plant Layout + 3D Model ───────────────────────────
A(h1("Section VIII-C — Plant Layout and 3D Model"))
A(body_p("This section presents the wind farm site layout and a 3D visualization of the "
         "Pasuquin Wind Farm, based on the spacing calculations in Section III-B.2 "
         "and the siting factors in Section IV."))
A(blank())
A(h2("C.1 Plant Site Layout"))
A(body_p("The layout shows 20 WTGs (WTG-101 to WTG-120) across the Pasuquin coastline "
         "and ridges, with the on-site substation, access roads, and ancillary facilities. "
         "Turbines: 5 rows × 4 units, 7D (910 m) inline spacing, 4D (520 m) cross-wind "
         "spacing per Section III-B.1 wake effect calculations."))
A(blank())
A(img_placeholder("PLANT LAYOUT — Site plan: turbine positions, substation, access roads, facilities", 16, 12))
A(caption("Figure VIII-C.1 — Pasuquin Wind Farm Site Layout Plan"))
A(blank())
A(h2("C.2 3D Model / Rendering"))
A(body_p("The 3D model illustrates the wind farm along the Pasuquin coastline, showing "
         "the scale of turbines (100 m hub height, 130 m rotor) against the terrain, "
         "the substation, and the eco-tourism visitor area."))
A(blank())
A(img_placeholder("3D MODEL / RENDERING — Perspective view of Pasuquin Wind Farm along the coastline", 16, 11))
A(caption("Figure VIII-C.2 — Pasuquin Wind Farm 3D Visualization"))
A(blank())
A(note("Note: Insert finalized layout drawing and 3D rendering above. "
       "Recommended: PNG or JPG, minimum 300 DPI, landscape orientation."))
A(pb())

# ── Section VIII-D: Foundation Calcs ─────────────────────────────────
A(h1("Section VIII-D — Foundation and Anchor Bolt Calculations"))
A(body_p("Gravity foundation and anchor bolt cage sizing for each turbine tower. "
         "Loads: IEC 61400-1 Class II; foundation checked per NSCP 2015."))
A(blank())
A(h2("D.1 Design Loads at Tower Base"))
A(tbl([
    ["Load Case", "Symbol", "Value", "Source"],
    ["Vertical dead load (rotor + nacelle + tower)", "F_z", "4,200 kN", "Manufacturer load document"],
    ["Extreme thrust (DLC 1.3 / DLC 6.2)", "F_y", "1,050 kN", "IEC 61400-1 Class II"],
    ["Overturning moment at tower base", "M_y", "115,000 kN·m", "F_y × hub height (~100 m) + lever arm"],
    ["Torsional moment", "M_z", "8,500 kN·m", "Extreme yaw + yaw misalignment"],
], widths=[3600, 1200, 1800, 2760]))
A(blank())
A(h2("D.2 Octagonal Gravity Mat — Sizing"))
A(tbl([
    ["Parameter", "Value", "Note"],
    ["Outer diameter (across flats)", "21.0 m", "Trial size — checked below"],
    ["Slab thickness (edge / center)", "1.0 m / 3.0 m", "Tapered frustum for economy"],
    ["Concrete grade", "fc' = 35 MPa", "ACI 318 / NSCP 2015"],
    ["Reinforcement", "Grade 60 (fy = 414 MPa)", "Top + bottom mats, both directions"],
    ["Concrete volume", "≈ 590 m³", "Octagonal frustum formula"],
    ["Foundation self-weight (W_f)", "14,160 kN", "γ_c = 24 kN/m³"],
    ["Soil cover (1.0 m × 380 m²)", "6,840 kN", "γ_s = 18 kN/m³"],
    ["Total stabilizing weight (W_total)", "25,200 kN", "F_z + W_f + W_soil"],
    ["Lever arm to edge (e = D/2)", "10.5 m", "Half of outer diameter"],
    ["Stabilizing moment (M_s = W_total × e)", "264,600 kN·m", "Resisting overturning"],
    ["Safety factor vs. overturning (M_s / M_y)", "2.30", "Required ≥ 1.5 ✓ PASS"],
    ["Max. edge bearing pressure", "≈ 175 kPa", "Trapezoidal distribution"],
    ["Required allowable soil bearing capacity", "≥ 250 kPa", "Geotech investigation required"],
], widths=[4000, 2600, 2760]))
A(blank())
A(h2("D.3 Anchor Bolt Cage"))
A(tbl([
    ["Parameter", "Value", "Basis"],
    ["Total number of bolts", "144 (72 inner + 72 outer ring)", "5 MW class industry standard"],
    ["Bolt grade", "ISO 898-1, Class 10.9 (fy ≈ 900 MPa)", "High-strength structural"],
    ["Bolt diameter", "M48 — effective stress area Aₛ = 1,470 mm²", ""],
    ["Embedded + flange length", "5,500 mm total", "Full anchorage depth"],
    ["Pitch circle diameter", "Inner 4.20 m / Outer 4.90 m", "Matches tower base flange"],
    ["Allowable tension per bolt (Pt)", "Aₛ × 0.7 × fy = 925 kN", "FoS = 1.5 applied"],
    ["Max. bolt tension (T_max)", "M_y × c / I_bolt-group ≈ 605 kN", "Linear bolt-group analysis"],
    ["Utilization (T_max / Pt)", "65%", "Required ≤ 85% ✓ PASS"],
    ["Pretension torque", "≈ 4,200 N·m via hydraulic tensioner", "70% of fy, per IEC 61400-6"],
    ["Fatigue check", "S-N curve ISO 1099, Class 71", "DLC 1.2 normal operation"],
], widths=[3200, 3400, 2760]))
A(note("NOTE: All values are preliminary. Final design must be certified by a Philippine-licensed "
       "Civil/Structural Engineer based on geotechnical investigation and actual turbine load report."))
A(pb())


# ── Section VIII-E: Facility Inventory ───────────────────────────────
A(h1("Section VIII-E — Facility Inventory and Floor Areas"))
A(body_p("All facilities on site based on Section IV plant layout. "
         "Land between turbines (~9.45 km²) remains usable for farming and grazing."))
A(blank())
A(tbl([
    ["No.", "Facility", "Floor Area (m²)", "Function"],
    ["1", "Wind turbine foundations (20 × ~380 m²)", "7,600 *", "Octagonal RC mat per Sec. VIII-D"],
    ["2", "On-site 33/115 kV substation", "5,000", "Step-up transformer, switchgear, control room"],
    ["3", "SCADA / Control room", "180", "24/7 supervisory wind-farm control"],
    ["4", "Maintenance workshop and stores", "650", "Spare parts, gearbox oil, blade tools"],
    ["5", "Administrative building", "320", "Project management, visitor reception"],
    ["6", "Eco-tourism visitor center", "400", "Public viewing deck, exhibits, restrooms"],
    ["7", "Emergency and fire station", "240", "Nacelle fire response, high-angle rescue"],
    ["8", "Diesel genset house + fuel storage", "120", "500 kVA backup + 5,000 L diesel tank"],
    ["9", "Meteorological mast pad", "100", "Permanent met mast and instrumentation"],
    ["10", "Guard house and main gate", "60", "24/7 security checkpoint"],
    ["11", "Parking and crane turnaround", "1,200", "Heavy vehicles, mobile crane hardstand"],
    ["12", "Material laydown / crane hardstand", "2,500", "Blade/nacelle staging (construction)"],
    ["13", "Septic and sewage treatment", "150", "On-site sanitary waste treatment"],
    ["14", "Water tank and pump house", "80", "Blade wash, fire-reserve tank (50,000 L)"],
    ["15", "Internal access roads (7 km × 6 m wide)", "≈ 42,000", "Heavy-haul gravel roads to each WTG"],
    ["", "TOTAL BUILT-UP FACILITIES (excl. roads + turbine pads)", "≈ 11,000", ""],
    ["", "TOTAL TURBINE FOUNDATION FOOTPRINT (*)", "≈ 7,600", ""],
    ["", "TOTAL ROADS AND HARDSTANDS", "≈ 44,500", ""],
    ["", "GROSS WIND FARM AREA (including inter-turbine land)", "≈ 9.46 km²", "Per Sec. III-B.2"],
], widths=[600, 3800, 2400, 2560]))
A(pb())

# ── Section VIII-F: Energy Balance ───────────────────────────────────
A(h1("Section VIII-F — Energy Balance Diagram"))
A(body_p("Extends the Process Flow Diagrams in Section II-B by quantifying power losses "
         "at each stage — from wind kinetic energy to electricity at the NGCP grid metering point."))
A(blank())
A(h2("F.1 Stage-by-Stage Power Balance (per turbine at V = 12 m/s)"))
A(body_p("Basis: P_wind = ½ × ρ × A × V³ = ½ × 1.225 × 13,273 × 12³ = 13.96 MW "
         "(same Betz equation as Section III-A.1)."))
A(blank())
A(tbl([
    ["Stage", "Input (MW)", "Loss (MW)", "Output (MW)", "Loss Mechanism"],
    ["Wind kinetic energy at rotor disc", "13.96", "—", "13.96", "P_wind = ½ρAV³"],
    ["Aerodynamic conversion (Cp = 0.45)", "13.96", "7.68", "6.28", "Real rotor < Betz 59.3%"],
    ["Hub + main bearing friction", "6.28", "0.06", "6.22", "≈ 1% mechanical losses"],
    ["Gearbox (3-stage planetary, η = 97%)", "6.22", "0.19", "6.03", "≈ 3% gear losses"],
    ["DFIG generator (η = 97%)", "6.03", "0.18", "5.85", "≈ 3% electromagnetic losses"],
    ["Partial-scale power converter", "5.85", "0.10", "5.75", "≈ 1.7% IGBT losses"],
    ["Pad-mount transformer 690 V / 33 kV", "5.75", "0.05", "5.70", "≈ 0.9% iron + copper losses"],
    ["Inter-array 33 kV cable (average)", "5.70", "0.04", "5.66", "≈ 0.7% I²R losses"],
    ["Step-up transformer 33 kV / 115 kV", "5.66", "0.03", "5.63", "≈ 0.5% transformer losses"],
    ["115 kV OHL to Laoag (≈ 35 km)", "5.63", "0.05", "5.58", "≈ 0.9% line losses"],
    ["NET at NGCP grid metering point", "—", "—", "5.58", "Useful electricity delivered"],
], widths=[3200, 1100, 1100, 1300, 2660]))
A(blank())
A(h2("F.2 Annual Energy Yield — Wind Farm Total"))
A(tbl([
    ["Parameter", "Value", "Calculation / Basis"],
    ["Rated capacity per turbine", "5.0 MW", "Design specification (Sec. III-A.1)"],
    ["Max theoretical AEP per turbine", "43,800 MWh/yr", "5.0 MW × 8,760 h/yr"],
    ["Capacity factor — Pasuquin", "38%", "Consistent with Bangui 35–42%, Burgos 36–45%"],
    ["Gross AEP per turbine", "16,644 MWh/yr", "43,800 × 0.38"],
    ["Electrical system losses (avg. 7%)", "1,165 MWh/yr", "Cables, transformers, line"],
    ["Net AEP per turbine delivered", "15,479 MWh/yr", "16,644 × (1 − 0.07)"],
    ["Wind farm total (20 turbines)", "309,580 MWh/yr", "15,479 × 20 turbines"],
    ["Equivalent households served", "≈ 77,400", "At 4,000 kWh/yr per household (PH avg.)"],
    ["Annual FiT revenue (gross)", "PHP 2.64 Billion/yr", "309,580 MWh × PHP 8.53/kWh"],
], widths=[3200, 2400, 3760]))
A(pb())


# ── Section VIII-G: O&M Plan ──────────────────────────────────────────
A(h1("Section VIII-G — Operations and Maintenance (O&M) Plan"))
A(body_p("Sustains the 38% capacity factor assumed in the financial analysis "
         "(Section VI) across the 25-year project life."))
A(blank())
A(h2("G.1 Permanent Staffing"))
A(tbl([
    ["Role", "No.", "Shift Schedule"],
    ["Wind Farm Manager", "1", "Day"],
    ["Lead Turbine Technician (GWO Level 3)", "4", "Day + on-call rotation"],
    ["Turbine Technician (GWO Level 2)", "8", "Day, 4-on/4-off rotation"],
    ["Electrical / Substation Engineer", "2", "Day"],
    ["SCADA Operator", "4", "24/7, four-shift rotation"],
    ["HSE Officer / GWO Safety Trainer", "1", "Day"],
    ["Admin and Procurement Staff", "2", "Day"],
    ["Security Personnel", "6", "24/7"],
    ["TOTAL", "28", ""],
], widths=[4200, 900, 4260]))
A(blank())
A(h2("G.2 Preventive Maintenance Schedule (per turbine)"))
A(tbl([
    ["Frequency", "Activity"],
    ["Continuous (24/7)", "SCADA monitoring: vibration, temperature, power output, fault alarms"],
    ["Weekly", "Ground-level visual + drone oil-leak inspection"],
    ["Monthly", "Tower climb: bolt torque check, lubrication top-up, brake pad inspection"],
    ["Quarterly", "Gearbox oil spectroscopy; blade leading-edge visual check"],
    ["Semi-annual", "Yaw and pitch function test; emergency descender re-certification"],
    ["Annual", "Full PM: oil change, filter change, bolt re-torque (all 144 anchor bolts)"],
    ["Every 2 years", "Generator slip-ring service; Philippine Grid Code compliance test"],
    ["Every 5 years", "Major: gearbox borescope, blade UT scan, transformer oil DGA"],
    ["Every 10 years", "Mid-life overhaul: gearbox refurbishment, generator rewind assessment"],
], widths=[2400, 6960]))
A(blank())
A(h2("G.3 Spare Parts Strategy (Tiered Inventory)"))
for s in [
    "Tier A — On-site (immediate): pitch/yaw motors, fuses, contactors, sensors, lubricants, filters.",
    "Tier B — Regional depot, 24 h delivery: IGBT converter modules, slip rings, brake pads, hydraulic accumulators.",
    "Tier C — Manufacturer warehouse, 2–6 weeks: main bearing, gearbox sub-assembly, generator, blade section.",
]:
    A(bullet(s))
A(blank())
A(h2("G.4 Safety and Compliance"))
for s in [
    "All technicians must hold valid GWO BST certificate: Working at Heights, First Aid, Manual Handling, Fire Awareness.",
    "Two-person rule enforced for any work above 30 m inside or outside the tower.",
    "Hot-work permits required for welding or grinding inside the nacelle.",
    "Annual emergency drill: nacelle fire, tower rescue, medical evacuation.",
    "All work compliant with DOLE Department Order No. 198-18 (OSH Standards).",
]:
    A(bullet(s))
A(pb())

# ── Section VIII-H: Materials and Color Coding ────────────────────────
A(h1("Section VIII-H — Materials and Color-Coding Standards"))
A(blank())
A(h2("H.1 Material Specifications"))
A(tbl([
    ["Component", "Material / Grade / Standard"],
    ["Tower steel", "S355 J0+N per EN 10025-2"],
    ["Tower flange bolts", "ISO 898-1 Class 10.9, hot-dip galvanized"],
    ["Anchor bolts", "ISO 898-1 Class 10.9 or ASTM A615 Grade 75"],
    ["Foundation concrete", "fc' = 35 MPa; ACI 318 / NSCP 2015"],
    ["Foundation rebar", "Grade 60 (fy = 414 MPa) deformed bars — top + bottom mats"],
    ["External coating", "ISO 12944 C5-M: zinc primer + epoxy + PU topcoat, total DFT ≥ 320 µm"],
    ["Blade composite", "Glass-fiber reinforced epoxy (GFRP) + carbon spar caps; IEC 61400-5"],
    ["33 kV inter-array cable", "XLPE-insulated, copper, Class 2; IEC 60502-2"],
    ["115 kV overhead conductor", "ACSR 'Drake' per ANSI/NEMA; NESC clearances"],
    ["Substation grounding", "Bare copper 95 mm² mesh; IEEE 80 — target ≤ 1 Ω"],
    ["Control cabinets", "IP66 painted steel; RAL 7035 light grey"],
    ["Gearbox lubricant", "ISO VG 320 mineral or PAO synthetic"],
    ["Yaw/pitch grease", "NLGI-2 lithium-complex with EP additive"],
], widths=[2800, 6560]))
A(blank())
A(h2("H.2 Color-Coding Scheme"))
A(tbl([
    ["Item", "Color (RAL)", "Standard / Reason"],
    ["Tower body — lower 2/3", "RAL 9010 Pure White", "IEC / aviation standard"],
    ["Tower body — upper 1/3 (2 red bands)", "RAL 3020 Traffic Red", "CAAP MOS Part 14 — daytime obstruction"],
    ["Blade tips (last 6 m, alternating)", "RAL 3020 + RAL 9010", "Aviation visibility"],
    ["Aviation warning lights", "Red, medium-intensity flashing", "ICAO Annex 14 / CAAP"],
    ["Earthing / grounding conductors", "Green/Yellow striped", "IEC 60446"],
    ["33 kV cable phase marking", "L1=Red, L2=Yellow, L3=Blue", "PEC 2017 Philippine standard"],
    ["Fire-fighting equipment", "RAL 3000 Flame Red", "DOLE OSH standards"],
    ["Emergency exits and routes", "RAL 6024 Traffic Green + white text", "ISO 7010"],
    ["115 kV substation control panels", "RAL 7032 Pebble Grey", "Philippine substation convention"],
    ["Hazard / caution floor markings", "RAL 1023 Traffic Yellow + black stripes", "ISO 3864"],
    ["Transformer oil containment bund", "RAL 1018 Zinc Yellow", "High visibility for inspection"],
], widths=[2800, 2800, 2760]))



# ════════════════════════════════════════════════════════════════════
# ASSEMBLE THE COMPLETE DOCUMENT
# ════════════════════════════════════════════════════════════════════

# Combine: original body (without its sectPr) + new sections + sectPr
complete_body = original_body_no_sect + "\n" + "\n".join(S) + "\n" + ORIG_SECT_PR

# Rebuild document.xml using the ORIGINAL namespaces
new_doc_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<w:document{ORIG_NS}>\n'
    f'<w:body>\n'
    f'{complete_body}\n'
    f'</w:body>\n'
    f'</w:document>'
)

# Build updated relationships — keep all original + add any new image refs if needed
# (Original already has rId7-rId14, we don't add new images so rels unchanged)

# Content types — copy original and add any missing defaults
content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="wdp" ContentType="image/vnd.ms-photo"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
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
    z.writestr("word/document.xml", new_doc_xml)
    z.writestr("word/_rels/document.xml.rels", orig_rels_xml)
    z.writestr("word/styles.xml", orig_styles)
    z.writestr("word/numbering.xml", orig_numbering)
    z.writestr("word/header1.xml", orig_header)
    z.writestr("word/footer1.xml", orig_footer)
    z.writestr("word/settings.xml", orig_settings)
    # All original media files (logos + diagrams)
    for name, data in media.items():
        z.writestr(name, data)

size_kb = os.path.getsize(OUTPUT) // 1024
print(f"✅  {OUTPUT}  ({size_kb} KB)")
print(f"   Original body: {len(original_body_no_sect):,} chars")
print(f"   New sections:  {len(chr(10).join(S)):,} chars")
print(f"   Total doc.xml: {len(new_doc_xml):,} chars")
print(f"   Media files:   {len(media)} items")
