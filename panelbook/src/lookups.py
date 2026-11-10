"""Lookups sheet: every table the formulas look up into, plus defined names."""
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from common import *

DIST = [  # name, phases, wires, VLN, VLL, note, high-leg phase
    ("120/208 Wye", 3, 4, 120, 208, "", None),
    ("277/480 Wye", 3, 4, 277, 480, "", None),
    ("347/600 Wye", 3, 4, 347, 600, "", None),
    ("230/400 Wye", 3, 4, 230, 400, "", None),
    ("115/200 Wye", 3, 4, 115, 200, "", None),
    ("480 Delta", 3, 3, 277, 480, "No neutral: 1-pole loads are flagged. VLN used only for current math.", None),
    ("120/240 3Ph 4W High-Leg Delta", 3, 4, 120, 240,
     "Closed or open delta, centre-tapped A-C winding. High leg 208 V to N: no 1-pole loads (NEC 110.15, 408.3(E)(1), 408.3(F)(1)).", "B"),
    ("115/230 3Ph 4W High-Leg Delta", 3, 4, 115, 230,
     "As above at 115/230 V. High leg 199 V to N.", "B"),
    ("120/240 1Ph 3W", 1, 3, 120, 240, "Split-phase", None),
    ("115/230 1Ph 3W", 1, 3, 115, 230, "Split-phase", None),
    ("120/208 1Ph 3W", 1, 3, 120, 208, "Network: 2 phases + N of a 120/208 wye", None),
    ("115/200 1Ph 3W", 1, 3, 115, 200, "Network: 2 phases + N of a 115/200 wye", None),
    ("120 1Ph 2W", 1, 2, 120, 120, "One hot + N", None),
]
DIST_ROWS = 16

NEC42 = "NEC 2020 Table 220.42 (Table 220.45 in NEC 2023). Verify against your adopted edition."
CLASSES = [  # name, method, f1, l1, f2, l2, f3, basis, continuous (3 h or more; engineer to confirm)
    ("Cooling", "Constant", 1.0, None, None, None, None, "100%", "Yes"),
    ("EV Charging", "Constant", 1.0, None, None, None, None, "100% (continuous; size OCPD at 125% separately)", "Yes"),
    ("Heating", "Constant", 1.0, None, None, None, None, "100%", "Yes"),
    ("HVAC", "Constant", 1.0, None, None, None, None, "100%", "Yes"),
    ("Kitchen Equipment", "Quantity", 1.0, None, None, None, None, "NEC 220.56: factor by number of units (table at right). Does not apply the 'two largest loads' floor.", "No"),
    ("Lighting", "Constant", 1.0, None, None, None, None, "NEC 220.42 'all others': 100%", "Yes"),
    ("Lighting - Dwelling Unit", "Tiered", 1.0, 3000, 0.35, 120000, 0.25, NEC42, "No"),
    ("Lighting - Hospital", "Tiered", 0.40, 50000, 0.20, None, None, NEC42, "Yes"),
    ("Lighting - Hotel/Motel", "Tiered", 0.60, 20000, 0.50, 100000, 0.35, NEC42, "Yes"),
    ("Lighting - Warehouse", "Tiered", 1.0, 12500, 0.50, None, None, NEC42, "Yes"),
    ("Motor", "Largest Motor +25%", 1.0, None, None, None, None, "NEC 430.24: 100% of all + 25% of the largest motor circuit", "Yes"),
    ("Other", "Constant", 1.0, None, None, None, None, "100%. Also catches loads entered with no class.", "No"),
    ("Power", "Constant", 1.0, None, None, None, None, "100%", "No"),
    ("Receptacle", "Tiered", 1.0, 10000, 0.50, None, None, "NEC 220.44 (non-dwelling): first 10 kVA @100%, rest @50%", "No"),
    ("Water Heater", "Constant", 1.0, None, None, None, None, "100%", "Yes"),
]
QTY = [("1", 1.0), ("2", 1.0), ("3", 0.90), ("4", 0.80), ("5", 0.70), ("6 and over", 0.65)]
BREAKERS = [15, 20, 25, 30, 35, 40, 45, 50, 60, 70, 80, 90, 100, 110, 125, 150, 175, 200,
            225, 250, 300, 350, 400, 450, 500, 600, 700, 800, 1000, 1200, 1600, 2000, 2500,
            3000, 4000]
BUS = [60, 100, 125, 150, 200, 225, 250, 400, 600, 800, 1000, 1200, 1600, 2000, 2500, 3000, 4000]
LISTS = {
    "X": ("Mounting", ["Surface", "Recessed", "Floor"]),
    "Y": ("Enclosure", ["Type 1", "Type 3R", "Type 4", "Type 4X", "Type 12"]),
    "Z": ("Mains Type", ["MCB", "MLO", "Fused Switch"]),
    "AA": ("A.I.C. Rating", ["10 kA", "14 kA", "18 kA", "22 kA", "25 kA", "35 kA", "42 kA", "65 kA", "100 kA"]),
    "AB": ("No. of Circuits", [12, 18, 24, 30, 36, 42, 48, 54, 60, 66, 72, 84]),
    "AC": ("Schedule Title", ["Load Center", "Panelboard", "Branch Panel", "Distribution Panel",
                              "Switchboard", "Lighting Panel", "Power Panel"]),
    "AD": ("Demand Method", ["Constant", "Tiered", "Quantity", "Largest Motor +25%"]),
}
SETTINGS = [  # label, value, fmt, name
    ("Default load class for circuits entered without a class", "Other", None, "LK_DefaultClass"),
    ("Flag circuit load above this % of its breaker trip", 0.80, "0%", "LK_BrkLimit"),
    ("Flag phase imbalance above (max deviation from average)", 0.20, "0%", "LK_ImbLimit"),
    ("Flag panel demand above this % of MCB / bus rating", 0.80, "0%", "LK_MCBWarn"),
    ("Label for empty slots (e.g. SPARE or SPACE; blank = leave empty)", "SPARE", None, "LK_EmptyLabel"),
    ("Default conductor material (Cu / Al)", "Cu", None, "LK_DefMat"),
    ("Default termination rating, °C (NEC 110.14(C))", 75, "0", "LK_DefTerm"),
    ("Default ambient temperature, °C (NEC 310.15(B))", 30, "0", "LK_DefAmb"),
    ("Default current-carrying conductors per raceway (NEC 310.15(C))", 3, "0", "LK_DefCCC"),
    ("Smallest conductor the tool will pick (branch practice)", "12", "@", "LK_MinWire"),
    ("Largest conductor per parallel set (above this, go parallel)", "500", "@", "LK_MaxWire"),
    ("Default insulation rating, °C: design basis (75 = any NEC building wire; 90 = THHN/THWN-2, XHHW-2 only)", 75, "0", "LK_DefIns"),
    ("Flag combined derating factor below (company practice)", 0.80, "0%", "LK_DerateMin"),
    ("Apply NEC 240.4(B) next size up (Yes / No)", "No", None, "LK_Use2404B"),
]

# NEC 2023 Table 310.16 (same values as 2017/2020). Cu 60/75/90 C, Al 60/75/90 C.
# Al 14 AWG is not listed: ampacity 0 so it is never selected.
WIRE = [  # size, label, short, cu60, cu75, cu90, al60, al75, al90, 240.4(D) Cu, 240.4(D) Al
    ("14", "14 AWG", "#14", 15, 20, 25, 0, 0, 0, 15, None),
    ("12", "12 AWG", "#12", 20, 25, 30, 15, 20, 25, 20, 15),
    ("10", "10 AWG", "#10", 30, 35, 40, 25, 30, 35, 30, 25),
    ("8", "8 AWG", "#8", 40, 50, 55, 35, 40, 45, None, None),
    ("6", "6 AWG", "#6", 55, 65, 75, 40, 50, 55, None, None),
    ("4", "4 AWG", "#4", 70, 85, 95, 55, 65, 75, None, None),
    ("3", "3 AWG", "#3", 85, 100, 115, 65, 75, 85, None, None),
    ("2", "2 AWG", "#2", 95, 115, 130, 75, 90, 100, None, None),
    ("1", "1 AWG", "#1", 110, 130, 145, 85, 100, 115, None, None),
    ("1/0", "1/0 AWG", "#1/0", 125, 150, 170, 100, 120, 135, None, None),
    ("2/0", "2/0 AWG", "#2/0", 145, 175, 195, 115, 135, 150, None, None),
    ("3/0", "3/0 AWG", "#3/0", 165, 200, 225, 130, 155, 175, None, None),
    ("4/0", "4/0 AWG", "#4/0", 195, 230, 260, 150, 180, 205, None, None),
    ("250", "250 kcmil", "#250kcmil", 215, 255, 290, 170, 205, 230, None, None),
    ("300", "300 kcmil", "#300kcmil", 240, 285, 320, 195, 230, 260, None, None),
    ("350", "350 kcmil", "#350kcmil", 260, 310, 350, 210, 250, 280, None, None),
    ("400", "400 kcmil", "#400kcmil", 280, 335, 380, 225, 270, 305, None, None),
    ("500", "500 kcmil", "#500kcmil", 320, 380, 430, 260, 310, 350, None, None),
    ("600", "600 kcmil", "#600kcmil", 350, 420, 475, 285, 340, 385, None, None),
    ("700", "700 kcmil", "#700kcmil", 385, 460, 520, 315, 375, 425, None, None),
    ("750", "750 kcmil", "#750kcmil", 400, 475, 535, 320, 385, 435, None, None),
    ("800", "800 kcmil", "#800kcmil", 410, 490, 555, 330, 395, 445, None, None),
    ("900", "900 kcmil", "#900kcmil", 435, 520, 585, 355, 425, 480, None, None),
    ("1000", "1000 kcmil", "#1000kcmil", 455, 545, 615, 375, 445, 500, None, None),
]
# NEC Table 250.122: OCPD rating not exceeding (A) -> EGC size (Cu, Al)
EGC = [(15, "14", "12"), (20, "12", "10"), (60, "10", "8"), (100, "8", "6"), (200, "6", "4"),
       (300, "4", "2"), (400, "3", "1"), (500, "2", "1/0"), (600, "1", "2/0"), (800, "1/0", "3/0"),
       (1000, "2/0", "4/0"), (1200, "3/0", "250"), (1600, "4/0", "350"), (2000, "250", "400"),
       (2500, "350", "600"), (3000, "400", "600"), (4000, "500", "750")]
# NEC Table 310.15(B)(1)(1): ambient correction, 30 C base, by insulation rating
AMB = [  # upper bound C, factor 60 C, 75 C, 90 C (0 = conductor not permitted at that ambient)
    (10, 1.29, 1.20, 1.15), (15, 1.22, 1.15, 1.12), (20, 1.15, 1.11, 1.08), (25, 1.08, 1.05, 1.04),
    (30, 1.00, 1.00, 1.00), (35, 0.91, 0.94, 0.96), (40, 0.82, 0.88, 0.91), (45, 0.71, 0.82, 0.87),
    (50, 0.58, 0.75, 0.82), (55, 0.41, 0.67, 0.76), (60, 0, 0.58, 0.71), (65, 0, 0.47, 0.65),
    (70, 0, 0.33, 0.58), (75, 0, 0, 0.50), (80, 0, 0, 0.41), (85, 0, 0, 0.29)]
# NEC Table 310.15(C)(1): more than 3 current-carrying conductors. (upper bound count, factor)
CCC = [(3, 1.00), (6, 0.80), (9, 0.70), (20, 0.50), (30, 0.45), (40, 0.40), (9999, 0.35)]
BRKTYPES = ["Standard", "100% rated", "Straight 240V", "GFCI", "AFCI", "AFCI/GFCI (DF)", "GFEP (30 mA)", "Shunt Trip", "HACR",
            "Lock-on", "SWD", "Fused"]


def add_name(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def hdr(ws, cell, text, width_cols=None):
    c = ws[cell]
    c.value = text
    c.font = font(8, True, C_LABEL)
    c.fill = F_ZHDR
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = GRID_LIGHT


def build_lookups(wb):
    ws = wb.create_sheet("Lookups")
    ws.sheet_properties.tabColor = "808080"
    ws.sheet_view.showGridLines = False
    ws["A1"] = "LOOKUP TABLES"
    ws["A1"].font = font(14, True, C_LABEL)
    ws["A2"] = ("Every panel sheet reads these tables. Edit or add rows freely (blue = editable). "
                "Keep table positions fixed; add new rows inside the blank rows provided.")
    ws["A2"].font = font(9, False, "595959", italic=True)
    ws.row_dimensions[3].height = 30

    def inp(cell, val, fmt=None, align=AL_L):
        c = ws[cell]
        c.value = val
        c.font = font(9, False, C_INPUT)
        c.fill = F_INPUT
        c.border = GRID_LIGHT
        c.alignment = align
        if fmt:
            c.number_format = fmt

    def fixed(cell, val, fmt=None, align=AL_C):
        c = ws[cell]
        c.value = val
        c.font = font(9)
        c.border = GRID_LIGHT
        c.alignment = align
        if fmt:
            c.number_format = fmt

    # --- Distribution systems A..G, rows 4..19
    for col, t in zip("ABCDEFG", ["Distribution System", "Phases", "Wires", "V L-N", "V L-L", "Notes",
                                  "High-leg\nphase"]):
        hdr(ws, f"{col}3", t)
    for i in range(DIST_ROWS):
        r = 4 + i
        row = DIST[i] if i < len(DIST) else (None,) * 7
        inp(f"A{r}", row[0]); inp(f"B{r}", row[1], "0", AL_C); inp(f"C{r}", row[2], "0", AL_C)
        inp(f"D{r}", row[3], "0", AL_C); inp(f"E{r}", row[4], "0", AL_C); inp(f"F{r}", row[5] or None)
        ws[f"F{r}"].font = font(8, False, "595959")
        ws[f"F{r}"].alignment = Alignment(wrap_text=True, vertical="center")
        inp(f"G{r}", row[6], None, AL_C)
    add_name(wb, "LK_DistName", "Lookups!$A$4:$A$19")
    add_name(wb, "LK_DistPh", "Lookups!$B$4:$B$19")
    add_name(wb, "LK_DistW", "Lookups!$C$4:$C$19")
    add_name(wb, "LK_DistVLN", "Lookups!$D$4:$D$19")
    add_name(wb, "LK_DistVLL", "Lookups!$E$4:$E$19")
    add_name(wb, "LK_DistHL", "Lookups!$G$4:$G$19")
    add_name(wb, "DD_Dist", "OFFSET(Lookups!$A$4,0,0,MAX(1,COUNTA(Lookups!$A$4:$A$19)),1)")
    dvh = DataValidation(type="list", formula1='"A,B,C"', allow_blank=True)
    ws.add_data_validation(dvh); dvh.add("G4:G19")

    # --- Settings A22..B36
    ws["A22"] = "SETTINGS"
    ws["A22"].font = font(10, True, C_LABEL)
    for i, (lab, val, fmt, nm) in enumerate(SETTINGS):
        r = 23 + i
        ws[f"A{r}"] = lab
        ws[f"A{r}"].font = font(8, False, C_LABEL)
        ws[f"A{r}"].alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = 24
        inp(f"B{r}", val, fmt, AL_C)
        add_name(wb, nm, f"Lookups!$B${r}")
    dv = DataValidation(type="list", formula1="DD_Class", allow_blank=False)
    ws.add_data_validation(dv); dv.add("B23")
    for addr, f1 in [("B28", '"Cu,Al"'), ("B29", '"60,75,90"'), ("B32", "LK_WSize"), ("B33", "LK_WSize"),
                     ("B34", '"75,90"'), ("B36", '"Yes,No"')]:
        d = DataValidation(type="list", formula1=f1, allow_blank=False)
        ws.add_data_validation(d); d.add(addr)

    # --- Load classifications H..O rows 4..23
    for col, t in zip(["H", "I", "J", "K", "L", "M", "N", "O"],
                      ["Load Classification", "Demand Method", "Tier 1 %\n(or flat %)",
                       "Tier 1 applies up to (VA)", "Tier 2 %", "Tier 2 applies up to (VA)",
                       "Tier 3 % (remainder)", "Basis / notes"]):
        hdr(ws, f"{col}3", t)
    for col, t in [("P", "Continuous\nload?")]:
        hdr(ws, f"{col}3", t)
    for i in range(NCLS):
        r = 4 + i
        row = CLASSES[i] if i < len(CLASSES) else (None,) * 9
        inp(f"H{r}", row[0]); inp(f"I{r}", row[1])
        inp(f"J{r}", row[2], "0%", AL_C); inp(f"K{r}", row[3], "#,##0", AL_C)
        inp(f"L{r}", row[4], "0%", AL_C); inp(f"M{r}", row[5], "#,##0", AL_C)
        inp(f"N{r}", row[6], "0%", AL_C); inp(f"O{r}", row[7])
        ws[f"O{r}"].font = font(8, False, "595959")
        ws[f"O{r}"].alignment = Alignment(wrap_text=True, vertical="center")
        inp(f"P{r}", row[8], None, AL_C)
    dvc = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    ws.add_data_validation(dvc); dvc.add("P4:P23")
    add_name(wb, "LK_Cont", "Lookups!$P$4:$P$23")
    for nm, col in [("LK_ClassName", "H"), ("LK_Method", "I"), ("LK_F1", "J"), ("LK_L1", "K"),
                    ("LK_F2", "L"), ("LK_L2", "M"), ("LK_F3", "N")]:
        add_name(wb, nm, f"Lookups!${col}$4:${col}$23")
    add_name(wb, "DD_Class", "OFFSET(Lookups!$H$4,0,0,MAX(1,COUNTA(Lookups!$H$4:$H$23)),1)")
    dvm = DataValidation(type="list", formula1="LK_Methods", allow_blank=True)
    ws.add_data_validation(dvm); dvm.add("I4:I23")
    ws["H25"] = ("How the methods work:  Constant = Tier 1 % on everything.  Tiered = Tier 1 % up to the Tier 1 limit, "
                 "Tier 2 % up to the Tier 2 limit, Tier 3 % on the rest (blank limit = no further tier).  "
                 "Quantity = connected load x factor for the number of circuits of that class (table at right).  "
                 "Largest Motor +25% = Tier 1 % of all + 25% of the largest single circuit of that class.  "
                 "Demand is applied to the class total of each panel INCLUDING its sub-panels (as Revit does).")
    ws["H25"].font = font(8, False, "595959", italic=True)
    ws["H25"].alignment = AL_WRAP_TL
    ws.merge_cells("H25:O28")

    # --- Qty factors Q..R rows 4..9
    hdr(ws, "Q3", "No. of units\n(NEC 220.56)"); hdr(ws, "R3", "Demand factor")
    for i, (n, f) in enumerate(QTY):
        r = 4 + i
        ws[f"Q{r}"] = n; ws[f"Q{r}"].font = font(9); ws[f"Q{r}"].alignment = AL_C; ws[f"Q{r}"].border = GRID_LIGHT
        inp(f"R{r}", f, "0%", AL_C)
    add_name(wb, "LK_QtyFactor", "Lookups!$R$4:$R$9")

    # --- Breakers T, Bus V
    hdr(ws, "T3", "Breaker trip\nratings (A)")
    for i, b in enumerate(BREAKERS):
        inp(f"T{4+i}", b, "0", AL_C)
    add_name(wb, "LK_Breakers", f"Lookups!$T$4:$T${3+len(BREAKERS)}")
    hdr(ws, "V3", "Bus / mains\nratings (A)")
    for i, b in enumerate(BUS):
        inp(f"V{4+i}", b, "0", AL_C)
    add_name(wb, "LK_Bus", f"Lookups!$V$4:$V${3+len(BUS)}")

    # --- simple lists
    names = {"X": "LK_Mounting", "Y": "LK_Enclosure", "Z": "LK_MainsType", "AA": "LK_AIC",
             "AB": "LK_Slots", "AC": "LK_Titles", "AD": "LK_Methods"}
    for col, (title, items) in LISTS.items():
        hdr(ws, f"{col}3", title)
        for i, it in enumerate(items):
            inp(f"{col}{4+i}", it, "0" if isinstance(it, int) else None, AL_C)
        add_name(wb, names[col], f"Lookups!${col}$4:${col}${3+len(items)}")

    # --- Conductors AF..AP rows 4..27 (NEC 310.16)
    ws["AF1"] = "CONDUCTOR SELECTION"
    ws["AF1"].font = font(10, True, C_LABEL)
    ws["AF2"] = ("Ampacity: NEC Table 310.16 (2017-2023 values), not more than 3 current-carrying conductors, "
                 "30 °C. The tool derates the 90 °C value for ambient and bundling, caps it at the "
                 "termination column (110.14(C)), applies 240.4(D), and picks the smallest size whose "
                 "allowable ampacity ≥ breaker trip. Verify against your adopted edition.")
    ws["AF2"].font = font(8, False, "595959", italic=True)
    ws["AF2"].alignment = AL_WRAP_TL
    ws.merge_cells("AF2:AP2")
    ws.row_dimensions[2].height = 36
    heads = ["Size", "Label", "Short", "Cu 60 °C", "Cu 75 °C", "Cu 90 °C", "Al 60 °C", "Al 75 °C",
             "Al 90 °C", "240.4(D) max OCPD Cu", "240.4(D) max OCPD Al"]
    cols = ["AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP"]
    for col, t in zip(cols, heads):
        hdr(ws, f"{col}3", t)
    for i, row in enumerate(WIRE):
        r = 4 + i
        fixed(f"AF{r}", row[0], "@"); fixed(f"AG{r}", row[1]); fixed(f"AH{r}", row[2])
        for j, col in enumerate(cols[3:9]):
            inp(f"{col}{r}", row[3 + j], "0", AL_C)
        inp(f"AO{r}", row[9], "0", AL_C); inp(f"AP{r}", row[10], "0", AL_C)
    n = 3 + len(WIRE)
    add_name(wb, "LK_WSize", f"Lookups!$AF$4:$AF${n}")
    add_name(wb, "LK_WLabel", f"Lookups!$AG$4:$AG${n}")
    add_name(wb, "LK_WShort", f"Lookups!$AH$4:$AH${n}")
    add_name(wb, "LK_Amp", f"Lookups!$AI$4:$AN${n}")
    add_name(wb, "LK_Lim", f"Lookups!$AO$4:$AP${n}")
    ws[f"AF{n+2}"] = ("Al 14 AWG is not listed in 310.16 (ampacity 0 = never picked). 240.4(D): blank = no "
                      "small-conductor limit. Parallel sets: 1/0 minimum (310.10(G)); EGC full size in each "
                      "raceway (250.122(F)). The tool does not apply 240.4(B) (next size up) or upsize the "
                      "EGC for voltage drop (250.122(B)).")
    ws[f"AF{n+2}"].font = font(8, False, "595959", italic=True)
    ws[f"AF{n+2}"].alignment = AL_WRAP_TL
    ws.merge_cells(f"AF{n+2}:AP{n+4}")

    # --- EGC AR..AT rows 4..20 (NEC 250.122)
    for col, t in zip(["AR", "AS", "AT"], ["OCPD not\nexceeding (A)", "EGC Cu", "EGC Al"]):
        hdr(ws, f"{col}3", t)
    for i, (a, cu, al) in enumerate(EGC):
        r = 4 + i
        inp(f"AR{r}", a, "0", AL_C); inp(f"AS{r}", cu, "@", AL_C); inp(f"AT{r}", al, "@", AL_C)
    m = 3 + len(EGC)
    add_name(wb, "LK_EgcA", f"Lookups!$AR$4:$AR${m}")
    add_name(wb, "LK_Egc", f"Lookups!$AS$4:$AT${m}")
    ws["AR1"] = "NEC 250.122"
    ws["AR1"].font = font(9, True, C_LABEL)

    # --- Ambient AV..AY (310.15(B)(1)(1): 60/75/90 C columns), CCC BA..BB (310.15(C)(1)), breaker types BD
    for col, t in zip(["AV", "AW", "AX", "AY"], ["Ambient up\nto (°C)", "60 °C\ninsulation", "75 °C\ninsulation",
                                                 "90 °C\ninsulation"]):
        hdr(ws, f"{col}3", t)
    for i, (t, f60, f75, f90) in enumerate(AMB):
        inp(f"AV{4+i}", t, "0", AL_C)
        for col, f in zip(["AW", "AX", "AY"], [f60, f75, f90]):
            inp(f"{col}{4+i}", f, "0.00", AL_C)
    add_name(wb, "LK_AmbT", f"Lookups!$AV$4:$AV${3+len(AMB)}")
    add_name(wb, "LK_AmbF", f"Lookups!$AW$4:$AY${3+len(AMB)}")
    ws["AV1"] = "NEC 310.15(B)(1)(1)"
    ws["AV1"].font = font(9, True, C_LABEL)
    ws["AV2"] = "0.00 = not permitted"
    ws["AV2"].font = font(8, False, "595959", italic=True)
    for col, t in zip(["BA", "BB"], ["Current-carrying\nconductors up to", "Factor"]):
        hdr(ws, f"{col}3", t)
    for i, (k, f) in enumerate(CCC):
        inp(f"BA{4+i}", k, "0", AL_C); inp(f"BB{4+i}", f, "0.00", AL_C)
    add_name(wb, "LK_CccN", f"Lookups!$BA$4:$BA${3+len(CCC)}")
    add_name(wb, "LK_CccF", f"Lookups!$BB$4:$BB${3+len(CCC)}")
    ws["BA1"] = "NEC 310.15(C)(1)"
    ws["BA1"].font = font(9, True, C_LABEL)

    hdr(ws, "BD3", "Breaker type")
    for i, b in enumerate(BRKTYPES):
        inp(f"BD{4+i}", b, None, AL_C)
    add_name(wb, "LK_BrkType", f"Lookups!$BD$4:$BD$19")   # blank rows free for additions

    widths = {"A": 30, "B": 8, "C": 7, "D": 7, "E": 7, "F": 34, "G": 9, "H": 24, "I": 19,
              "J": 9, "K": 11, "L": 8, "M": 11, "N": 10, "O": 46, "P": 11, "Q": 11, "R": 10,
              "S": 2, "T": 11, "U": 2, "V": 11, "W": 2, "X": 11, "Y": 11, "Z": 12, "AA": 12,
              "AB": 11, "AC": 18, "AD": 20, "AE": 2, "AF": 7, "AG": 11, "AH": 11,
              "AI": 7, "AJ": 7, "AK": 7, "AL": 7, "AM": 7, "AN": 7, "AO": 10, "AP": 10, "AQ": 2,
              "AR": 10, "AS": 8, "AT": 8, "AU": 2, "AV": 9, "AW": 10, "AX": 10, "AY": 10, "AZ": 2,
              "BA": 13, "BB": 8, "BC": 2, "BD": 16}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    for r in range(4, 28):
        ws.row_dimensions[r].height = 22
    ws.freeze_panes = "A4"
    return ws
