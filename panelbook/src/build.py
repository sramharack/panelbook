import sys
import openpyxl
from openpyxl.workbook.properties import CalcProperties
from common import *
from lookups import build_lookups
from panel import build_panel
from summary import build_panels, build_topology

OUT = "Panel_Schedule_Tool.xlsx"

LDU = "Lighting - Dwelling Unit"
PG = "Lighting - Dwelling Unit Parking Garage P01"
HLD = "120/240 3Ph 4W High-Leg Delta"

# name: (header values, circuits, capacity, hide rows beyond No. of Circuits)
PANELS = {
    "HDP": (dict(name="HDP", title="Distribution Panel", loc="Main Electrical Room", manual="Utility XFMR T1",
                 mount="Surface", encl="Type 1", slots=42, dist="120/208 Wye", aic="22 kA", mtype="MCB",
                 mains=400, mcb=400, notes1="Example distribution panel feeding LP000 and LP001."),
            {1: dict(sub="LP000", pol=3, trip=100),
             2: dict(sub="LP001", pol=3, trip=100),
             7: dict(desc="AHU-1", cls="Motor", add=12000, pol=3, trip=50),
             8: dict(desc="CU-1", cls="Cooling", add=9000, pol=3, trip=40),
             17: dict(desc="EF-1", cls="Motor", add=2500, pol=2, trip=20),
             14: dict(desc="Site Lighting", cls="Lighting", q1=12, v1=75, trip=20),
             16: dict(desc="SPARE", trip=20)}, 42, False),
    "LP000": (dict(name="LP000", title="Branch Panel", loc="Utility P05", manual="HDP", mount="Surface",
                   encl="Type 1", slots=42, dist="120/208 Wye", mains=100, mcb=225),
              {1: dict(desc=PG, cls=LDU, q1=14, v1=40, trip=20),
               2: dict(desc=PG, cls=LDU, q1=24, v1=40, trip=20),
               3: dict(desc=PG, cls=LDU, q1=20, v1=40, trip=20),
               4: dict(desc=PG, cls=LDU, q1=26, v1=40, trip=20),
               5: dict(desc=PG, cls=LDU, q1=8, v1=40, trip=20),
               6: dict(desc=LDU, cls=LDU, q1=8, v1=40, trip=20),
               7: dict(desc=LDU, cls=LDU, q1=10, v1=40, trip=20),
               8: dict(desc="Lighting - Dwelling Unit Machine RM P03", cls=LDU, q1=4, v1=40, trip=20),
               9: dict(desc="Other Stair-1 S2-1", cls="Other", add=2400, trip=40),
               10: dict(desc="Other Stair-1 S1-1", cls="Other", add=2600, trip=40),
               11: dict(desc="Other Stair-1 S3-1", cls="Other", add=2400, trip=40),
               12: dict(desc="Lighting - Dwelling Unit Lobby 13", cls=LDU, q1=6, v1=50, trip=20),
               13: dict(desc="Receptacle", cls="Receptacle", q1=2, v1=180, trip=20)}, 42, False),
    "LP001": (dict(name="LP001", title="Branch Panel", loc="Level 1 Corridor", manual="", mount="Recessed",
                   encl="Type 1", slots=42, dist="120/208 Wye", aic="10 kA", mtype="MCB", mains=125, mcb=100),
              {1: dict(desc="Receptacles Office 101", cls="Receptacle", q1=8, v1=180, trip=20),
               2: dict(desc="Receptacles Office 102", cls="Receptacle", q1=6, v1=180, trip=20),
               3: dict(desc="Lighting Office 101-102", cls="Lighting", q1=12, v1=36, q2=4, v2=18, trip=20),
               4: dict(desc="SPARE", trip=20),
               5: dict(desc="Kitchenette - Coffee Maker", cls="Kitchen Equipment", add=1200, trip=20),
               6: dict(desc="Corridor Lighting", cls="Lighting", q1=15, v1=24, trip=20),
               7: dict(desc="Kitchenette - Microwave", cls="Kitchen Equipment", add=1500, trip=20),
               8: dict(desc="Kitchenette - Dishwasher", cls="Kitchen Equipment", add=1800, trip=20),
               13: dict(desc="Water Heater WH-1", cls="Water Heater", add=4500, pol=2, trip=30),
               16: dict(sub="LP001A", pol=2, trip=80)}, 42, False),
    "LP001A": (dict(name="LP001A", title="Branch Panel", loc="Level 1 Lab 104", manual="", mount="Surface",
                    encl="Type 1", slots=24, dist="120/208 1Ph 3W", aic="10 kA", mtype="MLO", mains=100,
                    notes1="Single-phase 3-wire (network) panel: phases alternate A/B; column C is not used."),
               {1: dict(desc="Lab Receptacles North", cls="Receptacle", q1=10, v1=180, trip=20),
                3: dict(desc="Lab Lighting", cls="Lighting", q1=10, v1=40, trip=20),
                4: dict(desc="EV Charger EVC-1", cls="EV Charging", add=5760, pol=2, trip=40),
                5: dict(desc="Fume Hood Fan", cls="Motor", add=1200, trip=20),
                7: dict(desc="Lab Bench Power", cls="Power", add=1500, trip=20),
                8: dict(desc="Lab Receptacles South", cls="Receptacle", q1=10, v1=180, trip=20)}, 42, True),
    # --- anonymised retail store on a 120/240 V high-leg delta service (84-circuit board, 54 used)
    "RMDP": (dict(name="RMDP", title="Distribution Panel", loc="Retail Store - Electrical Room",
                  manual="Utility bank (closed delta)", mount="Surface", encl="Type 1", slots=54, dist=HLD,
                  aic="42 kA", mtype="MCB", mains=600, mcb=500, amb=35,
                  notes1="High-leg delta: 120 V loads on A and C only. 240 V 3-pole chillers and RTUs; "
                         "120/208 V sales-floor panel LP-R1 through a 30 kVA transformer."),
             {1: dict(desc="Chiller CH-1", cls="Cooling", add=30000, pol=3, trip=100),
              2: dict(desc="Chiller CH-2", cls="Cooling", add=30000, pol=3, trip=100),
              7: dict(desc="XFMR T-R1 30 kVA 240 Δ - 120/208 Y", sub="LP-R1", xf="XFMR", pol=3, trip=100),
              8: dict(desc="RTU-1", cls="Cooling", add=15000, pol=3, trip=60),
              13: dict(desc="Receptacles Sales Floor North", cls="Receptacle", q1=10, v1=180, trip=20),
              14: dict(desc="Receptacles Sales Floor South", cls="Receptacle", q1=10, v1=180, trip=20),
              15: dict(desc="AHU-1", cls="Motor", add=5000, pol=3, trip=30),
              16: dict(desc="Walk-in Cooler CU-1", cls="Cooling", add=9000, pol=3, trip=40),
              21: dict(desc="SPARE", pol=3, trip=20),
              22: dict(desc="RTU-2", cls="Cooling", add=15000, pol=3, trip=60),
              27: dict(desc="Exhaust Fan EF-1", cls="Motor", add=2000, pol=3, trip=15),
              30: dict(desc="Water Heater WH-1", cls="Water Heater", add=4500, pol=2, trip=30),
              35: dict(desc="Exterior Signage", cls="Lighting", add=1200, trip=20),
              36: dict(desc="Sales Floor Lighting 1", cls="Lighting", q1=40, v1=36, trip=20),
              37: dict(desc="Exterior Lighting", cls="Lighting", q1=8, v1=100, trip=20),
              38: dict(desc="Sales Floor Lighting 2", cls="Lighting", q1=40, v1=36, trip=20),
              41: dict(desc="Receptacles Stockroom", cls="Receptacle", q1=6, v1=180, trip=20),
              42: dict(desc="Receptacles Office", cls="Receptacle", q1=6, v1=180, trip=20),
              44: dict(desc="POS / Cash Wrap", cls="Power", add=1200, brk="GFCI", trip=20)}, 84, True),
    "LP-R1": (dict(name="LP-R1", title="Panelboard", loc="Retail Store - Sales Floor", manual="",
                   mount="Recessed", encl="Type 1", slots=42, dist="120/208 Wye", aic="10 kA", mtype="MCB",
                   mains=125, mcb=100, notes1="Fed from RMDP through 30 kVA transformer T-R1."),
              {1: dict(desc="Sales Lighting 1", cls="Lighting", q1=30, v1=40, trip=20),
               2: dict(desc="Sales Lighting 2", cls="Lighting", q1=30, v1=40, trip=20),
               3: dict(desc="Display Case Lighting", cls="Lighting", q1=20, v1=24, trip=20),
               4: dict(desc="Receptacles Checkout", cls="Receptacle", q1=8, v1=180, trip=20),
               5: dict(desc="Refrigerated Case RC-1", cls="Power", add=1800, trip=20),
               6: dict(desc="Refrigerated Case RC-2", cls="Power", add=1800, trip=20),
               15: dict(desc="Refrigerated Case RC-3", cls="Power", add=1800, trip=20),
               8: dict(desc="Receptacles Office", cls="Receptacle", q1=6, v1=180, trip=20),
               9: dict(desc="POS Terminals", cls="Power", add=1000, trip=20),
               10: dict(desc="Security / CCTV", cls="Power", add=600, trip=20),
               11: dict(desc="Freezer Case FC-1", cls="Power", add=3000, pol=2, trip=20),
               12: dict(desc="IT Rack", cls="Power", add=1500, trip=20),
               14: dict(desc="Receptacles Stockroom", cls="Receptacle", q1=8, v1=180, trip=20),
               16: dict(desc="Breakroom Microwave", cls="Kitchen Equipment", add=1500, brk="GFCI", trip=20),
               18: dict(desc="Breakroom Refrigerator", cls="Kitchen Equipment", add=800, trip=20)}, 42, False),
    "Template": (dict(name="Template", title="Panelboard", loc=None, manual=None, mount="Surface",
                      encl="Type 1", slots=42, dist="120/208 Wye", mtype="MCB"), {}, 42, False),
    "Template 84": (dict(name="Template 84", title="Panelboard", loc=None, manual=None, mount="Surface",
                         encl="Type 1", slots=84, dist="120/208 Wye", mtype="MCB"), {}, 84, False),
}
TEMPLATES = ("Template", "Template 84")


def build_readme(wb):
    ws = wb.create_sheet("Read Me", 0)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "44546A"
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 120
    lines = [
        ("t", "Panel Schedule Tool"),
        ("s", "No macros. Every number is a formula driven by lookup tables and cell references, so it works in desktop Excel, Excel online and on a Mac. The printed schedule follows the company v005 layout."),
        ("h", "Sheets"),
        ("p", "Panels: register and summary of every panel (type each panel's tab name in column B)."),
        ("p", "Topology: tree of the distribution system built from the sub-panel assignments, with loads, loading, voltage changes (transformers) and flags."),
        ("p", "HDP, LP000, LP001, LP001A: 120/208 V worked example. LP000 reproduces the Revit schedule (4,860 VA Lighting - Dwelling Unit → 75.12%)."),
        ("p", "RMDP, LP-R1: anonymised retail store on a 120/240 V high-leg delta. RMDP is an 84-circuit board with 54 circuits used (unused rows hidden). LP-R1 is a 120/208 V panel fed through a 30 kVA transformer."),
        ("p", "Template (42 circuits) and Template 84 (84 circuits): blank panels to copy."),
        ("p", "Lookups: distribution systems, load classes and demand factors, conductor ampacity (310.16), EGC (250.122), ambient and bundling factors, breaker and bus sizes, lists and settings. Edit freely."),
        ("h", "Panel sheet layout"),
        ("p", "Middle (print area) = the schedule in the company layout. It is 100% formulas. Never type in it."),
        ("p", "Left of the schedule = panel data + ODD circuit inputs.  Right of the schedule = panel checks + EVEN circuit inputs. Each input row lines up with its schedule row."),
        ("p", "Yellow cells with blue text are inputs. Grey cells are calculated. Load (VA) = Fixture 1 Qty × VA + Fixture 2 Qty × VA + Added VA."),
        ("p", "Phase letters follow the usual rotation: CKT 1/2 = A (L-1), 3/4 = B (L-2), 5/6 = C (L-3), repeating (A/B alternating for 1-phase panels). The active phase cells are shaded grey."),
        ("h", "Bigger boards"),
        ("p", "Each sheet has a fixed capacity: 42 circuits (Template) or 84 circuits (Template 84). Copy the one that fits, set No. of Circuits (12 to 84), and hide the unused circuit rows to shorten the printout. 84-circuit sheets print portrait."),
        ("h", "Multi-pole breakers"),
        ("p", "Enter the breaker on its first circuit only (Poles = 2 or 3). The next slot(s) on the same side are taken automatically and show '(pole of CKT x)'. Entering anything in a taken slot is flagged ERR."),
        ("h", "Line currents (TOTAL AMPS)"),
        ("p", "Each circuit adds its line current to the phases it lands on: 1-pole = VA / V(L-N); 2-pole = VA / V(L-L) on each of its two lines; 3-pole = VA / (√3 × V(L-L)) on each line. This is exact for 1- and 2-pole loads and balanced 3-pole loads on every system, including high-leg delta. (Phase VA ÷ V(L-N) overstates 3-pole loads by 15% and understates 2-pole loads on 120/208 V by 13%.)"),
        ("p", "Panel current = the highest line current. Est. demand amps = that × (total demand VA / total connected VA). Loading, Panels and Topology all use it. MCA = next standard breaker at or above the est. demand amps, as on the v005 sheets."),
        ("h", "High-leg (centre-tapped) delta: 120/240 V or 115/230 V, 3-phase 4-wire"),
        ("p", "Pick the high-leg system as the Distribution System. The high leg defaults to B (L-2) from Lookups; 'High-leg phase' on the panel can set A or C (NEC 408.3(E)(1) allows C only in meter equipment, so anything but B is flagged WARN)."),
        ("p", "High-leg rows are shaded orange in the input zones. A 1-pole load there is blocked when you type it (set Poles to 2 or 3 first) and flagged ERR if it gets in anyway (208 V to neutral). A 1-pole breaker with no load there is WARN."),
        ("p", "2-pole breakers that include the high leg are WARN: use a straight-rated 240 V breaker, not 120/240 slash-rated (NEC 240.85). Single-phase 240 V loads are better on A–C. 3-pole loads on 240 V are fine."),
        ("p", "The schedule header shows 'HIGH LEG: L-2 (B) = 208V TO N' and the L-2 column is labelled HIGH LEG. Phase balance on these panels compares A and C only (the lighting winding)."),
        ("h", "Sub-panels (replaces the old IMPORT SUB PANEL button)"),
        ("p", "1. Make sure the sub-panel exists as its own tab and is listed on the Panels sheet."),
        ("p", "2. On the upstream panel, pick an empty slot and choose the sub-panel in 'Sub-Panel (feeds)'. Set Poles (3 for a 3-phase sub, 2 for 1-phase 3-wire) and Trip."),
        ("p", "3. The sub-panel's load comes in per phase and its load classes roll up into this panel's demand table, so demand factors apply to the combined totals."),
        ("p", "4. Through a transformer? Type XFMR in 'Via XFMR'. The sub's VA is then spread evenly over the feeder's poles and its current worked out at this panel's voltage. Without XFMR, a voltage mismatch is WARN, and a neutral-load panel fed from the high leg is ERR. Topology shows the voltage change."),
        ("h", "Cable selection"),
        ("p", "Each circuit with a trip rating gets a cable: the smallest conductor whose allowable ampacity is at least the trip. Allowable = 90 °C ampacity × ambient factor × bundling factor, capped at the termination column (110.14(C), 75 °C default) and the 240.4(D) small-conductor limits. Above the 'largest per set' setting (500 kcmil) it goes to parallel sets (1/0 minimum). The EGC comes from 250.122 by trip, full size in each set."),
        ("p", "Set material, termination, ambient and conductors-per-raceway per panel (header) or as defaults on Lookups. 'Wire override' forces a size, e.g. for voltage drop; the tool does not upsize the EGC for that (250.122(B)) or apply 240.4(B)."),
        ("p", "The panel's own feeder (WIRE SIZE in the header) is sized from the MCB rating, or the bus rating for MLO panels."),
        ("h", "Adding a new panel (replaces NEW SHEET / COPY)"),
        ("p", "1. Right-click Template (or Template 84) > Move or Copy > tick 'Create a copy' > OK.   2. Rename the new tab (e.g. LP002).   3. Type the same name in Panel Name (cell D2).   4. Add the name to the next empty row of Panels column B."),
        ("p", "Tab names are how panels find each other (INDIRECT). If you rename a tab, also update Panel Name and the Panels sheet."),
        ("h", "Demand factors and calculation basis"),
        ("p", "Load classes and demand methods are on Lookups (Constant, Tiered, Quantity, Largest Motor +25%). Defaults follow NEC 220.42/220.44/220.56 and 430.24."),
        ("p", "Totals are straight sums of VA. Revit adds loads vectorially using power factor, so its totals can be slightly lower (LP000: 12,620 VA here, 12,470 VA in Revit)."),
        ("h", "Keep it working"),
        ("p", "Do not insert or delete rows or columns on panel sheets: cross-panel links read fixed cells (hidden calculation columns BK onward). Hiding rows is safe. Add new load classes, systems or breaker types in the blank rows on Lookups."),
        ("p", "Hidden columns hold the calculations. Unhide them to audit any number."),
    ]
    r = 2
    for kind, text in lines:
        c = ws[f"B{r}"]
        c.value = text
        if kind == "t":
            c.font = font(18, True, C_LABEL); ws.row_dimensions[r].height = 28
        elif kind == "s":
            c.font = font(10, False, "595959", italic=True)
        elif kind == "h":
            r += 1
            c = ws[f"B{r}"]; c.value = text
            ws[f"B{r-1}"].value = None
            c.font = font(11, True, "1F4E79")
            c.border = Border(bottom=Side(style="thin", color="1F4E79"))
        else:
            c.font = font(10)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1


def main(out=None, extra=None):
    """extra: optional dict of additional panels (same shape as PANELS) for testing."""
    out = out or OUT
    panels = dict(PANELS)
    if extra:
        panels.update(extra)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_lookups(wb)
    for name, (hv, ckts, cap, hide) in panels.items():
        build_panel(wb, name, hv, ckts, cap=cap, hide_unused=hide,
                    tab_color="7F7F7F" if name in TEMPLATES else "1F4E79")
    listed = [n for n in panels if n not in TEMPLATES]
    build_panels(wb, listed)
    build_topology(wb)
    build_readme(wb)
    order = ["Read Me", "Panels", "Topology"] + listed + list(TEMPLATES) + ["Lookups"]
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.save(out)
    print("saved", out)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
