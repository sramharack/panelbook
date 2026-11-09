"""Shared constants, styles and layout maps for the panel schedule workbook.

v2 (2026-11-02): company (v005) schedule look, variable circuit capacity per sheet
(42 or 84), high-leg delta, line-current method, cable selection.
"""
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter as CL, column_index_from_string as CI

FONT = "Arial"


# ---------- styles ----------
def font(size=8, bold=False, color="000000", italic=False, name=FONT):
    return Font(name=name, size=size, bold=bold, color=color, italic=italic)


def fill(hex_):
    return PatternFill("solid", fgColor=hex_)


LINE = "6E6E6E"
thin = Side(style="thin", color=LINE)
thin_light = Side(style="thin", color="BFBFBF")
blk = Side(style="thin", color="000000")          # schedule lines (company look)
hair = Side(style="hair", color="000000")
med = Side(style="medium", color="000000")
none = Side(style=None)

GRID = Border(left=thin, right=thin, top=thin, bottom=thin)
GRID_LIGHT = Border(left=thin_light, right=thin_light, top=thin_light, bottom=thin_light)
BOX = Border(left=blk, right=blk, top=blk, bottom=blk)

F_INPUT = fill("FFF8D6")        # input cells
F_CALC = fill("F2F2F2")         # calculated cells in input zones
F_ZHDR = fill("DCE3EB")         # zone column headers
F_ZTITLE = fill("44546A")       # zone title band
F_WHITE = fill("FFFFFF")
F_PH = "D9D9D9"                 # active-phase VA cells on the schedule (company look)
F_HL = "FCE4D6"                 # high-leg rows in the input zones (orange, NEC 110.15)

C_INPUT = "0000FF"
C_LABEL = "333F4F"

AL_L = Alignment(horizontal="left", vertical="center")
AL_R = Alignment(horizontal="right", vertical="center")
AL_RI = Alignment(horizontal="right", vertical="center", indent=1)
AL_C = Alignment(horizontal="center", vertical="center")
AL_CW = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL_HDR = Alignment(horizontal="center", vertical="bottom", wrap_text=True)
AL_WRAP_TL = Alignment(horizontal="left", vertical="top", wrap_text=True)

# ---------- panel sheet layout ----------
R0 = 10                    # first circuit row (CKT 1/2)
CAPACITIES = (42, 84)      # sheet sizes built from the templates
NCLS = 20                  # load-class slots (Lookups rows 4..23)
CLS_R0 = 10                # hidden class block rows 10..29
NFEED = 16                 # max sub-panels fed per panel
FEED_R0 = 10               # hidden feeder list rows 10..25
SCH_CLS_ROWS = 10          # load-classification rows printed on the schedule


class Layout:
    """Row map for one panel sheet. Only the circuit rows and what sits below them move
    with capacity; every cell read from another sheet (params, class block, feeder list)
    is at a fixed address."""

    def __init__(self, cap):
        assert cap % 2 == 0 and cap >= 12
        self.cap = cap
        self.R0 = R0
        self.R1 = R0 + cap // 2 - 1
        self.TOT = self.R1 + 1          # TOTAL LOAD VA (each side)
        self.BOT = self.R1 + 4          # load classification header
        self.LAST = self.BOT + SCH_CLS_ROWS


# input zones (odd = left of schedule, even = right)
ZONE_KEYS = ["ckt", "ph", "desc", "cls", "q1", "v1", "q2", "v2", "add", "sub", "xf",
             "pol", "trip", "brk", "wov", "va", "amp", "wire", "stat"]
ODD = {k: CL(1 + i) for i, k in enumerate(ZONE_KEYS)}                     # A..S
ZONE_WIDTHS = {"ckt": 5, "ph": 5.5, "desc": 30, "cls": 22, "q1": 6, "v1": 7,
               "q2": 6, "v2": 7, "add": 8, "sub": 11, "xf": 7, "pol": 5.5, "trip": 6,
               "brk": 10, "wov": 8, "va": 9, "amp": 7, "wire": 22, "stat": 40}

# schedule (print area), company v005 look: I..AB there, U..AN here
SCH_KEYS = ["ckto", "desco", "tripo", "polo", "brko", "wireo", "g1",
            "L1o", "L1e", "L2o", "L2e", "L3o", "L3e", "g2",
            "wiree", "brke", "pole", "tripe", "desce", "ckte"]
_s0 = CI(ODD["stat"]) + 2
S = {k: CL(_s0 + i) for i, k in enumerate(SCH_KEYS)}                     # U..AN
SCH_WIDTHS = {"ckto": 10, "desco": 43.5, "tripo": 15.5, "polo": 11.5, "brko": 13,
              "wireo": 13, "g1": 2.5, "L1o": 10.5, "L1e": 10.5, "L2o": 10.5, "L2e": 10.5,
              "L3o": 10.5, "L3e": 10.5, "g2": 2.5, "wiree": 13, "brke": 13, "pole": 11.5,
              "tripe": 15.5, "desce": 43.5, "ckte": 14}
SCH_FIRST, SCH_LAST = S["ckto"], S["ckte"]
EVEN = {k: CL(CI(SCH_LAST) + 2 + i) for i, k in enumerate(ZONE_KEYS)}     # AP..BH
GAPS = [CL(CI(ODD["stat"]) + 1), CL(CI(SCH_LAST) + 1), CL(CI(EVEN["stat"]) + 1)]
PHASE_COLS = {"A": ("L1o", "L1e"), "B": ("L2o", "L2e"), "C": ("L3o", "L3e")}

# calc zone (hidden). Fixed start leaves spare columns for future input columns.
CALC0 = CI("BK")
assert CALC0 > CI(EVEN["stat"]) + 1
PL, PV, ROWPH = CL(CALC0), CL(CALC0 + 1), CL(CALC0 + 2)
HKEYS = ["sub", "valid", "sVA", "sA", "sB", "sC", "sPol", "sDem", "sIA", "sIB", "sIC",
         "sVLL", "sVLN", "sN", "sHL", "xfm", "own", "VA", "P", "sh1", "sh2", "sh3",
         "c1", "c2", "c3", "has", "cont", "slot", "sla", "amp", "key", "frank",
         "t", "oi", "oa", "ns", "wi", "ei", "nN", "wlab", "cab"]
OH = {k: CL(CALC0 + 3 + i) for i, k in enumerate(HKEYS)}
EH = {k: CL(CALC0 + 3 + len(HKEYS) + i) for i, k in enumerate(HKEYS)}
_c = CALC0 + 3 + 2 * len(HKEYS)
CKEYS = ["cname", "loc", "lcnt", "lmax", "subc", "subn", "subm", "conn", "cnt", "mx",
         "meth", "f1", "l1", "f2", "l2", "f3", "dem", "df", "ord"]
CB = {k: CL(_c + i) for i, k in enumerate(CKEYS)}
_c += len(CKEYS)
FKEYS = ["fname", "fcktn", "fpol", "ftrip", "fxf", "fckt", "fbrk", "fva", "fdem"]
FB = {k: CL(_c + i) for i, k in enumerate(FKEYS)}
_c += len(FKEYS)
GC = [CL(_c + j) for j in range(NFEED)]; _c += NFEED
GN = [CL(_c + j) for j in range(NFEED)]; _c += NFEED
GM = [CL(_c + j) for j in range(NFEED)]; _c += NFEED
CALC_LAST = CL(_c - 1)

# panel parameter / export cells ($PV$n). Rows 1..41 keep their v1 meaning except
# where noted; new parameters are appended from row 42.
P = {"name": 1, "conn": 2, "dem": 3, "demA": 4, "phA": 5, "phB": 6, "phC": 7, "hot": 8,
     "ph": 9, "wires": 10, "vln": 11, "vll": 12, "slots": 13, "div": 14, "loc": 15,
     "dist": 16, "mains": 17, "mcb": 18, "mtype": 19, "err": 20, "warn": 21, "imb": 22,
     "connA": 23, "title": 24, "manual": 25, "nfeed": 26, "aA": 27, "aB": 28, "aC": 29,
     "def": 30, "lim": 31, "empty": 32, "fed": 33, "fedckt": 34, "fedbrk": 35,
     "supply": 36, "avg": 37, "rating": 38, "loading": 39, "ncls": 40, "chk": 41,
     # v2
     "hl": 42, "vhl": 43, "amax": 44, "mat": 45, "fct": 46, "term": 47, "fcab": 48,
     "mca": 49, "neut": 50, "cap": 51, "amb": 52, "ccc": 53, "tcol": 54, "c90": 55,
     "lcol": 56, "minw": 57, "maxw": 58, "maxamp": 59, "i10": 60, "hlin": 61,
     "sysbox": 62, "hlnote": 63, "hl1p": 64, "ambf": 65, "cccf": 66, "fns": 67,
     "fwi": 68, "fei": 69, "aic": 70}
P_LABELS = {
    "name": "Panel name", "conn": "Total connected VA (incl. sub-panels)",
    "dem": "Total est. demand VA", "demA": "Est. demand A (max line A x demand/connected)",
    "phA": "Phase A VA", "phB": "Phase B VA", "phC": "Phase C VA",
    "hot": "Hot conductors (feeder poles)", "ph": "Phases", "wires": "Wires", "vln": "V L-N",
    "vll": "V L-L", "slots": "No. of circuits", "div": "Hot x VLN (reference only)",
    "loc": "Location", "dist": "Distribution system", "mains": "Bus rating",
    "mcb": "MCB rating", "mtype": "Mains type", "err": "Circuit ERR count",
    "warn": "Circuit WARN count", "imb": "Phase imbalance (A-C only on high-leg panels)",
    "connA": "Connected A (= max line A)", "title": "Schedule title",
    "manual": "Supply From (manual)", "nfeed": "Sub-panels fed",
    "aA": "Line A amps (sum of circuit line currents)", "aB": "Line B amps",
    "aC": "Line C amps", "def": "Default load class", "lim": "Breaker load limit",
    "empty": "Empty-slot label", "fed": "Fed from (register)", "fedckt": "Fed from CKT",
    "fedbrk": "Feeder breaker", "supply": "Supply From shown", "avg": "Average phase VA",
    "rating": "Rating used for loading", "loading": "Demand loading",
    "ncls": "Load classes used", "chk": "Phase sum - class sum (should be 0)",
    "hl": "High-leg phase ('' = none)", "vhl": "High leg V to N", "amax": "Max line amps",
    "mat": "Conductor material", "fct": "Ampacity factor (ambient x CCC)",
    "term": "Termination rating C", "fcab": "Feeder cable text",
    "mca": "MCA (next std OCPD >= demand A)",
    "neut": "Has neutral (1/0)", "cap": "Sheet capacity (circuits)", "amb": "Ambient C",
    "ccc": "Current-carrying conductors per raceway", "tcol": "LK_Amp column, termination",
    "c90": "LK_Amp column, 90C", "lcol": "LK_Lim column", "minw": "Min wire index",
    "maxw": "Max wire index per set", "maxamp": "Allowable A of max size per set",
    "i10": "Index of 1/0", "hlin": "High-leg phase input", "sysbox": "System box text",
    "hlnote": "High-leg schedule note", "hl1p": "1P loads on high leg",
    "ambf": "Ambient factor", "cccf": "CCC factor", "fns": "Feeder sets",
    "fwi": "Feeder wire index", "fei": "Feeder EGC index", "aic": "AIC text"}


def pv(key):
    return f"${PV}${P[key]}"


# input cells (header, left zone)
IN = {"name": "D2", "title": "D3", "loc": "D4", "manual": "D5", "mount": "D6",
      "encl": "D7", "slots": "D8", "dist": "I2", "aic": "I3", "mtype": "I4",
      "mains": "I5", "mcb": "I6", "notes1": "I7", "notes2": "I8",
      "hlin": "O2", "mat": "O3", "term": "O4", "amb": "O5", "ccc": "O6"}

# ---------- Panels (register) sheet layout ----------
REG_R0, REG_R1 = 5, 54
REG_VIS = ["idx", "panel", "check", "type", "loc", "fed", "fckt", "fbrk", "sys", "bus",
           "mcb", "connk", "demk", "demA", "loading", "aA", "aB", "aC", "imb", "level",
           "err", "warn", "flags"]
RV = {k: CL(1 + i) for i, k in enumerate(REG_VIS)}   # A..W
_r = CI("Y")
FM = [CL(_r + j) for j in range(NFEED)]; _r += NFEED      # feeder-name matrix
RH = {}
for k in ["cnt", "prow"]:
    RH[k] = CL(_r); _r += 1
PAR = [CL(_r + j) for j in range(8)]; _r += 8             # P1..P8
IDXP = [CL(_r + j) for j in range(7)]; _r += 7            # idx of P1..P7
for k in ["idx", "key", "rank", "islast"]:
    RH[k] = CL(_r); _r += 1
ISLP = [CL(_r + j) for j in range(7)]; _r += 7            # isLast of P1..P7
for k in ["tree", "ok", "manual", "flagstr", "sysx"]:
    RH[k] = CL(_r); _r += 1
REG_LAST = CL(_r - 1)


def reg_rng(col):
    return f"Panels!${col}${REG_R0}:${col}${REG_R1}"
