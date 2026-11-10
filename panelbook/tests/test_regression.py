"""Regression tests: build the workbook with the test panels, recalc in LibreOffice, check numbers.

    pytest panelbook/tests -v
Skipped automatically if LibreOffice is not installed.
"""
import math
import pathlib
import shutil
import sys

import openpyxl
import pytest

HERE = pathlib.Path(__file__).resolve().parent
TOOL = HERE.parent
sys.path.insert(0, str(TOOL / "src"))
sys.path.insert(0, str(TOOL))
from common import P, PV, ODD, EVEN, EH, CB, CLS_R0, NCLS, Layout  # noqa: E402

pytestmark = pytest.mark.skipif(
    not (shutil.which("soffice") or shutil.which("libreoffice")
         or pathlib.Path(r"C:\Program Files\LibreOffice\program\soffice.exe").exists()),
    reason="LibreOffice not installed")


TYPED_CELL = None


def type_over_schedule_cell(raw):
    """Simulate someone typing a number over a schedule VA cell on E13 (CKT 5 row, L-3 odd column)."""
    global TYPED_CELL
    from common import S, R0
    w = openpyxl.load_workbook(raw)
    TYPED_CELL = f"{S['L3o']}{R0 + 2}"
    w["E13"][TYPED_CELL] = 999
    w.save(raw)


@pytest.fixture(scope="session")
def wb(tmp_path_factory):
    import make
    out = tmp_path_factory.mktemp("build") / "test.xlsx"
    _, calc, n = make.make(out, test=True, post=type_over_schedule_cell)
    assert n == 0, f"{n} formula errors after recalc"
    return openpyxl.load_workbook(calc, data_only=True)


def p(wb, sheet, key):
    return wb[sheet][f"{PV}{P[key]}"].value


def status(wb, sheet, ckt):
    L = Layout(p(wb, sheet, "cap"))
    Z = ODD if ckt % 2 else EVEN
    return wb[sheet][f"{Z['stat']}{L.R0 + (ckt - 1) // 2}"].value or ""


def cable(wb, sheet, ckt):
    L = Layout(p(wb, sheet, "cap"))
    Z = ODD if ckt % 2 else EVEN
    return wb[sheet][f"{Z['wire']}{L.R0 + (ckt - 1) // 2}"].value or ""


# ---------- high-leg regression (from the old-tool check)
def test_highleg_line_amps(wb):
    exp = {"aA": 1000 / 120 + 3000 / 240 + 6000 / (math.sqrt(3) * 240),
           "aB": 500 / 120 + 3000 / 240 + 6000 / (math.sqrt(3) * 240),
           "aC": 800 / 120 + 6000 / (math.sqrt(3) * 240)}
    for k, v in exp.items():
        assert p(wb, "HLTEST", k) == pytest.approx(v, abs=0.005)
    assert round(p(wb, "HLTEST", "aA"), 2) == 35.27
    assert round(p(wb, "HLTEST", "aB"), 2) == 31.10
    assert round(p(wb, "HLTEST", "aC"), 2) == 21.10
    assert round(p(wb, "HLTEST", "amax"), 2) == 35.27


def test_highleg_flags(wb):
    assert status(wb, "HLTEST", 3).startswith("ERR: 1P load on high leg B")
    assert "240.85" in status(wb, "HLTEST", 2)
    assert status(wb, "HLTEST", 1) == "OK"
    assert status(wb, "HLTEST", 7) == "OK"


# ---------- Revit example must not move
def test_lp000_matches_revit(wb):
    ws = wb["LP000"]
    rows = {ws[f"{CB['cname']}{r}"].value: r for r in range(CLS_R0, CLS_R0 + NCLS)}
    r = rows["Lighting - Dwelling Unit"]
    assert ws[f"{CB['conn']}{r}"].value == pytest.approx(4860)
    assert round(ws[f"{CB['df']}{r}"].value, 4) == 0.7512
    assert p(wb, "LP000", "conn") == pytest.approx(12620)


# ---------- examples are clean
@pytest.mark.parametrize("sheet", ["HDP", "LP000", "LP001", "LP001A", "RMDP", "LP-R1"])
def test_examples_have_no_circuit_errors(wb, sheet):
    assert p(wb, sheet, "err") == 0
    assert abs(p(wb, sheet, "chk")) < 0.5


# ---------- edge cases
def test_highleg_on_c(wb):
    assert p(wb, "E1", "hl") == "C"
    assert status(wb, "E1", 11).startswith("ERR: 1P load on high leg C")
    assert status(wb, "E1", 7) == "OK"


def test_neutral_sub_on_high_leg_is_error(wb):
    assert status(wb, "E1", 2).startswith("ERR: E3 neutral loads")
    assert status(wb, "E1", 1) == "OK"          # 120/240 1-ph sub on A-B (lighting winding when HL = C)


def test_115_230_high_leg_voltage(wb):
    assert round(p(wb, "E4", "vhl")) == 199
    assert "199 V" in status(wb, "E4", 3)


def test_highleg_input_ignored_on_wye(wb):
    assert p(wb, "E5", "hl") in ("", None)


def test_cable_selection(wb):
    assert cable(wb, "HLTEST", 1) == "2#12 + #12G Cu  [OCPD]"            # 20 A 1P
    assert cable(wb, "LP000", 9) == "2#8 + #10G Cu  [OCPD]"              # 40 A: 10 AWG is 35 A at 75 C
    assert p(wb, "HDP", "fcab").startswith("2# 3/0 AWG x 4C")            # 400 A feeder
    assert p(wb, "HDP", "fcab").endswith("Cu 75°C")
    # Al, 75 C basis, 40 C (x0.88) and 6 CCC (x0.80): factor 0.704. 200 A needs 500 kcmil (310 x 0.704 = 218 A)
    assert p(wb, "E1", "fcab").startswith("500 kcmil x 4C")
    assert cable(wb, "E1", 8).startswith("4 sets: 3#4/0")                # 4/0 Al: 180 x 0.704 = 127 A per set
    assert status(wb, "E1", 14).startswith("WARN: wire override fails OCPD")   # #8 on 150 A: one set, too small
    assert status(wb, "E1", 20).startswith("WARN: wire override not in")


# ---------- v2.1 cable basis (insulation column, ambient per insulation, four NEC tests)
def test_75C_basis_upsizes_hot_bundled_12awg(wb):
    # 75 C: 25 A x 0.94 x 0.80 = 18.8 A < 20 A -> #10 (the v2.0 bug kept #12)
    assert p(wb, "E6", "ins") == 75
    assert round(p(wb, "E6", "fct"), 3) == 0.752
    assert cable(wb, "E6", 1) == "2#10 + #12G Cu  [OCPD]"


def test_90C_basis_keeps_12awg(wb):
    # 90 C: 30 A x 0.96 x 0.80 = 23.0 A >= 20 A, cap 25 A at 75 C termination -> #12
    assert p(wb, "E7", "ins") == 90
    assert cable(wb, "E7", 1) == "2#12 + #12G Cu  [OCPD]"
    assert "90 °C MIN." in p(wb, "E7", "basis")


def test_derating_floor_warning(wb):
    ws = wb["E6"]
    from common import EVEN
    flags = [ws[f"{EVEN['stat']}{r}"].value for r in range(2, 9)]
    texts = [ws[f"{EVEN['cls']}{r}"].value or "" for r in range(2, 9)]
    i = [k for k, t in enumerate(texts) if "basis" in t][0]
    assert flags[i] == "WARN" and "split raceway" in texts[i]
    i7 = [k for k, t in enumerate([wb["E7"][f"{EVEN['cls']}{r}"].value or "" for r in range(2, 9)]) if "basis" in t][0]
    assert [wb["E7"][f"{EVEN['stat']}{r}"].value for r in range(2, 9)][i7] == "WARN"   # 0.96 x 0.80 = 0.768 < 0.80


def test_continuous_125pct_governs(wb):
    # 2500 VA / 120 V = 20.8 A continuous -> 26.0 A at 75 C -> #10 (#12 is 25 A)
    assert cable(wb, "E8", 1) == "2#10 + #12G Cu  [125%]"
    assert cable(wb, "E8", 3) == "2#12 + #12G Cu  [OCPD]"               # same load, noncontinuous
    assert status(wb, "E8", 1).startswith("WARN: breaker below 125%")    # breaker undersized: still flagged


def test_override_checked_against_tests(wb):
    assert status(wb, "E8", 5).startswith("WARN: wire override fails OCPD")   # #12 on 30 A


# ---------- v2.1 breaker rating (80% standard vs 100% rated)
def check(wb, sheet, label):
    from common import EVEN
    ws = wb[sheet]
    for r in range(2, 9):
        if (ws[f"{EVEN['desc']}{r}"].value or "") == label:
            return ws[f"{EVEN['cls']}{r}"].value or "", ws[f"{EVEN['stat']}{r}"].value
    raise KeyError(label)


def test_standard_breaker_125pct_check(wb):
    # 16.7 A continuous on 20 A: 125% = 20.8 A > 20 A
    assert status(wb, "E8", 7) == "WARN: breaker below 125% of continuous load (210.20(A)/215.3)"
    assert "80% RATED" in p(wb, "E8", "brknote")


def test_all_100pct_rated(wb):
    assert p(wb, "E9", "brkmode") == 2
    assert status(wb, "E9", 7) == "OK"                                    # 16.7 A on 100% rated 20 A
    assert status(wb, "E9", 1) == "WARN: load exceeds 100%-rated breaker"  # 20.8 A on 20 A
    assert cable(wb, "E9", 1) == "2#12 + #12G Cu  [OCPD]"                 # (a) at 100%: 20.8 A <= 25 A
    assert p(wb, "E9", "brknote").startswith("ALL BREAKERS 100% RATED")


def test_mcb_nec_minimum(wb):
    # 89.9 A continuous: 80% MCB needs 112.4 A > 100 A -> ERR; 100% rated MCB needs 89.9 A -> not ERR
    i_dem = 32400 / (math.sqrt(3) * 208)                                  # 89.93 A
    assert p(wb, "E10", "mcbreq") == pytest.approx(1.25 * i_dem, abs=0.01)
    assert check(wb, "E10", "Demand vs MCB / bus")[1] == "ERR"
    assert p(wb, "E11", "mcbreq") == pytest.approx(i_dem, abs=0.01)
    assert check(wb, "E11", "Demand vs MCB / bus")[1] == "WARN"           # 90% loading still above 80% practice flag
    assert p(wb, "E11", "brknote").startswith("MAIN BREAKER 100% RATED")


def test_per_circuit_100pct_feeder(wb):
    # 15.8 A continuous on a 100%-rated 15 A breaker: over 100% -> flagged; type shows on schedule
    assert status(wb, "E12", 2) == "WARN: load exceeds 100%-rated breaker"
    assert "UNLESS NOTED 100% RATED" in p(wb, "E12", "brknote")


# ---------- v2.1 overrides, typed-over detection, schedule display
def row_of(ckt):
    from common import R0
    return R0 + (ckt - 1) // 2


def test_override_flows_through(wb):
    from common import ODD, OH
    ws = wb["E13"]
    r = row_of(1)
    assert ws[f"{OH['own0']}{r}"].value == 1000
    assert ws[f"{ODD['va']}{r}"].value == 1500                       # load column uses the override
    assert p(wb, "E13", "phA") == pytest.approx(1500 + 6000 / 3 + 3000 / 2)   # phase A total follows
    assert p(wb, "E13", "nov") == 1
    assert status(wb, "E13", 1) == "WARN: VA override (calculated 1,000 VA)"


def test_override_log(wb):
    ws = wb["E13"]
    # find the log header and read the first entry
    for r in range(1, 120):
        if ws[f"A{r}"].value == "OVERRIDES IN EFFECT":
            row = [ws[f"{c}{r + 2}"].value for c in "ABDEJ"]
            assert row == [1, "override me", 1000, 1500, "site survey"]
            return
    raise AssertionError("override log not found")


def test_typed_over_schedule_is_flagged(wb):
    assert p(wb, "E13", "ntyped") == 1
    text, flag = check(wb, "E13", "Circuit flags")
    assert flag == "ERR" and "TYPED OVER" in text
    assert p(wb, "HDP", "ntyped") == 0                                # untouched sheets are clean


def test_multipole_text_on_merged_rows(wb):
    from common import S
    ws = wb["E13"]
    d = lambda ckt: ws[f"{S['desco' if ckt % 2 else 'desce']}{row_of(ckt)}"].value
    assert d(3) in ("", None) and d(5) == "THREE POLE LOAD" and d(7) in ("", None)   # 3P: middle row
    assert d(2) == "TWO POLE LOAD" and d(4) in ("", None)                            # 2P: first row
    assert d(9) == "SPARE"                                                           # empty slot label


def test_quiet_240_85_warning(wb):
    assert status(wb, "E14", 2) == "OK"                       # 2P spare, no load
    assert "240.85" in status(wb, "E14", 10)                  # real 2P load on B-C with a standard breaker
    assert status(wb, "E14", 14) == "OK"                      # straight-rated 240 V breaker


def test_system_box_spans_all_va_columns(wb):
    from common import S
    merged = [str(m) for m in wb["E13"].merged_cells.ranges]
    assert f"{S['L1o']}1:{S['L3e']}1" in merged


def test_blank_starter_builds(tmp_path):
    import make
    _, calc, n = make.make(tmp_path / "blank.xlsx", blank=True)
    assert n == 0
    names = openpyxl.load_workbook(calc).sheetnames
    assert "Template" in names and "HDP" not in names and "LP000" not in names
