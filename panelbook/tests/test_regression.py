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


@pytest.fixture(scope="session")
def wb(tmp_path_factory):
    import make
    out = tmp_path_factory.mktemp("build") / "test.xlsx"
    _, calc, n = make.make(out, test=True)
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
    assert cable(wb, "HLTEST", 1) == "2#12 + #12G Cu"            # 20 A 1P
    assert cable(wb, "LP000", 9) == "2#8 + #10G Cu"              # 40 A: 10 AWG is 35 A at 75 C
    assert p(wb, "HDP", "fcab").startswith("2# 3/0 AWG x 4C")    # 400 A feeder
    assert p(wb, "E1", "fcab").startswith("350 kcmil x 4C")      # 200 A Al, 40 C, 6 CCC (factor 0.728)
    assert cable(wb, "E1", 8).startswith("3 sets: 3#4/0")        # override 4/0 Al at 400 A
    assert "310.10(G)" in status(wb, "E1", 14)                   # #8 in parallel
    assert status(wb, "E1", 20).startswith("WARN: wire override not in")
