"""Carry a project workbook built with panelbook v2.0.x forward to the current version.

    python panelbook/tools/migrate.py OLD.xlsx NEW.xlsx [--skip SHEET ...]

Reads every panel sheet's inputs (header cells and circuit input columns) from the old workbook using the
v2.0 column layout, rebuilds a fresh workbook with the current source, and writes the inputs back in.
Also rescues text that was typed over a schedule description cell (moved into the circuit's Description
input) and keeps the Panels register order. Calculated cells are never copied; the new tool recalculates.
"""
import argparse
import pathlib
import sys

import openpyxl
from openpyxl.utils import get_column_letter as CL

HERE = pathlib.Path(__file__).resolve().parent
TOOL = HERE.parent
sys.path.insert(0, str(TOOL / "src"))
sys.path.insert(0, str(TOOL))

import build  # noqa: E402
import make  # noqa: E402
from common import IN, Layout  # noqa: E402

# ---- v2.0 layout (panelbook 2.0.0): input zones A..S and AP..BH, schedule U..AN
V20_KEYS = ["ckt", "ph", "desc", "cls", "q1", "v1", "q2", "v2", "add", "sub", "xf",
            "pol", "trip", "brk", "wov", "va", "amp", "wire", "stat"]
V20_INPUTS = ["desc", "cls", "q1", "v1", "q2", "v2", "add", "sub", "xf", "pol", "trip", "brk", "wov"]
V20_ODD = {k: CL(1 + i) for i, k in enumerate(V20_KEYS)}
V20_EVEN = {k: CL(42 + i) for i, k in enumerate(V20_KEYS)}
V20_DESC_SCHED = {"odd": "V", "even": "AM"}
V20_HEADER = ["name", "title", "loc", "manual", "mount", "encl", "slots", "dist", "aic", "mtype", "mains", "mcb",
              "notes1", "notes2", "hlin", "mat", "term", "amb", "ccc"]
NOT_PANELS = {"Read Me", "Panels", "Topology", "Lookups", "Template", "Template 84"}


def read_panel(ws_f, ws_v):
    """Return (header dict, circuits dict, capacity) from a v2.0 panel sheet."""
    cap = 84 if "$65" in (ws_f.print_area or "") else 42
    L = Layout(cap)
    hdr = {k: ws_f[IN[k]].value for k in V20_HEADER}
    circuits, rescued = {}, []
    for side, Z, first in (("odd", V20_ODD, 1), ("even", V20_EVEN, 2)):
        for r in range(L.R0, L.R1 + 1):
            ckt = 2 * (r - L.R0) + first
            data = {k: ws_f[f"{Z[k]}{r}"].value for k in V20_INPUTS
                    if ws_f[f"{Z[k]}{r}"].value not in (None, "")}
            typed = ws_f[f"{V20_DESC_SCHED[side]}{r}"].value
            if typed not in (None, "") and not str(typed).startswith("=") and "desc" not in data:
                data["desc"] = str(typed)
                rescued.append(ckt)
            if data:
                circuits[ckt] = data
    return hdr, circuits, cap, rescued


def migrate(old, new, skip=()):
    wf = openpyxl.load_workbook(old)
    wv = openpyxl.load_workbook(old, data_only=True)
    register = [wf["Panels"][f"B{r}"].value for r in range(5, 55)] if "Panels" in wf.sheetnames else []
    register = [n for n in register if n]
    panels, report = {}, []
    for name in wf.sheetnames:
        if name in NOT_PANELS or name in skip:
            continue
        hdr, ckts, cap, rescued = read_panel(wf[name], wv[name])
        hdr["name"] = name
        hide = isinstance(hdr.get("slots"), int) and hdr["slots"] < cap
        panels[name] = (hdr, ckts, cap, hide)
        report.append(f"{name}: {len(ckts)} circuits, {cap}-circuit sheet"
                      + (f", rescued typed-over descriptions on CKT {rescued}" if rescued else ""))
    # register order first, then any unregistered panel sheets
    order = [n for n in register if n in panels] + [n for n in panels if n not in register]
    build.PANELS = {**{n: panels[n] for n in order},
                    **{t: build.PANELS[t] for t in build.TEMPLATES}}
    out, calc, n = make.make(new)
    calc.unlink()
    return report, n


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("old")
    ap.add_argument("new")
    ap.add_argument("--skip", nargs="*", default=[], help="panel sheets to leave out")
    a = ap.parse_args()
    report, n = migrate(a.old, a.new, a.skip)
    for line in report:
        print("  " + line)
    sys.exit(1 if n else 0)
