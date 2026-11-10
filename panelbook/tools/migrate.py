"""Carry a project workbook built with any panelbook version (2.0 onward) forward to the current version.

    python panelbook/tools/migrate.py OLD.xlsx NEW.xlsx [--skip SHEET ...]

Reads every panel sheet's inputs (header cells and circuit input columns, found by their row-9 headers), rebuilds a fresh workbook with the current source, and writes the inputs back in.
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

# Column headers (row 9 of each input zone) -> input keys. Works for every panelbook version since 2.0.
HEADER_KEYS = {
    "circuit description": "desc", "load classification": "cls", "fixture 1 qty": "q1", "fixture 1 va each": "v1",
    "fixture 2 qty": "q2", "fixture 2 va each": "v2", "added va": "add", "sub-panel (feeds)": "sub",
    "via xfmr": "xf", "poles": "pol", "trip (a)": "trip", "breaker type": "brk", "wire override": "wov",
    "va override": "ovr", "override reason": "why", "continuous (blank = class)": "cont"}
HEADER_FIELDS = ["name", "title", "loc", "manual", "mount", "encl", "slots", "dist", "aic", "mtype", "mains", "mcb",
                 "notes1", "notes2", "hlin", "mat", "term", "amb", "ccc", "brk100"]
NOT_PANELS = {"Read Me", "Panels", "Topology", "Lookups", "Template", "Template 84"}


def norm(v):
    return " ".join(str(v).replace("\n", " ").split()).lower() if v is not None else ""


def zone_maps(ws):
    """Find the odd and even input zones and the schedule description columns from the row-9 headers."""
    zones, cur, desc_sched = [], None, []
    for ci in range(1, ws.max_column + 1):
        h = norm(ws.cell(9, ci).value)
        if h == "ckt":
            cur = {"ckt": ci}
            zones.append(cur)
        elif h == "ckt description":
            desc_sched.append(ci)
        elif cur is not None and h in HEADER_KEYS:
            cur[HEADER_KEYS[h]] = ci
        elif h == "status" and cur is not None:
            cur = None
    return zones[:2], desc_sched[:2]


def read_panel(ws_f):
    """Return (header dict, circuits dict, capacity, rescued) from a panel sheet of any panelbook version."""
    last = int(str(ws_f.print_area).split("$")[-1]) if ws_f.print_area else 44
    cap = 84 if last > 50 else 42
    L = Layout(cap)
    hdr = {k: ws_f[IN[k]].value for k in HEADER_FIELDS if k in IN}
    (odd, even), desc_cols = zone_maps(ws_f)
    circuits, rescued = {}, []
    for Z, first, dcol in ((odd, 1, desc_cols[0]), (even, 2, desc_cols[1])):
        for r in range(L.R0, L.R1 + 1):
            ckt = 2 * (r - L.R0) + first
            data = {k: ws_f.cell(r, c).value for k, c in Z.items()
                    if k != "ckt" and ws_f.cell(r, c).value not in (None, "")}
            typed = ws_f.cell(r, dcol).value
            if typed not in (None, "") and not str(typed).startswith("=") and "desc" not in data:
                data["desc"] = str(typed)
                rescued.append(ckt)
            if data:
                circuits[ckt] = data
    return hdr, circuits, cap, rescued


def migrate(old, new, skip=()):
    wf = openpyxl.load_workbook(old)
    register = [wf["Panels"][f"B{r}"].value for r in range(5, 55)] if "Panels" in wf.sheetnames else []
    register = [n for n in register if n]
    panels, report = {}, []
    for name in wf.sheetnames:
        if name in NOT_PANELS or name in skip:
            continue
        if name.endswith(" PRINT"):
            continue
        hdr, ckts, cap, rescued = read_panel(wf[name])
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
