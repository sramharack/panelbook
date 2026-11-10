"""Print tabs: a print/CAD copy of each panel schedule.

Every cell is a live formula pointing at the panel sheet, so loads, descriptions and totals stay current.
What a formula cannot do, the print tab does at build time: 2- and 3-pole breakers become real merged cells
with the text centred across their rows, and phase shading is a fixed fill (it survives an AutoCAD paste).

Merges follow the breaker layout at build time. If poles change later, a banner under the schedule turns
red ("layout changed"); rebuild with tools/migrate.py to refresh.
"""
from copy import copy

from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import FormulaRule
from common import *
from lookups import DIST

TEXT_KEYS = {"odd": ("desco", "tripo", "polo", "brko", "wireo"), "even": ("desce", "tripe", "pole", "brke", "wiree")}
NUMERIC = {"tripo", "polo", "tripe", "pole"}


def system_phases(dist):
    for name, ph, wires, *_ in DIST:
        if name == dist:
            return ph, wires
    return 3, 4


def breaker_poles(data, panels):
    """Poles a circuit's breaker takes, as the panel sheet will compute it."""
    if isinstance(data.get("pol"), (int, float)):
        return max(1, min(3, int(data["pol"])))
    sub = (data.get("sub") or "").strip()
    if sub and sub in panels:
        ph, wires = system_phases(panels[sub][0].get("dist") or "120/208 Wye")
        return 3 if ph == 3 else (2 if wires >= 3 else 1)
    return 1


def layout(circuits, cap, panels):
    """Per side: {row: poles of the breaker starting there}, and the per-row P values the sheet computes."""
    L = Layout(cap)
    starts, prow = {"odd": {}, "even": {}}, {"odd": {}, "even": {}}
    for side, first in (("odd", 1), ("even", 2)):
        taken = set()
        for r in range(L.R0, L.R1 + 1):
            ckt = 2 * (r - L.R0) + first
            data = circuits.get(ckt, {})
            p = breaker_poles(data, panels) if data else 1
            prow[side][r] = p
            if r in taken:
                continue
            if data and p > 1:
                span = [rr for rr in range(r, min(r + p, L.R1 + 1))]
                starts[side][r] = len(span)
                taken.update(span[1:])
    return starts, prow


def signature(prow, L):
    """Same number the sheet computes in its psig parameter."""
    return (sum(prow["odd"][r] * r * r for r in range(L.R0, L.R1 + 1))
            + 3 * sum(prow["even"][r] * r * r for r in range(L.R0, L.R1 + 1)))


def build_print_tab(wb, name, hdr, circuits, cap, hide, panels):
    live = wb[name]
    L = Layout(cap)
    tab = wb.create_sheet(f"{name} PRINT"[:31])
    tab.sheet_properties.tabColor = "C65911"
    tab.sheet_view.showGridLines = False
    tab.sheet_view.zoomScale = 85
    q = f"'{name}'"
    c0, c1 = CI(SCH_FIRST), CI(SCH_LAST)

    for ci in range(1, c1 + 3):
        col = CL(ci)
        tab.column_dimensions[col].width = live.column_dimensions[col].width if ci >= c0 else 1.5
    for r in range(1, L.LAST + 3):
        tab.row_dimensions[r].height = live.row_dimensions[r].height
        if live.row_dimensions[r].hidden:
            tab.row_dimensions[r].hidden = True

    starts, prow = layout(circuits, cap, panels)
    block_cells = {}
    for side, keys in TEXT_KEYS.items():
        for r, n in starts[side].items():
            for k in keys:
                block_cells[(S[k], r)] = n
                for rr in range(r + 1, r + n):
                    block_cells[(S[k], rr)] = 0      # covered by the merge

    live_merges = [m for m in live.merged_cells.ranges if m.min_col >= c0 and m.max_col <= c1]
    for r in range(1, L.LAST + 1):
        for ci in range(c0, c1 + 1):
            col = CL(ci)
            src, dst = live.cell(r, ci), tab.cell(r, ci)
            dst._style = copy(src._style)
            n = block_cells.get((col, r))
            if n == 0:
                continue
            if n:
                rng = f"{q}!{col}{r}:{col}{r + n - 1}"
                if any(col == S[k] for k in NUMERIC):
                    dst.value = f'=IF(SUM({rng})=0,"",SUM({rng}))'
                else:
                    dst.value = "=" + "&".join(f"{q}!{col}{rr}" for rr in range(r, r + n))
                continue
            v = src.value
            if isinstance(v, str) and v.startswith("="):
                dst.value = f"={q}!{col}{r}"
            elif v is not None:
                dst.value = v
    for m in live_merges:
        tab.merge_cells(str(m))
    for (col, r), n in block_cells.items():
        if n:
            tab.merge_cells(f"{col}{r}:{col}{r + n - 1}")
            top = tab[f"{col}{r}"]
            top.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for rr in range(r, r + n):
                tab[f"{col}{rr}"].border = Border(left=blk, right=blk, top=blk if rr == r else None,
                                                  bottom=blk if rr == r + n - 1 else None)
    # single-row circuits: full grid lines
    for side, keys in TEXT_KEYS.items():
        for r in range(L.R0, L.R1 + 1):
            if any((S[k], r) in block_cells for k in keys):
                continue
            for k in keys:
                tab[f"{S[k]}{r}"].border = BOX

    # phase shading for this panel's system (fixed fill)
    ph, wires = system_phases(hdr.get("dist") or "120/208 Wye")
    letters = "ABC" if ph == 3 else ("AB" if wires >= 3 else "A")
    slots = hdr.get("slots") if isinstance(hdr.get("slots"), int) else cap
    for r in range(L.R0, L.R1 + 1):
        active = letters[(r - L.R0) % len(letters)] if r - L.R0 < slots // 2 else None
        for letter, (a, z) in PHASE_COLS.items():
            for col in (S[a], S[z]):
                tab[f"{col}{r}"].fill = fill(F_PH if letter == active else "FFFFFF")

    # sync banner (below the print area, not printed)
    sig = signature(prow, L)
    b = L.LAST + 2
    tab.merge_cells(f"{SCH_FIRST}{b}:{SCH_LAST}{b}")
    cell = tab[f"{SCH_FIRST}{b}"]
    cell.value = (f'=IF({q}!${PV}${P["psig"]}={sig},"Print tab in sync with {name} (breaker layout unchanged).",'
                  f'"BREAKER LAYOUT CHANGED ON {name}: merges below are out of date. Regenerate the print tabs '
                  f'(panelbook tools/migrate.py).")')
    cell.font = font(11, True, "2E7D32")
    cell.alignment = Alignment(horizontal="left", vertical="center")
    tab.row_dimensions[b].height = 24
    tab.conditional_formatting.add(f"{SCH_FIRST}{b}", FormulaRule(
        formula=[f'{q}!${PV}${P["psig"]}<>{sig}'], fill=fill("FDE2E1"), font=font(11, True, "C00000")))

    tab.defined_names["Schedule"] = DefinedName("Schedule", attr_text=f"'{tab.title}'!${SCH_FIRST}$1:${SCH_LAST}${L.LAST}")
    tab.print_area = f"{SCH_FIRST}1:{SCH_LAST}{L.LAST}"
    tab.page_setup.orientation = live.page_setup.orientation
    tab.page_setup.paperSize = live.page_setup.paperSize
    tab.page_setup.fitToWidth = 1
    tab.page_setup.fitToHeight = 1
    tab.sheet_properties.pageSetUpPr.fitToPage = True
    tab.print_options.horizontalCentered = True
    tab.page_margins = copy(live.page_margins)
    tab.oddFooter.right.text = f"panelbook {VERSION}"
    tab.oddFooter.right.size = 7
    tab.sheet_view.selection[0].activeCell = f"{SCH_FIRST}1"
    tab.sheet_view.selection[0].sqref = f"{SCH_FIRST}1"
    return tab
