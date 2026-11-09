<p align="center">
  <img src="panelbook/docs/img/panelbook-logo.png" width="112" alt="panelbook logo">
</p>

<h1 align="center">panelbook</h1>
<p align="center"><b>Panel schedules that check themselves.</b><br>
Macro-free Excel panel schedules, generated from Python and tested before every release.</p>

---

panelbook builds an Excel workbook of electrical panel schedules. Engineers and techs fill in yellow input cells;
everything else (phase loading, demand, line currents, cable sizes, checks) is calculated by ordinary Excel
formulas. No macros, so it opens the same in desktop Excel, Excel online, on a Mac and in LibreOffice.

The printed schedule keeps the familiar company layout, ready to paste into AutoCAD.

## What it does

- **Schedules in the company layout:** CKT#, description, trip, poles, breaker type, wire size, L-1/L-2/L-3 VA,
  totals, load classification with demand factors, MCA.
- **42- and 84-circuit sheets** (*Template* and *Template 84*), 12 to 84 circuits.
- **Demand by load class:** constant, tiered (NEC 220.42/220.44), quantity (220.56) and largest-motor +25%
  (430.24), applied to each panel's totals including its sub-panels.
- **Sub-panels:** pick a panel in any slot; its load comes in per phase and rolls up. The *Panels* register
  and *Topology* tree show the whole distribution system.
- **Exact line currents** for 1-, 2- and 3-pole loads on every system. Panel current is the highest line.
- **High-leg delta (120/240 and 115/230 V):** selectable high-leg phase, 1-pole loads blocked on it, orange
  rows, NEC 110.15 / 408.3 marking note, 240.85 breaker warning.
- **Transformer-fed sub-panels** with voltage-mismatch checks.
- **Cable and EGC selection** from NEC 310.16 and 250.122 with ambient and bundling factors, termination
  limits, 240.4(D) and parallel sets. Per-circuit override.
- **Checks on every circuit and panel:** ERR = must fix, WARN = check.

## Using it

Download `Panel_Schedule_Tool.xlsx` from the latest **Release** and start on its *Read Me* sheet.
Copy *Template* (or *Template 84*) for each new panel and list it on the *Panels* sheet.

## Status

**v2.0.0 baseline.** In use for review on a live project. Fixes in progress:

| Area | Change |
|---|---|
| ⚠️ Cable sizing | Derating always starts from the 90 °C column. **Wrong for 75 °C cable when ambient or bundling factors apply** (can undersize, e.g. #12 at 35 °C with 4–6 conductors). Unaffected when the factor is 1.0 (30 °C, ≤ 3 conductors). Fix: separate insulation rating (75 °C default, 90 °C explicit), 75/60 °C ambient columns, 125% continuous-load check. |
| Breakers | 80% / 100% rated breakers, continuous-load check, schedule footnote. |
| Schedule | Multi-pole rows shown merged; system box spans all VA columns. |
| Overrides | VA override + reason columns, "overrides in effect" log, detection of values typed onto the schedule. |
| Warnings | Quieter high-leg 2P (240.85) warning. |
| Workflow | Row and tab reset, working/print views, CAD Data Link range, blank starter workbook. |
| Assurance | Source-cited tables, rule tests, independent reference model, approved-panel tests, in-workbook self-test. |

## Building from source

Needs Python 3.10+ and LibreOffice (used headless to recalculate).

```bash
pip install -r requirements.txt
python panelbook/make.py          # -> dist/Panel_Schedule_Tool.xlsx
pytest panelbook/tests -v         # regression tests
```

Never hand-edit the workbook: change `panelbook/src/` and rebuild. Design notes and verification record:
[`panelbook/docs/design-notes.md`](panelbook/docs/design-notes.md).

```
panelbook/
  src/        workbook generator (layout, lookups, panel sheet, register/topology, assembly)
  tests/      regression panels and checks
  docs/       design notes, images
  make.py     build → recalc → finished workbook
shared/       LibreOffice recalc and cached-value injection
```

## Releases

Tag `panelbook-vX.Y.Z` and push. GitHub Actions runs the tests, builds the workbook and attaches it to a
Release with the changelog.

## Responsibility

panelbook is a design aid. Results must be checked by a qualified engineer against the adopted code edition
and manufacturer data; the engineer of record remains responsible for the design.

© 2026 Shankar Ramharack. All rights reserved.
