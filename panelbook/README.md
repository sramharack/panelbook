# panelbook: source

| File | Role |
|---|---|
| `src/common.py` | layout, column maps, parameter addresses, styles |
| `src/lookups.py` | Lookups sheet: systems, load classes, ampacity / EGC / correction tables, settings |
| `src/panel.py` | one panel sheet: schedule, input zones, hidden calc zone, checks |
| `src/summary.py` | Panels register and Topology |
| `src/build.py` | example panels, templates, Read Me, workbook assembly |
| `tests/test_panels.py` | regression panels (never shipped) |
| `tests/test_regression.py` | checks: high-leg line amps, LP000 vs Revit, cable picks, edge cases |
| `make.py` | build → LibreOffice recalc → inject cached values → `dist/` |

See [`docs/design-notes.md`](docs/design-notes.md) for how the workbook is organised.
