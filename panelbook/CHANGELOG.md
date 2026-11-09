# Changelog: panelbook

## Unreleased
- Planned for v2.1: see the Status table in the main README.


## v2.0.0 (2026-11-02)
- Printed schedule in the company v005 layout (BRK TYPE, WIRE SIZE, L-1/L-2/L-3 VA pairs, TOTAL AMPS, MCA).
- 42- and 84-circuit sheets (Template, Template 84); No. of Circuits 12-84.
- 120/240 and 115/230 V high-leg delta: selectable high-leg phase, 1-pole loads blocked/ERR, 240.85 WARN, orange rows,
  schedule note.
- Line currents per circuit type (1P VA/V(L-N), 2P VA/V(L-L), 3P VA/(√3·V(L-L))); panel current = max line.
- Transformer-fed sub-panels (Via XFMR), voltage-mismatch checks, Topology shows the voltage change.
- Cable selection: NEC 310.16, ambient and bundling factors, 110.14(C), 240.4(D), parallel sets, 250.122 EGC,
  wire override. Breaker type column.
- Parameter/export column moved from AX to BL.

## v1.0.0 (2026-10-28)
- Revit-style macro-free schedule, sub-panels, demand classes, Panels register, Topology.
