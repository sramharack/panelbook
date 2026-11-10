# Changelog: panelbook

## v2.1.0 (2026-11-10)
**Cable sizing (correctness fix).** v2.0 always derated from the 90 °C column, which undersizes 75 °C cable
whenever ambient or bundling factors apply (e.g. #12 on 20 A at 35 °C with 4–6 conductors).
- Insulation rating is now the design basis: Cu/Al 75C (default, any NEC building wire) or 90C (explicit,
  printed on the schedule). Derating uses the insulation column with 60/75/90 °C ambient factors
  (310.15(B)(1)(1)); the result is capped at the lower of termination and insulation columns (110.14(C)).
- Every cable passes four tests: (a) 125% continuous + noncontinuous at the cap column, (b) derated
  ampacity ≥ load, (c) derated ampacity ≥ breaker (240.4(B) optional, off), 240.4(D). The governing test is
  shown with each cable ([OCPD], [125%], [derated load], [min size], [override]).
- Continuous-load flag per load class (Lookups column P). Panel "Cable basis" check; WARN when the combined
  derating factor is below 0.80 (company practice, Lookups setting).
- Wire overrides below 1/0 stay one set and are checked against the tests.

**Breakers.** Standard 80% / MCB 100% / All 100% per panel (R7), "100% rated" per circuit. NEC 210.20(A)/215.3
breaker check (125% continuous, 100% when 100% rated), NEC minimum MCB in the panel checks (ERR below it),
"(100%)" after the MCB rating, breaker footnote on the schedule.

**Overrides.** VA override + reason columns per circuit: flows through totals, amps, demand, breaker and
cable checks; row WARN; "Overrides in effect" list; count in Circuit flags and on the Panels sheet. Values
typed over the schedule itself are detected and flagged ERR.

**Schedule look.** 2P/3P breakers read as one merged block (text on the first row of a 2P, middle row of a
3P). "Three Phase" box spans all VA columns. Conductor and breaker footnotes. Empty slots read SPARE.

**Workflow.** Working/print view (column outline 1/2). Reset a row (select its inputs, Delete) or a panel
(Name Box → Circuit_Inputs). Named range "Schedule" per sheet for AutoCAD Data Link. Version in footers and
Read Me. Blank starter workbook (`make.py --blank`). `tools/migrate.py` carries v2.0 workbooks forward.

**Warnings.** High-leg 2P (240.85) warning skips spares/unloaded breakers and "Straight 240V" breakers.

**Compatibility.** Odd-side inputs keep their columns (C:O); even-side inputs moved two columns right; the
hidden parameter column moved from BL to BQ. Use `tools/migrate.py` rather than copy-paste for old files.
Typed-over detection uses ISFORMULA (Excel 2013+).


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
