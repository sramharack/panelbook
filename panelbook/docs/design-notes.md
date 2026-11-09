# Panel Schedule Tool (Excel, no macros): design notes

Deliverable: `Panel_Schedule_Tool.xlsx`, v2 built 2026-11-02 (v1 2026-10-28). It replaces the macro-based panel workbook (NEW SHEET / IMPORT SUB PANEL / COPY buttons). Everything runs on lookup tables, INDEX/MATCH and INDIRECT; Excel 2007 functions only.

## v2 changes (2026-11-02)
- **Company look:** the print area now follows the v005 sheets: header block, CKT# | DESCRIPTION | TRIP | POLES | BRK TYPE | WIRE SIZE | L-1/L-2/L-3 VA (odd/even pairs) | mirror, TOTAL LOAD VA ×2 / TOTAL AMPS, load classification + totals block with MCA. The Revit print view was dropped.
- **Bigger boards:** sheet capacity is set at build time: `Template` (42) and `Template 84` (84). No. of Circuits 12–84.
- **High-leg delta:** `120/240 3Ph 4W High-Leg Delta` and `115/230 3Ph 4W High-Leg Delta` in Lookups, new column G "High-leg phase" (default B). Per-panel override (A/B/C; not-B = WARN, 408.3(E)(1)).
  - 1P load on a high-leg row: blocked by data validation on entry (set Poles first) and ERR.
  - 1P breaker with no load on high leg: WARN. 2P including high leg: WARN (straight 240 V breaker, 240.85).
  - Orange rows in input zones only; schedule note "HIGH LEG: L-2 (B) = 208V TO N" and "L-2 VA (HIGH LEG)" header.
  - Phase imbalance on high-leg panels compares A and C only.
- **Line currents:** per pole VA × k: k1 = 1/V(L-N), k2 = 2/V(L-L), k3 = √3/V(L-L), summed per line. All systems. Direct sub-panel feeders use the sub's own line currents. Panel current = MAX(line amps); demand A = that × demand/connected. Used by loading, Panels, Topology.
- **Transformer feeders:** "Via XFMR" input column. Sub VA spread evenly over the feeder poles, current at this panel's voltage. Without XFMR: V mismatch = WARN; neutral-load sub on a feeder that includes the high leg = ERR. Topology shows "XFMR: 120/240 … → 120/208 Wye", or ⚠ if voltages differ with no XFMR.
- **Cable selection:** NEC 310.16 (Cu/Al, 60/75/90 °C), 90 °C ampacity × ambient (310.15(B)(1)) × bundling (310.15(C)(1)) factor, capped at the termination column (110.14(C)) and 240.4(D). Smallest size with allowable ≥ trip; above `LK_MaxWire` (500 kcmil) it goes parallel (1/0 min). EGC from 250.122, full size per set, never larger than the phase conductor. Per-panel material / termination / ambient / CCC; defaults on Lookups. Per-circuit "Wire override". Panel feeder (header WIRE SIZE) sized from MCB, else bus.
- **New input columns:** Via XFMR, Breaker Type (Lookups list), Wire override. New outputs: Cable (auto).
- Empty-slot label default is now SPACE (Lookups setting).
- MCA = next standard breaker ≥ est. demand A (fixes the #REF! in v005).

## Sheets
- **Read Me**, **Panels** (register), **Topology** (tree + KPIs), panel tabs (HDP, LP000, LP001, LP001A, RMDP, LP-R1, Template, Template 84), **Lookups**.
- Panel tabs: A:S odd inputs, U:AN schedule (print area), AP:BH even inputs, hidden calc zone from BK.
- **Parameters/exports moved from AX to BL** (BL1:BL70; rows 1–41 keep their v1 meaning except aA/aB/aC = line amps, connA = max line A, demA = max line A × DF). Rows 42+ are new (hl, vhl, amax, mat, fct, term, fcab, mca, neut, cap…).
- Hidden class block rows 10–29, feeder list rows 10–25 (16 feeders) are fixed regardless of capacity.

## Key conventions
- Circuit load VA = Fx1 Qty × VA + Fx2 Qty × VA + Added VA.
- Phase per row: CKT 1/2 = A (L-1), 3/4 = B (L-2), 5/6 = C (L-3). 1-phase 3-wire alternates A/B.
- Multi-pole breakers on their first CKT; the next slot(s) on the same side become "(pole of CKT x)".
- Tab name must equal Panel Name (D2). Never insert or delete rows/columns on panel sheets.
- Totals are arithmetic VA sums (LP000 12,620 VA vs Revit 12,470 VA vector sum).

## Verification (2026-11-02)
- 41,740 formulas, 0 errors after LibreOffice recalc.
- High-leg regression (`tests/test_panels.py`, HLTEST): A 35.27 / B 31.10 / C 21.10 A, max 35.27. 1P on B = ERR; 2P on A–B = 240.85 WARN.
- LP000: Lighting - Dwelling Unit 4,860 VA → 75.12%; total 12,620 VA.
- Edge cases (`tests/test_panels.py` E1–E5): high leg on C, 115/230 (199 V), XFMR missing, neutral sub on high leg, Al / 40 °C / 6 CCC feeder (350 kcmil at 200 A), overrides.

## Build pipeline
- `python panelbook/make.py`: openpyxl build → LibreOffice recalc (`shared/recalc.py`) → cached values copied back (`shared/inject.py`) → `dist/Panel_Schedule_Tool.xlsx`. `fullCalcOnLoad` stays on.
- `pytest panelbook/tests -v`: builds with the regression panels (`tests/test_panels.py`) and checks the numbers above.
- Tag `panelbook-vX.Y.Z` and push: GitHub Actions runs the tests and attaches the workbook to a Release.

## Known limits / next steps
- No voltage drop yet: use Wire override; EGC is not upsized for it (250.122(B)), 240.4(B) next-size-up not applied.
- Ambient / CCC are per panel, not per circuit. Neutral counted for 1P circuits and neutral sub-panels only (not 120/240 V 2P loads).
- Rows beyond No. of Circuits must be hidden by hand to shorten the print.
- Per-phase demand not computed (demand A scales the max line current by panel DF).
