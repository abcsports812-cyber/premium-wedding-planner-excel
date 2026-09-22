# Premium Wedding Planner — Final QA Report (Post-Fix Pass)

This report supersedes the prior QA report. It documents the fixes applied for all 4 FAILs
and 5 WARNINGs raised in the Final Visual QA Report, and the full re-verification performed
afterward. No commit/push has been made — this report is for review before that step.

Build method unchanged: Python + `openpyxl` generator (`scripts/`), recalculated with
LibreOffice headless (`soffice`) as the calculation engine, both DEMO and BLANK built from the
same shared builder (`scripts/main.py` → `build_workbook(demo=True/False)`).

---

## Required explicit confirmations

| Check | Result |
|---|---|
| Chart overlap | **ZERO** — verified mathematically from the saved XLSX's chart anchor XML (all 8 charts, pairwise) |
| Formula errors | **ZERO** in both DEMO and BLANK (LibreOffice recalculation + full cached-cell scan for `#REF!/#VALUE!/#DIV/0!/#NAME?/#N/A/#NULL!/#NUM!`) |
| KPI reconciliation | **PASS** — all 17 Dashboard KPI cells matched independent Python-computed ground truth exactly |
| Dynamic propagation | **PASS** — live-tested on a temporary copy, discarded afterward |
| Print setup | **PASS** — all 17 visible sheets now have an explicit `print_area` and `fitToPage=True` |
| Header clipping | **PASS** — zero headers exceed their column's effective width, workbook-wide |
| Wrapped-text clipping | **PASS** — row height now grows automatically wherever wrapped content needs 2+ lines |
| Dropdown validation | **PASS** — all 29 rules point at correctly-sized, correctly-typed Lists ranges |

---

## PASS (verified this pass)

| # | Item | Evidence |
|---|---|---|
| 1 | Zero formula errors, both files | Full cached-cell scan, 0 hits, both DEMO and BLANK |
| 2 | Zero chart overlap | Extracted every chart's two-cell anchor from the saved file and pairwise-tested bounding boxes — 0 intersections (see **F1** below for the numbers) |
| 3 | All 8 charts have valid, non-broken `val`/`cat` references after LibreOffice recalculation | Checked directly against the saved chart XML |
| 4 | Dashboard KPI reconciliation | All 17 KPI cells (Financial Snapshot, Guest Snapshot, Planning & Honeymoon, Attention Needed) matched independently-computed Python totals exactly |
| 5 | Dynamic propagation | Raised `Budget!G6` by $15,000 and changed `Guest List!I6` Attending→Declined on a temp copy; Actual Spending, Remaining Budget, Confirmed, Not Attending, and Over-Budget Categories all updated by the exact expected deltas after recalculation; charts retained valid data references afterward; temp copy discarded |
| 6 | Dropdown validation | All 29 rules re-verified against the Lists sheet — correct range, correct size, non-empty |
| 7 | Rating value/list type consistency | `Lists` Rating list and `Vendors!O`/`Venue Comparison!M` cell values are now uniformly integers |
| 8 | Header clipping | Zero headers exceed their (now auto-widened) column width, checked across all 12 table sheets |
| 9 | Wrapped-text row height | Every checked wrap cell that needs 2 lines now gets a 33.75pt row (vs. the fixed 19.5pt before) |
| 10 | Print setup | All 17 visible sheets have `print_area` + `fitToPage=True`; Lists (hidden) is correctly excluded from printing entirely |
| 11 | Blank template has no leaked demo data | 0 leaked cells found across all 12 table sheets (names, vendors, addresses, etc.) |
| 12 | Demo data still realistic/consistent | Sophia & Daniel, 15 June 2027; Budget planned $67,900 / actual $66,305 vs. $70,000 setting — unchanged, still internally consistent |
| 13 | Structure preserved | Both files: 18 sheets (17 visible + hidden Lists), 920 formulas, 8 charts, 12 Excel Tables, 29 dropdowns — identical counts to before the fix pass; no feature removed |

---

## FAILS — fixed

### F1 — Chart overlap (was: all 8 charts overlapping)
**Fix:** Redesigned the chart grid in `scripts/dashboard.py`. Charts are now a uniform 15.0 cm × 8.5 cm, anchored at column B (left) and column M (right) — a 10-column gutter (~19.4 cm) that comfortably exceeds the 15 cm chart width — and each chart-pair row is stepped by 20 rows (17 rows of actual chart height + 3-row gap), computed from the chart height rather than a fixed guess.

**Verified from the saved file's actual anchors** (post-recalculation):
```
Budget vs. Actual by Category   cols[1-8]   rows[27-43]
Expenses by Category            cols[12-19] rows[27-43]
Paid vs. Remaining               cols[1-8]   rows[47-63]
Top 5 Wedding Expenses           cols[12-19] rows[47-63]
RSVP Status                      cols[1-8]   rows[67-83]
Guest Meal Preferences           cols[12-19] rows[67-83]
Checklist Completion by Period   cols[1-8]   rows[87-103]
Honeymoon: Planned vs. Actual    cols[12-19] rows[87-103]
```
Pairwise intersection test across all 28 chart pairs: **0 overlaps**, both horizontally (columns 9–11 are a clear gutter) and vertically (each row-pair has a clear gap before the next).

**A second, unplanned bug was found and fixed during this work:** LibreOffice's recalculation was silently dropping the `<c:cat>`/`<c:val>` XML elements (emptying the chart) for 3 of the 8 charts specifically when both their value and category data lived on the chart's own host sheet (Dashboard). This was invisible to the "0 formula errors" check because it's a chart-rendering defect, not a formula error — caught only by inspecting the saved chart XML directly. Fixed by moving all Dashboard chart helper data onto the (already-hidden) `Lists` sheet, matching the pattern already proven safe by the 5 charts that referenced Lists/other sheets and never broke. All 8 charts now retain valid data after recalculation, confirmed above.

### F2 — Wrapped text / row height clipping
**Fix:** Generic fix in `scripts/tablesheet.py`, not per-cell: every data row now computes, for each `wrap=True` column, how many lines the actual cell text needs at that column's width (using a conservative characters-per-width-unit estimate), and sets the row height to the max across the row (minimum 20pt, ~14pt per line + padding for more). Blank/formula-only rows are unaffected (stay at the compact default). Verified on all 6 originally-flagged examples — all now get a 33.75pt (2-line) row instead of the fixed 19.5pt.

### F3 — Clipped headers
**Fix:** Generic fix in `scripts/tablesheet.py`: every column's width is now `max(declared_width, header_text_length + 2)` at generation time, applied to every table sheet, not just the 5 flagged cells. Verified: `Guest List!M5/N5` and `Venue Comparison!I5/J5/L5` are all now wide enough (widths 22/20/19/20/15 respectively), and a workbook-wide scan found zero remaining header/column-width mismatches.

### F4 — Dashboard print area / page setup
**Fix:** Dashboard now has `print_area = A1:U107` (covers the banner, all 18 KPI cards, and all 8 charts), `fitToPage=True`, `fitToWidth=1`, and `fitToHeight=0` (fit width to one page, let height paginate naturally — deliberately not forced to one page, since squeezing an 18-KPI, 8-chart dashboard onto a single sheet would make it illegibly small; this matches the instruction to prioritize readable output). The old hidden helper columns on Dashboard no longer exist at all (moved to the Lists sheet, see F1), so there is nothing to accidentally include in the print area.

---

## WARNINGS — fixed

### W1 — Print settings on Calendar / Settings / START HERE
**Fix:** All three now have an explicit `print_area`, `fitToPage=True`, `fitToWidth=1`. Calendar (94 rows, two sections) additionally gets `print_title_rows` repeating **both** section header rows ("Task Deadlines" and "Payment Deadlines") on every printed page, and `fitToHeight=0` so it paginates naturally instead of being squeezed. START HERE and Settings, being short, use `fitToHeight=1` (comfortably one page).

### W2 — Seating Plan freeze pane
**Fix:** Reviewed and improved rather than left as-is. The Table Overview reference block was restructured from one 20-row column into two side-by-side 10-row columns (Tables 1–10, Tables 11–20), directly shrinking the section the freeze pane has to sit below. The freeze pane still anchors at the guest roster's own header row (the practically-necessary choice — Excel freeze panes can't skip the middle of a sheet), but it moved from **B31 to B21**, a 10-row (one-third) reduction in the frozen/non-scrollable region, while keeping the full 20-table capacity reference visible above the roster.

### W3 — Dashboard KPI card spacing
**Fix:** Redesigned the KPI grid in `scripts/dashboard.py` from 5 cards touching edge-to-edge (`B,E,H,K,N`, zero gap) to 5 cards with a genuine 1-column gutter between each (`B,F,J,N,R` — verified via the saved merged-cell ranges: `B8:D9, F8:H9, J8:L9, N8:P9, R8:T9`). This required widening the Dashboard's visible content area from 15 to 19 columns (B→P to B→T) — a deliberate, modest increase in service of the explicit ask, not an incidental one; card widths themselves are unchanged (still 3 columns, still comfortably fit the longest label, "Outstanding Vendor Pmts" / "Over-Budget Categories") so no label overflow risk was introduced. All 18 KPI cards remain; none were removed or resized down.

### W4 — Rating dropdown data type
**Fix:** Changed `Lists["Rating"]` from `["1","2","3","4","5"]` (text) to `[1,2,3,4,5]` (integer) in `scripts/lists_settings.py`. Demo cell values in `Vendors!O` and `Venue Comparison!M` were already integers, so list and stored values are now consistently numeric. No formula or chart references the Rating column, so nothing else needed updating.

### W5 — Checklist chart labels
**Fix:** Added a chart-only short-label column on the Lists sheet (`"12+ Months", "9-12 Months", "6-9 Months", "3-6 Months", "1-3 Months", "Final Month", "Wedding Week", "Wedding Day", "Post-Wedding"`), positionally mirrored 1:1 to the real `Lists!N` (ChecklistPeriod) order and used only as the "Checklist Completion by Period" chart's category axis. The underlying `Master Checklist!Period` column and the `Lists!ChecklistPeriod` dropdown source are completely unchanged — still the full "12+ Months Before" etc. text everywhere a human reads or selects it.

---

## Remaining limitations (unchanged from before, still honest)

- Print/PDF output and exact pixel-level rendering (chart legibility at 15×8.5cm, row-height visual comfort, card gutter appearance) were not visually rendered in this session — verified structurally (geometry math against the saved XML, character-length-vs-width heuristics) rather than by screenshot, since no Excel/LibreOffice GUI is available here. A manual look in Excel before publishing is still recommended.
- The LibreOffice same-sheet-category-reference chart bug (see F1) was diagnosed empirically against this specific file/LibreOffice version (24.2.7); it's avoided architecturally now (all Dashboard chart helper data lives on a different sheet than its chart), which is also simply a cleaner design, but the underlying LO behavior itself wasn't root-caused to a specific mechanism.
- Chart category-label crowding on "Checklist Completion by Period" (9 categories) is improved by shorter labels (W5) but was not visually confirmed to be free of rotation/overlap, since that requires rendering.

---

## Not yet done (per your instructions)
Commit and push are intentionally **not** performed. Everything above reflects the current state of `dist/Premium_Wedding_Planner_DEMO.xlsx` and `dist/Premium_Wedding_Planner_BLANK.xlsx` on disk, both freshly regenerated from the shared builder and recalculated. Awaiting your approval before committing/pushing.
