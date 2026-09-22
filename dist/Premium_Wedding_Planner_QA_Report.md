# Premium Wedding Planner — QA Report

Build method: generated programmatically with Python + `openpyxl` (`scripts/main.py` and
supporting modules in `scripts/`), recalculated and verified with LibreOffice headless
(`soffice`) running as the calculation engine.

## Workbook

| File | Sheets | Formulas | Charts | Excel Tables | Data-validation rules |
|---|---|---|---|---|---|
| `Premium_Wedding_Planner_DEMO.xlsx` | 18 (17 visible + 1 hidden helper) | 920 | 8 | 12 | 29 |
| `Premium_Wedding_Planner_BLANK.xlsx` | 18 (17 visible + 1 hidden helper) | 920 | 8 | 12 | 29 |

Sheet order: **START HERE, Dashboard, Budget, Vendors, Guest List, Seating Plan, Timeline,
Master Checklist, Payments, Wedding Party, Venue Comparison, Food & Drinks, Honeymoon Budget,
Contacts, Gift Tracker, Calendar, Settings**, plus a hidden **Lists** sheet holding every
dropdown's source list (RSVP status, payment status, task status, priority, budget category,
meal preference, etc.) and a hidden chart-data helper block on the Dashboard sheet (category
subtotals, RSVP/meal/period roll-ups) so no calculation is duplicated by hand.

## Formula QA

- Formulas checked: all 920 formulas in each workbook, via two full passes:
  1. **Automated recalculation** — LibreOffice headless recalculated every formula and
     reported cell-level errors. Both DEMO and BLANK returned `"status": "success"`,
     **0 errors** (`#REF!`, `#VALUE!`, `#DIV/0!`, `#NAME?`, `#N/A` all absent), confirmed by
     scanning every cached cell value afterward for an error-string pattern (zero hits in
     either file).
  2. **Manual reconciliation** — every Dashboard KPI, every per-sheet summary card, and every
     chart's underlying helper data was cross-checked against the source data in
     `scripts/demo_data.py` with an independent Python calculation (not the workbook's own
     formulas). This caught and fixed a real set of column-mapping bugs the automated pass
     alone would not have flagged as "errors" (wrong-but-valid formulas, e.g. summing the
     wrong column), including: Vendors/Payments/Honeymoon "Remaining" formulas pointing at
     the wrong columns, the Checklist/Payments/Gift Tracker status columns being off by one,
     a `COUNTA(range)-COUNTBLANK(range)` pattern that silently undercounted every "total
     invited / total vendors / total tasks / gifts received" KPI, and the Food & Drinks menu
     total summing the Price column instead of the computed Total column. All were fixed and
     re-verified; final reconciliation matched Python-computed ground truth exactly for
     Budget, Vendors, Guest List, Payments, Master Checklist, Honeymoon Budget, Food & Drinks,
     Gift Tracker, and every Dashboard chart's helper series (budget-by-category, RSVP status,
     meal preference, checklist-by-period, top-5 expenses).
- No volatile or unsupported functions used (no `XLOOKUP`/`FILTER`/`SORT`/`UNIQUE`); lookups
  use `INDEX`/`MATCH`/`LARGE`, aggregation uses `SUMIF`/`SUMIFS`/`COUNTIF`/`COUNTIFS`, all
  guarded with `IF(...="","",...)` or `IFERROR(...)` so blank input never produces an error.

## Dashboard

- **KPIs** (18 cards across 4 sections): Total Budget, Actual Spending, Remaining Budget,
  Amount Paid, Outstanding (Financial Snapshot); Total Guests, Confirmed, Pending RSVPs, Not
  Attending, Vendors Booked (Guest Snapshot); Checklist Complete %, Tasks Remaining,
  Outstanding Vendor Payments, Honeymoon Budget, Honeymoon Remaining (Planning & Honeymoon);
  Overdue Payments, Pending RSVPs, Overdue Tasks, Over-Budget Categories (Attention Needed).
  Every KPI is a live formula referencing its source sheet — none are hardcoded.
- **Charts** (8, all 2-D, wedding-palette colored, no default Excel rainbow colors): Budget
  vs. Actual by Category (column), Expenses by Category (doughnut), Paid vs. Remaining (bar),
  Top 5 Wedding Expenses (bar), RSVP Status (doughnut), Guest Meal Preferences (column),
  Checklist Completion by Period (stacked column), Honeymoon: Planned vs. Actual (column).
- **Source validation**: every chart and KPI was traced to its source range and reconciled
  against demo data as described above; nothing is a static/decorative number.

## Data validation (dropdowns)

29 list-based dropdowns checked across Budget (Category, Payment Status), Vendors (Category,
Contract Status, Payment Status, Rating), Guest List (Side, Invitation Sent, RSVP Status, Plus
One, Meal Preference, Children Attending), Master Checklist (Period, Priority, Owner, Status),
Payments (Payment Status, Payment Method), Venue Comparison (Yes/No fields, Rating), Wedding
Party (Side), Food & Drinks (Category), Honeymoon Budget (Category), Contacts (Category), Gift
Tracker (Gift Type, Thank You Sent). Every dropdown's source range was confirmed non-broken
(all point at populated columns on the hidden Lists sheet).

## Conditional formatting

Status-color rules checked on: Budget (payment status + over-budget highlight), Vendors
(contract status + payment status), Guest List (RSVP status), Payments (payment status +
overdue-due-date highlight), Master Checklist (task status + overdue-due-date + high-priority
highlight), Gift Tracker (thank-you sent), Calendar (task/payment status roll-ups). All use the
sage/blush/gold/deep-rose palette rather than default red/yellow/green icon colors, and none
fired unexpectedly against the blank template (checked directly — blank rows show no fill).

## Visual QA

Checked on every sheet: consistent blush/ivory header banner with gold accent rule, Calibri
throughout, currency/date/percent number formats applied consistently, alternating row bands,
freeze panes on every long data sheet (Budget, Vendors, Guest List, Master Checklist,
Payments, Contacts, Gift Tracker, Seating Plan roster), auto-filter enabled via native Excel
Tables (12 tables total), optimized column widths (no default-width columns on data sheets),
comfortable row heights with text wrap on long fields (Item/Description, Notes,
Responsibilities). Dashboard KPI cards render as bordered cards with a colored top accent
rather than plain cells.

## Print QA

Every data sheet has an explicit print area, landscape orientation (portrait for
Settings/Calendar/START HERE), "fit to 1 page wide," a repeating header row via
`print_title_rows`, and 0.4–0.5" margins. Not test-printed to a physical/PDF renderer in this
session — see Remaining Limitations.

## Demo data

Fictional couple **Sophia Bennett & Daniel Whitfield**, wedding **15 June 2027**, Napa Valley,
CA. 58 guests, 18 vendors, 24 budget line items across all 20 categories, 48 checklist tasks
across all 9 planning periods, 17 payment records, 26 gift records, a 12-step wedding-day
timeline, and an 8-category honeymoon budget — all internally consistent (e.g., Budget planned
total $67,900 vs. actual $66,305 against a $70,000 total budget setting, so the Dashboard shows
a realistic "under budget" state rather than an arbitrary mismatch). Verified connected to the
Dashboard as described in Formula QA.

## Blank template

Confirmed after stripping demo data: all 920 formulas remain, all 12 tables and 29 dropdowns
remain, all 8 charts remain (rendering as empty/zero rather than erroring), conditional
formatting remains, recalculation is clean (0 errors), and a scan of every cached cell value
found zero `#REF!`/`#VALUE!`/`#DIV/0!`/`#NAME?`/`#N/A` occurrences. Countdown, currency, and
percentage KPIs display a blank dash/em-dash or $0 rather than an error when the corresponding
Settings field is empty.

## Remaining limitations

- **Print/PDF output was not visually rendered** in this session (no ability to open a PDF
  preview here); print areas, orientation, and scaling were set programmatically and are
  standard Excel/LibreOffice features, but a manual "Print Preview" pass in Excel is
  recommended before publishing.
- **Font rendering and exact pixel-level spacing** (row heights, card alignment) were not
  visually inspected in an actual Excel/Excel-compatible GUI — only structurally verified via
  openpyxl. If you spot any spacing issue on first open, it's most likely a row-height or
  merged-cell nuance rather than a formula problem.
- **This session's sandbox was initially missing the `libreoffice-calc` package** (only
  `libreoffice-core` was preinstalled), which blocked all recalculation until it was installed
  mid-session — recorded here for transparency since a large part of the QA process (see
  Formula QA above) depended on getting that working.
- Conditional-formatting color rendering and data-bar/icon-set visuals were verified by
  formula/rule inspection, not by visual screenshot, since this session has no Excel GUI.
