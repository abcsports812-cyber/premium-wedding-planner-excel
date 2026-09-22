from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from styles import (
    set_sheet_defaults, F_TITLE, F_SUBTITLE, F_SECTION, F_BODY, F_KPI_LABEL, kpi_card,
    FMT_CURRENCY0, FMT_PCT1, FMT_INT, FMT_DATE, col_idx, DEEP_ROSE, SAGE, GOLD, BLUSH,
    LIGHT_BLUSH, LIGHT_SAGE, MUTED_GOLD, CHARCOAL, DUSTY_ROSE, IVORY, FILL_LIGHT_BLUSH,
    fill, F_NAV,
)
from lists_settings import LISTS

LIST_COL = {k: i + 1 for i, k in enumerate(LISTS.keys())}

PALETTE = [DUSTY_ROSE, SAGE, GOLD, "D9B8B4", "C7D2C0", MUTED_GOLD, "9D646B", "715B52"]

# ------------------------------------------------------------------ LAYOUT --
# Empirical: a Dashboard column at width COL_WIDTH renders at ~CM_PER_COL cm
# (measured from LibreOffice's own recalculated two-cell chart anchors).
COL_WIDTH = 8.2
CM_PER_COL = 1.9375

CARD_COLS = 3          # columns a KPI card spans
GUTTER_COLS = 1         # blank spacer column between cards
CARD_STRIDE = CARD_COLS + GUTTER_COLS
N_CARDS_PER_ROW = 5
FIRST_CARD_COL = 2      # column B
CARD_START_COLS = [FIRST_CARD_COL + i * CARD_STRIDE for i in range(N_CARDS_PER_ROW)]
LAST_CONTENT_COL = CARD_START_COLS[-1] + CARD_COLS - 1     # column T (20)
PRINT_RIGHT_COL = LAST_CONTENT_COL + 1                      # 1-col safety margin -> U (21)

CHART_WIDTH_CM = 15.0
CHART_HEIGHT_CM = 8.5
CHART_ROWS = -(-int(CHART_HEIGHT_CM * 100) // int(0.529 * 100))  # ceil(height / default-row-cm)
CHART_ROW_GAP = 3
CHART_ROW_STEP = CHART_ROWS + CHART_ROW_GAP
CHART_LEFT_COL = FIRST_CARD_COL                              # B
CHART_GAP_COLS = -(-int(CHART_WIDTH_CM * 100) // int(CM_PER_COL * 100)) + 3  # width + generous gutter
CHART_RIGHT_COL = CHART_LEFT_COL + CHART_GAP_COLS             # L

# Hidden chart-data helper block lives well clear of all visible content/print area.
HELPER_START = LAST_CONTENT_COL + 3   # 2-col buffer past the print margin
H_PLANNED = HELPER_START
H_ACTUAL = HELPER_START + 1
H_PAIDREM_LABEL = HELPER_START + 2
H_PAIDREM_VALUE = HELPER_START + 3
H_TOP5_VAL = HELPER_START + 4
H_TOP5_LABEL = HELPER_START + 5
H_RSVP = HELPER_START + 6
H_MEAL = HELPER_START + 7
H_CHK_COMPLETED = HELPER_START + 8
H_CHK_REMAINING = HELPER_START + 9
H_SHORT_PERIOD = HELPER_START + 10
HELPER_END = H_SHORT_PERIOD

SHORT_PERIOD_LABELS = [
    "12+ Months", "9-12 Months", "6-9 Months", "3-6 Months", "1-3 Months",
    "Final Month", "Wedding Week", "Wedding Day", "Post-Wedding",
]


def style_chart(chart, title, height=CHART_HEIGHT_CM, width=CHART_WIDTH_CM):
    chart.title = title
    chart.style = 2
    chart.height = height
    chart.width = width
    if hasattr(chart, "y_axis"):
        chart.y_axis.majorGridlines = None
    if hasattr(chart, "gapWidth"):
        chart.gapWidth = 60
    for i, series in enumerate(chart.series):
        color = PALETTE[i % len(PALETTE)]
        series.graphicalProperties.solidFill = color
        series.graphicalProperties.line.noFill = True
    return chart


def colorize_doughnut(chart, n):
    series = chart.series[0]
    pts = []
    for i in range(n):
        dp = DataPoint(idx=i)
        dp.graphicalProperties.solidFill = PALETTE[i % len(PALETTE)]
        pts.append(dp)
    series.data_points = pts


def build_dashboard(wb, refs, demo):
    """refs: dict with first/last rows for Budget, Vendors, Guest List, Payments,
    Master Checklist, Honeymoon Budget sheets."""
    ws = wb.create_sheet("Dashboard", 1)
    set_sheet_defaults(ws, GOLD)

    last_banner_col = LAST_CONTENT_COL

    # ---- Wedding identity banner ----
    for r in range(1, 6):
        for c in range(1, last_banner_col + 1):
            ws.cell(row=r, column=c).fill = FILL_LIGHT_BLUSH
    ws.row_dimensions[1].height = 6
    ws.row_dimensions[2].height = 16
    lbl = ws.cell(row=2, column=2, value="OUR WEDDING")
    lbl.font = Font(name="Calibri", size=11, bold=True, color=DEEP_ROSE)
    ws.row_dimensions[3].height = 34
    names = ws.cell(row=3, column=2, value='=IF(Partner1Name="","Partner One",Partner1Name)&"  &  "&IF(Partner2Name="","Partner Two",Partner2Name)')
    names.font = Font(name="Calibri", size=26, bold=True, color=CHARCOAL)
    ws.row_dimensions[4].height = 18
    datecell = ws.cell(row=4, column=2, value='=IF(WeddingDate="","Wedding date not yet set",TEXT(WeddingDate,"dddd, d mmmm yyyy"))')
    datecell.font = Font(name="Calibri", size=13, italic=True, color="715B52")

    countdown_left = last_banner_col - 3  # 4-column countdown box on the right edge
    for r in range(2, 5):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=countdown_left - 1)

    ws.merge_cells(start_row=2, start_column=countdown_left, end_row=5, end_column=last_banner_col)
    cd = ws.cell(row=2, column=countdown_left,
                 value='=IF(WeddingDate="","—",IF(WeddingDate>=TODAY(),WeddingDate-TODAY()&" DAYS TO GO",TEXT(WeddingDate,"d mmm yyyy")&" ❦ Married!"))')
    cd.font = Font(name="Calibri", size=17, bold=True, color=DEEP_ROSE)
    cd.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(2, 6):
        ws.cell(row=r, column=countdown_left).border = Border(left=Side(style="thin", color=GOLD))
    for c in range(1, last_banner_col + 1):
        ws.cell(row=5, column=c).border = Border(bottom=Side(style="medium", color=GOLD))
    nav = ws.cell(row=1, column=2, value="‹ Start Here")
    nav.font = F_NAV
    nav.hyperlink = "#'START HERE'!A1"

    B, V, G, P, C, H = refs["Budget"], refs["Vendors"], refs["Guest List"], refs["Payments"], refs["Master Checklist"], refs["Honeymoon Budget"]

    c0, c1, c2, c3, c4 = CARD_START_COLS

    row = 7
    sect = ws.cell(row=row, column=2, value="FINANCIAL SNAPSHOT")
    sect.font = F_SECTION
    row += 1
    cards_row = row
    kpi_card(ws, cards_row, c0, "Total Budget", "=TotalBudget", FMT_CURRENCY0, DEEP_ROSE)
    kpi_card(ws, cards_row, c1, "Actual Spending", f"=SUM(Budget!G{B[0]}:G{B[1]})", FMT_CURRENCY0, DUSTY_ROSE)
    kpi_card(ws, cards_row, c2, "Remaining Budget", f'=IF(TotalBudget="","",TotalBudget-SUM(Budget!G{B[0]}:G{B[1]}))', FMT_CURRENCY0, SAGE)
    kpi_card(ws, cards_row, c3, "Amount Paid", f"=SUM(Budget!J{B[0]}:J{B[1]})", FMT_CURRENCY0, SAGE)
    kpi_card(ws, cards_row, c4, "Outstanding", f"=SUM(Budget!K{B[0]}:K{B[1]})", FMT_CURRENCY0, GOLD)

    row = cards_row + 4
    sect = ws.cell(row=row, column=2, value="GUEST SNAPSHOT")
    sect.font = F_SECTION
    row += 1
    cards_row2 = row
    kpi_card(ws, cards_row2, c0, "Total Guests", f"=COUNTA('Guest List'!B{G[0]}:B{G[1]})", FMT_INT, DEEP_ROSE)
    kpi_card(ws, cards_row2, c1, "Confirmed", f"=COUNTIF('Guest List'!I{G[0]}:I{G[1]},\"Attending\")", FMT_INT, SAGE)
    kpi_card(ws, cards_row2, c2, "Pending RSVPs", f"=COUNTIF('Guest List'!I{G[0]}:I{G[1]},\"Pending\")", FMT_INT, GOLD)
    kpi_card(ws, cards_row2, c3, "Not Attending", f"=COUNTIF('Guest List'!I{G[0]}:I{G[1]},\"Declined\")", FMT_INT, "9D646B")
    kpi_card(ws, cards_row2, c4, "Vendors Booked", f"=COUNTA(Vendors!B{V[0]}:B{V[1]})", FMT_INT, DUSTY_ROSE)

    row = cards_row2 + 4
    sect = ws.cell(row=row, column=2, value="PLANNING PROGRESS & HONEYMOON")
    sect.font = F_SECTION
    row += 1
    cards_row3 = row
    kpi_card(ws, cards_row3, c0, "Checklist Complete", f"=IFERROR(COUNTIF('Master Checklist'!H{C[0]}:H{C[1]},\"Completed\")/(COUNTA('Master Checklist'!C{C[0]}:C{C[1]})),\"\")", FMT_PCT1, SAGE)
    kpi_card(ws, cards_row3, c1, "Tasks Remaining", f"=COUNTIFS('Master Checklist'!C{C[0]}:C{C[1]},\"<>\",'Master Checklist'!H{C[0]}:H{C[1]},\"<>Completed\")", FMT_INT, GOLD)
    kpi_card(ws, cards_row3, c2, "Outstanding Vendor Pmts", f"=SUM(Vendors!K{V[0]}:K{V[1]})", FMT_CURRENCY0, GOLD)
    kpi_card(ws, cards_row3, c3, "Honeymoon Budget", "=HoneymoonBudgetSetting", FMT_CURRENCY0, DEEP_ROSE)
    kpi_card(ws, cards_row3, c4, "Honeymoon Remaining", f"=SUM('Honeymoon Budget'!G{H[0]}:G{H[1]})", FMT_CURRENCY0, SAGE)

    # ---- Attention Needed ----
    row = cards_row3 + 4
    sect = ws.cell(row=row, column=2, value="ATTENTION NEEDED")
    sect.font = F_SECTION
    row += 1
    att_row = row
    kpi_card(ws, att_row, c0, "Overdue Payments",
             f'=COUNTIFS(Payments!H{P[0]}:H{P[1]},"<>",Payments!H{P[0]}:H{P[1]},"<"&TODAY(),Payments!I{P[0]}:I{P[1]},"<>Paid in Full")',
             FMT_INT, "9D646B")
    kpi_card(ws, att_row, c1, "Pending RSVPs",
             f'=COUNTIF(\'Guest List\'!I{G[0]}:I{G[1]},"Pending")', FMT_INT, GOLD)
    kpi_card(ws, att_row, c2, "Overdue Tasks",
             f'=COUNTIFS(\'Master Checklist\'!G{C[0]}:G{C[1]},"<>",\'Master Checklist\'!G{C[0]}:G{C[1]},"<"&TODAY(),\'Master Checklist\'!H{C[0]}:H{C[1]},"<>Completed")',
             FMT_INT, "9D646B")
    over_budget_formula = (
        "=SUMPRODUCT((Lists!$" + get_column_letter(H_ACTUAL) + "$2:$" + get_column_letter(H_ACTUAL) + "$21>"
        + "Lists!$" + get_column_letter(H_PLANNED) + "$2:$" + get_column_letter(H_PLANNED) + "$21)*1)"
    )
    kpi_card(ws, att_row, c3, "Over-Budget Categories", over_budget_formula, FMT_INT, "9D646B")

    charts_start = att_row + 5

    # ================= CHART-DATA HELPER AREA =================
    # Lives on the (fully hidden) Lists sheet rather than on Dashboard itself.
    # LibreOffice's chart engine reliably preserves a chart's category/value
    # references when they point at a DIFFERENT sheet than the chart's own host
    # sheet; helper data placed directly on Dashboard was found (empirically,
    # verified against the saved/recalculated XML) to have its <c:cat>/<c:val>
    # elements dropped by LibreOffice's resave for some chart shapes. Routing
    # every helper value through Lists sidesteps that risk entirely and also
    # means Dashboard's print area no longer needs to carve out hidden columns.
    lst = wb["Lists"]
    BC_COL = get_column_letter(LIST_COL["BudgetCategory"])       # Lists col for budget categories
    RSVP_COL = get_column_letter(LIST_COL["RSVPStatus"])
    MEAL_COL = get_column_letter(LIST_COL["MealPreference"])
    PERIOD_COL = get_column_letter(LIST_COL["ChecklistPeriod"])

    lst.cell(row=1, column=H_PLANNED, value="Planned")
    lst.cell(row=1, column=H_ACTUAL, value="Actual")
    for i in range(20):
        r = 2 + i
        lst.cell(row=r, column=H_PLANNED, value=f"=SUMIF(Budget!$B${B[0]}:$B${B[1]},Lists!{BC_COL}{r},Budget!$F${B[0]}:$F${B[1]})")
        lst.cell(row=r, column=H_ACTUAL, value=f"=SUMIF(Budget!$B${B[0]}:$B${B[1]},Lists!{BC_COL}{r},Budget!$G${B[0]}:$G${B[1]})")

    lst.cell(row=2, column=H_PAIDREM_LABEL, value="Paid")
    lst.cell(row=2, column=H_PAIDREM_VALUE, value=f"=SUM(Budget!J{B[0]}:J{B[1]})")
    lst.cell(row=3, column=H_PAIDREM_LABEL, value="Remaining")
    lst.cell(row=3, column=H_PAIDREM_VALUE, value=f"=SUM(Budget!K{B[0]}:K{B[1]})")

    for k in range(1, 6):
        r = 1 + k
        lst.cell(row=r, column=H_TOP5_VAL, value=f"=IFERROR(LARGE(Budget!$G${B[0]}:$G${B[1]},{k}),\"\")")
        lst.cell(row=r, column=H_TOP5_LABEL,
                 value=f"=IFERROR(INDEX(Budget!$D${B[0]}:$D${B[1]},MATCH({get_column_letter(H_TOP5_VAL)}{r},Budget!$G${B[0]}:$G${B[1]},0)),\"\")")

    for i in range(4):
        r = 2 + i
        lst.cell(row=r, column=H_RSVP, value=f"=COUNTIF('Guest List'!$I${G[0]}:$I${G[1]},Lists!{RSVP_COL}{r})")
    for i in range(6):
        r = 2 + i
        lst.cell(row=r, column=H_MEAL, value=f"=COUNTIF('Guest List'!$L${G[0]}:$L${G[1]},Lists!{MEAL_COL}{r})")

    lst.cell(row=1, column=H_CHK_COMPLETED, value="Completed")
    lst.cell(row=1, column=H_CHK_REMAINING, value="Remaining")
    for i in range(9):
        r = 2 + i
        lst.cell(row=r, column=H_CHK_COMPLETED,
                 value=f"=COUNTIFS('Master Checklist'!$B${C[0]}:$B${C[1]},Lists!{PERIOD_COL}{r},'Master Checklist'!$H${C[0]}:$H${C[1]},\"Completed\")")
        lst.cell(row=r, column=H_CHK_REMAINING,
                 value=f"=COUNTIFS('Master Checklist'!$B${C[0]}:$B${C[1]},Lists!{PERIOD_COL}{r})-{get_column_letter(H_CHK_COMPLETED)}{r}")
        # Short, chart-only label mirroring the same period order (source Lists!N values are unchanged).
        lst.cell(row=r, column=H_SHORT_PERIOD, value=SHORT_PERIOD_LABELS[i])

    # ================= CHARTS =================
    catref = Reference(wb["Lists"], min_col=LIST_COL["BudgetCategory"], min_row=2, max_row=21)
    left_anchor = get_column_letter(CHART_LEFT_COL)
    right_anchor = get_column_letter(CHART_RIGHT_COL)

    c1c = BarChart(); c1c.type = "col"
    c1c.add_data(Reference(lst, min_col=H_PLANNED, max_col=H_ACTUAL, min_row=1, max_row=21), titles_from_data=True)
    c1c.set_categories(catref)
    style_chart(c1c, "Budget vs. Actual by Category")
    c1c.y_axis.numFmt = FMT_CURRENCY0
    ws.add_chart(c1c, f"{left_anchor}{charts_start}")

    c2c = DoughnutChart()
    c2c.add_data(Reference(lst, min_col=H_ACTUAL, min_row=2, max_row=21), titles_from_data=False)
    c2c.set_categories(catref)
    style_chart(c2c, "Expenses by Category")
    colorize_doughnut(c2c, 20)
    c2c.dataLabels = DataLabelList(); c2c.dataLabels.showPercent = True
    ws.add_chart(c2c, f"{right_anchor}{charts_start}")

    row2 = charts_start + CHART_ROW_STEP
    c3c = BarChart(); c3c.type = "bar"
    c3c.add_data(Reference(lst, min_col=H_PAIDREM_VALUE, min_row=2, max_row=3), titles_from_data=False)
    c3c.set_categories(Reference(lst, min_col=H_PAIDREM_LABEL, min_row=2, max_row=3))
    style_chart(c3c, "Paid vs. Remaining")
    c3c.legend = None
    ws.add_chart(c3c, f"{left_anchor}{row2}")

    c4c = BarChart(); c4c.type = "bar"
    c4c.add_data(Reference(lst, min_col=H_TOP5_VAL, min_row=2, max_row=6), titles_from_data=False)
    c4c.set_categories(Reference(lst, min_col=H_TOP5_LABEL, min_row=2, max_row=6))
    style_chart(c4c, "Top 5 Wedding Expenses")
    c4c.legend = None
    ws.add_chart(c4c, f"{right_anchor}{row2}")

    row3 = charts_start + 2 * CHART_ROW_STEP
    c5c = DoughnutChart()
    c5c.add_data(Reference(lst, min_col=H_RSVP, min_row=2, max_row=5), titles_from_data=False)
    c5c.set_categories(Reference(wb["Lists"], min_col=LIST_COL["RSVPStatus"], min_row=2, max_row=5))
    style_chart(c5c, "RSVP Status")
    colorize_doughnut(c5c, 4)
    c5c.dataLabels = DataLabelList(); c5c.dataLabels.showVal = True
    ws.add_chart(c5c, f"{left_anchor}{row3}")

    c6c = BarChart(); c6c.type = "col"
    c6c.add_data(Reference(lst, min_col=H_MEAL, min_row=2, max_row=7), titles_from_data=False)
    c6c.set_categories(Reference(wb["Lists"], min_col=LIST_COL["MealPreference"], min_row=2, max_row=7))
    style_chart(c6c, "Guest Meal Preferences")
    c6c.legend = None
    ws.add_chart(c6c, f"{right_anchor}{row3}")

    row4 = charts_start + 3 * CHART_ROW_STEP
    c7c = BarChart(); c7c.type = "col"; c7c.grouping = "stacked"; c7c.overlap = 100
    c7c.add_data(Reference(lst, min_col=H_CHK_COMPLETED, min_row=1, max_row=10), titles_from_data=True)
    c7c.add_data(Reference(lst, min_col=H_CHK_REMAINING, min_row=1, max_row=10), titles_from_data=True)
    c7c.set_categories(Reference(lst, min_col=H_SHORT_PERIOD, min_row=2, max_row=10))
    style_chart(c7c, "Checklist Completion by Period")
    ws.add_chart(c7c, f"{left_anchor}{row4}")

    c8c = BarChart(); c8c.type = "col"
    c8c.add_data(Reference(wb["Honeymoon Budget"], min_col=3, min_row=H[0]-1, max_row=H[1]), titles_from_data=True)
    c8c.add_data(Reference(wb["Honeymoon Budget"], min_col=4, min_row=H[0]-1, max_row=H[1]), titles_from_data=True)
    c8c.set_categories(Reference(wb["Honeymoon Budget"], min_col=2, min_row=H[0], max_row=H[1]))
    style_chart(c8c, "Honeymoon: Planned vs. Actual")
    ws.add_chart(c8c, f"{right_anchor}{row4}")

    last_row = row4 + CHART_ROWS + 2   # small bottom margin

    for c in range(1, LAST_CONTENT_COL + 1):
        ws.column_dimensions[get_column_letter(c)].width = COL_WIDTH

    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0   # tall, chart-rich sheet: fit width, paginate naturally in height
    ws.page_margins.left = 0.3
    ws.page_margins.right = 0.3
    ws.page_margins.top = 0.4
    ws.page_margins.bottom = 0.4
    ws.print_area = f"A1:{get_column_letter(PRINT_RIGHT_COL)}{last_row}"

    ws.sheet_view.zoomScale = 85
    return ws
