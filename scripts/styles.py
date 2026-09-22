"""Shared design system: palette, fonts, borders, and small styling helpers
used across every sheet of the Premium Wedding Planner workbook."""

from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.formatting.rule import FormulaRule, CellIsRule, DataBarRule

# ---------------------------------------------------------------- PALETTE --
IVORY = "FBF7F2"
BLUSH = "E8C7C3"
DUSTY_ROSE = "C98F91"
SAGE = "A8B5A0"
CHARCOAL = "333333"
GOLD = "B89B5E"

CREAM = "FFFDF9"
LIGHT_BLUSH = "F5E4E1"
LIGHT_SAGE = "E3E9DF"
TAUPE = "B7A69C"
BEIGE = "EDE3DA"
DEEP_ROSE = "9D646B"
BROWN = "715B52"
MUTED_GOLD = "D2BC8B"
WHITE = "FFFFFF"

# Row banding
ROW_A = WHITE
ROW_B = "FBF5F0"  # very light blush-ivory

FONT_NAME = "Calibri"

# ------------------------------------------------------------------ FONTS --
F_TITLE = Font(name=FONT_NAME, size=22, bold=True, color=CHARCOAL)
F_SUBTITLE = Font(name=FONT_NAME, size=11, italic=True, color=BROWN)
F_SHEET_TITLE = Font(name=FONT_NAME, size=18, bold=True, color=CHARCOAL)
F_SHEET_SUB = Font(name=FONT_NAME, size=10.5, italic=True, color=BROWN)
F_SECTION = Font(name=FONT_NAME, size=12.5, bold=True, color=DEEP_ROSE)
F_HEADER = Font(name=FONT_NAME, size=10.5, bold=True, color=CHARCOAL)
F_HEADER_LIGHT = Font(name=FONT_NAME, size=10.5, bold=True, color=WHITE)
F_BODY = Font(name=FONT_NAME, size=10.5, color=CHARCOAL)
F_BODY_MUTED = Font(name=FONT_NAME, size=9.5, italic=True, color=BROWN)
F_KPI_NUM = Font(name=FONT_NAME, size=20, bold=True, color=DEEP_ROSE)
F_KPI_LABEL = Font(name=FONT_NAME, size=9, bold=True, color=CHARCOAL)
F_NAV = Font(name=FONT_NAME, size=9.5, bold=True, color=BROWN, underline="single")
F_LINK = Font(name=FONT_NAME, size=10.5, color="1155CC", underline="single")

# ---------------------------------------------------------------- FILLS ----
def fill(color):
    return PatternFill("solid", fgColor=color)

FILL_IVORY = fill(IVORY)
FILL_CREAM = fill(CREAM)
FILL_BLUSH = fill(BLUSH)
FILL_LIGHT_BLUSH = fill(LIGHT_BLUSH)
FILL_SAGE = fill(SAGE)
FILL_LIGHT_SAGE = fill(LIGHT_SAGE)
FILL_GOLD = fill(GOLD)
FILL_MUTED_GOLD = fill(MUTED_GOLD)
FILL_BEIGE = fill(BEIGE)
FILL_CHARCOAL = fill(CHARCOAL)
FILL_WHITE = fill(WHITE)
FILL_HEADER = fill(DEEP_ROSE)

# --------------------------------------------------------------- BORDERS ---
THIN_GOLD = Side(style="thin", color=GOLD)
THIN_TAUPE = Side(style="thin", color=TAUPE)
THIN_LIGHT = Side(style="thin", color="E6DDD3")
MED_ROSE = Side(style="medium", color=DUSTY_ROSE)

B_GOLD_BOTTOM = Border(bottom=THIN_GOLD)
B_TABLE_HEADER = Border(bottom=Side(style="medium", color=GOLD))
B_ALL_LIGHT = Border(left=THIN_LIGHT, right=THIN_LIGHT, top=THIN_LIGHT, bottom=THIN_LIGHT)
B_CARD = Border(left=Side(style="thin", color=GOLD), right=Side(style="thin", color=GOLD),
                 top=Side(style="thin", color=GOLD), bottom=Side(style="thin", color=GOLD))

ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_LEFT_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
ALIGN_CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)

# ---------------------------------------------------------------- NUMBER FORMATS
FMT_CURRENCY = '$#,##0.00;[RED]-$#,##0.00'
FMT_CURRENCY0 = '$#,##0;[RED]-$#,##0'
FMT_PCT = '0%'
FMT_PCT1 = '0.0%'
FMT_DATE = 'd mmm yyyy'
FMT_TIME = 'h:mm AM/PM'
FMT_INT = '#,##0'

SHEET_TAB_COLOR = DUSTY_ROSE


def set_sheet_defaults(ws: Worksheet, tab_color=SHEET_TAB_COLOR):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = tab_color
    ws.sheet_view.zoomScale = 100


def sheet_header(ws: Worksheet, title, subtitle, last_col_letter="N", nav_links=None, row_height=34):
    """Draws the standard elegant sheet banner in rows 1-3, plus optional nav row 4."""
    for r in range(1, 4):
        for c in range(1, col_idx(last_col_letter) + 1):
            ws.cell(row=r, column=c).fill = FILL_LIGHT_BLUSH
    ws.row_dimensions[1].height = 8
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[3].height = 20
    t = ws.cell(row=2, column=2, value=title)
    t.font = F_SHEET_TITLE
    t.alignment = ALIGN_LEFT
    s = ws.cell(row=3, column=2, value=subtitle)
    s.font = F_SHEET_SUB
    s.alignment = ALIGN_LEFT
    # gold accent rule under banner
    for c in range(1, col_idx(last_col_letter) + 1):
        ws.cell(row=3, column=c).border = Border(bottom=Side(style="medium", color=GOLD))
    if nav_links:
        col = 2
        for label, target in nav_links:
            cell = ws.cell(row=1, column=col, value=f"‹ {label}")
            cell.font = F_NAV
            cell.hyperlink = f"#'{target}'!A1"
            col += 1
    return 5  # first free row after header


def col_idx(letter):
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(letter)


def style_band(ws, row, first_col, last_col, banded=True, idx=0):
    f = FILL_WHITE if idx % 2 == 0 else fill(ROW_B)
    for c in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = f
        cell.font = F_BODY
        cell.border = B_ALL_LIGHT


def kpi_card(ws, top_row, left_col, label, value_formula, number_format=FMT_CURRENCY0,
             accent=DUSTY_ROSE, height=3, width=3, label_first=False, text_value=False):
    """Draws a KPI card spanning `height` rows x `width` cols starting at (top_row, left_col).
    Uses two separate merge blocks (value, label) so only their own anchor cells are written."""
    r0, c0 = top_row, left_col
    r1, c1 = top_row + height - 1, left_col + width - 1
    label_h = 1
    value_h = height - label_h
    if label_first:
        label_r0, label_r1 = r0, r0
        value_r0, value_r1 = r0 + label_h, r1
    else:
        value_r0, value_r1 = r0, r0 + value_h - 1
        label_r0, label_r1 = value_r1 + 1, r1

    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = FILL_CREAM
            cell.border = B_CARD
    # top accent strip
    for c in range(c0, c1 + 1):
        ws.cell(row=r0, column=c).border = Border(
            top=Side(style="medium", color=accent),
            left=THIN_GOLD if c == c0 else None,
            right=THIN_GOLD if c == c1 else None,
        )

    if value_r1 > value_r0 or c1 > c0:
        ws.merge_cells(start_row=value_r0, start_column=c0, end_row=value_r1, end_column=c1)
    val = ws.cell(row=value_r0, column=c0)
    val.value = value_formula
    val.font = Font(name=FONT_NAME, size=19, bold=True, color=accent)
    if not text_value:
        val.number_format = number_format
    val.alignment = ALIGN_CENTER

    if label_r1 > label_r0 or c1 > c0:
        ws.merge_cells(start_row=label_r0, start_column=c0, end_row=label_r1, end_column=c1)
    lbl = ws.cell(row=label_r0, column=c0)
    lbl.value = label.upper()
    lbl.font = F_KPI_LABEL
    lbl.alignment = ALIGN_CENTER


def section_title(ws, row, col, text, last_col_letter=None):
    cell = ws.cell(row=row, column=col, value=text.upper())
    cell.font = F_SECTION
    if last_col_letter:
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col_idx(last_col_letter))
    return row + 1
