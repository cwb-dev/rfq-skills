from copy import copy
from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.properties import PageSetupProperties

BASE = Path(__file__).resolve().parent
OUT = BASE / "templates"
OUT.mkdir(parents=True, exist_ok=True)

FONT = "Microsoft YaHei"
INK = "FF222222"
GREY = "FF666666"
LIGHT_GREY = "FF8A8A8A"
BRAND = "FFFF6D23"
TINT = "FFFFF0E9"
BAND = "FFF7F7F7"
GRID = "FFE2E2E2"
LAST = "P"


def f(size=10, bold=False, color=INK, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(color):
    return PatternFill("solid", start_color=color, end_color=color)


def cells(ws, cell_range, font=None, fill_color=None, border=None, align=None):
    for row in ws[cell_range]:
        for cell in row:
            if font:
                cell.font = copy(font)
            if fill_color:
                cell.fill = fill(fill_color)
            if border:
                cell.border = copy(border)
            if align:
                cell.alignment = copy(align)


def merge(ws, cell_range, value="", font=None, fill_color=None, border=None, align=None):
    ws.merge_cells(cell_range)
    first = cell_range.split(":")[0]
    ws[first] = value
    cells(ws, cell_range, font=font, fill_color=fill_color, border=border, align=align)


wb = Workbook()
ws = wb.active
ws.title = "采购合同"
ws.sheet_view.showGridLines = False
widths = {
    "A": 7, "B": 7, "C": 7, "D": 10, "E": 10, "F": 7, "G": 2, "H": 7,
    "I": 9, "J": 7, "K": 10, "L": 8, "M": 6, "N": 6, "O": 8, "P": 8,
}
for col, width in widths.items():
    ws.column_dimensions[col].width = width

ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.orientation = "portrait"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
ws.page_margins.left = 0.25
ws.page_margins.right = 0.25
ws.page_margins.top = 0.3
ws.page_margins.bottom = 0.35
ws.page_margins.header = 0.15
ws.page_margins.footer = 0.15
ws.print_options.horizontalCentered = True

for row, height in {1: 24, 2: 15, 3: 15, 4: 5, 5: 4, 6: 7}.items():
    ws.row_dimensions[row].height = height
logo = XLImage(str(BASE / "assets" / "logo-runsun.png"))
logo.width = 124
logo.height = 55
ws.add_image(logo, "A1")
merge(ws, "D1:P1", "上海如榫工业科技有限公司", f(15, True), align=Alignment(horizontal="right", vertical="center"))
merge(ws, "D2:P2", "上海市青浦区香大东路1318号1幢6层", f(8.5, color=GREY), align=Alignment(horizontal="right", vertical="center"))
merge(ws, "D3:P3", "电话：+86 18006193111", f(8.5, color=GREY), align=Alignment(horizontal="right", vertical="center"))
cells(ws, "A5:P5", border=Border(bottom=Side(style="thick", color=BRAND)))

ws.row_dimensions[7].height = 25
ws.row_dimensions[8].height = 18
ws.row_dimensions[9].height = 7
merge(ws, "A7:H7", "外贸产品采购合同", f(17, True), align=Alignment(vertical="center"))
merge(ws, "K7:L7", "合同编号", f(8, color=LIGHT_GREY), align=Alignment(horizontal="right", vertical="center"))
merge(ws, "M7:P7", "", f(10.5, True), align=Alignment(horizontal="right", vertical="center"))
merge(ws, "K8:L8", "地点 / 日期", f(8, color=LIGHT_GREY), align=Alignment(horizontal="right", vertical="center"))
merge(ws, "M8:P8", "", f(10, True), align=Alignment(horizontal="right", vertical="center"))

cells(ws, "A10:G13", fill_color=BAND)
cells(ws, "I10:P13", fill_color=BAND)
for row in range(10, 14):
    ws[f"A{row}"].border = Border(left=Side(style="thick", color=BRAND))
    ws[f"I{row}"].border = Border(left=Side(style="thick", color=BRAND))
merge(ws, "A10:G10", "需方（采购方）", f(8, color=LIGHT_GREY), fill_color=BAND, align=Alignment(vertical="center"))
merge(ws, "I10:P10", "供方（销售方）", f(8, color=LIGHT_GREY), fill_color=BAND, align=Alignment(vertical="center"))
merge(ws, "A11:G13", "", f(10), fill_color=BAND, align=Alignment(vertical="top", wrap_text=True))
merge(ws, "I11:P13", "", f(10), fill_color=BAND, align=Alignment(vertical="top", wrap_text=True))
ws.row_dimensions[10].height = 15
for row in range(11, 14):
    ws.row_dimensions[row].height = 15
ws.row_dimensions[14].height = 7

header_font = f(8, True)
header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
header_border = Border(bottom=Side(style="medium", color=BRAND))
ws.row_dimensions[15].height = 42
merge(ws, "A15:C15", "产品名称", header_font, TINT, header_border, header_align)
merge(ws, "D15:D15", "产品图纸号", header_font, TINT, header_border, header_align)
merge(ws, "E15:E15", "客户图纸号", header_font, TINT, header_border, header_align)
merge(ws, "F15:H15", "产品英文名称", header_font, TINT, header_border, header_align)
merge(ws, "I15:I15", "数量\n千个", header_font, TINT, header_border, header_align)
merge(ws, "J15:J15", "千重\nKG/千个", header_font, TINT, header_border, header_align)
merge(ws, "K15:K15", "总重\nKG", header_font, TINT, header_border, header_align)
merge(ws, "L15:L15", "包装", header_font, TINT, header_border, header_align)
merge(ws, "M15:N15", "价格\n元/千个\n含税运价", header_font, TINT, header_border, header_align)
merge(ws, "O15:P15", "合计\n元", header_font, TINT, header_border, header_align)

row_border = Border(bottom=Side(style="thin", color=GRID))
left_top = Alignment(horizontal="left", vertical="top", wrap_text=True)
center_top = Alignment(horizontal="center", vertical="top", wrap_text=True)
right_top = Alignment(horizontal="right", vertical="top", wrap_text=True)
ws.row_dimensions[16].height = 64
merge(ws, "A16:C16", "", f(9), border=row_border, align=left_top)
merge(ws, "D16:D16", "", f(8.5), border=row_border, align=center_top)
merge(ws, "E16:E16", "", f(8.5), border=row_border, align=center_top)
merge(ws, "F16:H16", "", f(8.5), border=row_border, align=left_top)
merge(ws, "I16:I16", "", f(9), border=row_border, align=right_top)
merge(ws, "J16:J16", "", f(9), border=row_border, align=right_top)
merge(ws, "K16:K16", "", f(9), border=row_border, align=right_top)
merge(ws, "L16:L16", "", f(8.5), border=row_border, align=center_top)
merge(ws, "M16:N16", "", f(9), border=row_border, align=right_top)
merge(ws, "O16:P16", "", f(9), border=row_border, align=right_top)
ws["I16"].number_format = "#,##0.000"
for col in ("J", "K", "M", "O"):
    ws[f"{col}16"].number_format = "#,##0.00"

ws.row_dimensions[17].height = 21
cells(ws, "A17:P17", font=f(10, True), border=Border(top=Side(style="medium", color=BRAND)))
merge(ws, "A17:H17", "合计：", f(10, True), align=Alignment(horizontal="right", vertical="center"))
merge(ws, "I17:I17", "", f(10, True), align=Alignment(horizontal="right", vertical="center"))
merge(ws, "J17:J17", "", f(10, True))
merge(ws, "K17:K17", "", f(10, True), align=Alignment(horizontal="right", vertical="center"))
merge(ws, "L17:N17", "", f(10, True))
merge(ws, "O17:P17", "", f(10, True), align=Alignment(horizontal="right", vertical="center"))
ws["I17"].number_format = "#,##0.000"
ws["K17"].number_format = "#,##0.00"
ws["O17"].number_format = "#,##0.00"
ws.row_dimensions[18].height = 18
merge(ws, "A18:P18", "", f(9.5, italic=True, color=GREY), align=Alignment(vertical="center"))
ws.row_dimensions[19].height = 7

merge(ws, "A20:P20", "包装及标签要求", f(8, color=LIGHT_GREY), align=Alignment(vertical="center"))
ws.row_dimensions[20].height = 15
merge(ws, "A21:P21", "", f(8.5), fill_color=BAND, align=Alignment(vertical="center", wrap_text=True))
ws.row_dimensions[21].height = 26
def row_range(columns, row):
    first, last = columns.split(":")
    return f"{first}{row}:{last}{row}"


label_ranges = (("A:B", "C:D", "E:G"), ("I:J", "K:L", "M:P"))
for row in range(22, 25):
    for label_cn, label_en, value in label_ranges:
        merge(ws, row_range(label_cn, row), "", f(8), border=row_border, align=Alignment(vertical="center", wrap_text=True))
        merge(ws, row_range(label_en, row), "", f(8, color=GREY), border=row_border, align=Alignment(vertical="center", wrap_text=True))
        merge(ws, row_range(value, row), "", f(8.5), border=row_border, align=Alignment(vertical="center", wrap_text=True))
    ws.row_dimensions[row].height = 30
ws.row_dimensions[25].height = 7
merge(ws, "A26:P26", "合同条款", f(8, color=LIGHT_GREY), align=Alignment(vertical="center"))
ws.row_dimensions[26].height = 15
for row in range(27, 36):
    merge(ws, f"A{row}:P{row}", "", f(9), align=Alignment(vertical="center", wrap_text=True))
    ws.row_dimensions[row].height = 18
ws.row_dimensions[36].height = 7

merge(ws, "A37:G37", "需方（采购方）", f(9, True), fill_color=TINT, align=Alignment(vertical="center"))
merge(ws, "I37:P37", "供方（销售方）", f(9, True), fill_color=TINT, align=Alignment(vertical="center"))
ws.row_dimensions[37].height = 18
sign_rows = {38: 18, 39: 28, 40: 18, 41: 18, 42: 18, 43: 18, 44: 18}
for row, height in sign_rows.items():
    merge(ws, f"A{row}:G{row}", "", f(8.5), border=row_border, align=Alignment(vertical="center", wrap_text=True))
    merge(ws, f"I{row}:P{row}", "", f(8.5), border=row_border, align=Alignment(vertical="center", wrap_text=True))
    ws.row_dimensions[row].height = height
ws.row_dimensions[45].height = 28
merge(ws, "A45:G45", "签字 / 盖章：", f(9), align=Alignment(vertical="bottom"))
merge(ws, "I45:P45", "签字 / 盖章：", f(9), align=Alignment(vertical="bottom"))
ws.row_dimensions[46].height = 7
ws.row_dimensions[47].height = 14
cells(ws, "A47:P47", border=Border(top=Side(style="thin", color="FFDDDDDD")))
merge(ws, "A47:P47", "上海如榫工业科技有限公司   ·   上海市青浦区香大东路1318号1幢6层", f(7.5, color="FF999999"), align=Alignment(horizontal="center", vertical="center"))

ws.freeze_panes = "A15"
ws.print_area = "A1:P47"
wb.save(OUT / "PC-RUNSUN.xlsx")
print(OUT / "PC-RUNSUN.xlsx")
