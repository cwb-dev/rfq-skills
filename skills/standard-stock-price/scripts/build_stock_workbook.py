import argparse
import json
import re
from collections import defaultdict
from pathlib import Path
from urllib.parse import quote

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

FIELDS = ['lookup', 'line', 'display', 'platform', 'code', 'standard', 'spec', 'material_grade', 'finish', 'price', 'weight', 'moq', 'stock', 'pack', 'match', 'drawing', 'remarks', 'url', 'group']
GROUPS = ['原标准/等效', '等级替代需确认', '表色待确认', 'MOQ待确认', '不匹配/无货']
THIN = Side(style='thin', color='B7C3D0')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill('solid', fgColor='1F4E78')
SUBHEADER_FILL = PatternFill('solid', fgColor='D9EAF7')
INPUT_FILL = PatternFill('solid', fgColor='FFF2CC')
ALT_FILL = PatternFill('solid', fgColor='FCE4D6')
GREEN_FILL = PatternFill('solid', fgColor='E2F0D9')
REF_THIN = Side(style='thin', color='000000')
REF_BORDER = Border(left=REF_THIN, right=REF_THIN, top=REF_THIN, bottom=REF_THIN)
REF_YELLOW = PatternFill('solid', fgColor='FFFF00')
REF_LIGHT = PatternFill('solid', fgColor=Color(theme=3, tint=0.8))


def source_url(option):
    if option.get('url'):
        return option['url']
    code = str(option['code'])
    if option['platform'] == '工品一号':
        return 'https://www.gpyh.com/quickbuy/goodsSearchPage?isGoodsFilterFlg=no&searchKey=' + quote(code) + '&searchTab=jgj&isDirectory=false&provinceId=861&cityId=3524'
    return 'https://www.taodingding.com/searchList?search=' + quote(code)


def display_label(option, item):
    reasons = []
    match = str(option.get('match', ''))
    drawing = str(option.get('drawing', ''))
    remarks = str(option.get('remarks', ''))
    grade = re.search(r'(\d+(?:\.\d+)?)(?:级)?\s*（原要求\s*(\d+(?:\.\d+)?)', str(option.get('material_grade', '')))
    if grade:
        reasons.append(f'等级{grade.group(2)}→{grade.group(1)}')
    elif '等级替代' in match:
        reasons.append('等级替代')
    if option.get('standard') in ('特30', 'GB30') and item.get('standard') == 'DIN 933':
        reasons.append('特30≠DIN933')
    if option.get('standard') in ('GB6170', 'GB6170A') and item.get('standard') == 'DIN 934':
        reasons.append('GB6170尺寸差异')
    if '牙型不符' in match or '客户要半牙' in drawing:
        reasons.append('全牙≠半牙')
    if '图纸尺寸差异' in match or '内六角不同' in drawing or '内六角约17不符' in drawing:
        reasons.append('内六角不符')
    if '厂标' in match:
        reasons.append('厂标尺寸差异')
    finish = str(option.get('finish', ''))
    if '白锌（' in finish:
        reasons.append('白锌≠蓝白锌')
    elif option.get('group') == '表色待确认':
        reasons.append(f'表色待确认（{finish}）')
    if '无网站图纸' in drawing or '无图纸' in drawing:
        reasons.append('无网站图纸')
    elif '需核对图纸' in drawing or '需与DIN933图纸核对' in drawing:
        reasons.append('图纸待核')
    elif '图中未标注尺寸' in drawing:
        reasons.append('图纸未注尺寸')
    stock = option.get('stock')
    if stock is not None and stock < item['quantity']:
        reasons.append('库存不足')
    if '卖完下架' in remarks:
        reasons.append('卖完下架')
    reasons = list(dict.fromkeys(reasons))
    status = '完全符合' if not reasons else '不完全符合：' + '、'.join(reasons)
    if option.get('group') == '表色待确认' and not reasons:
        status = '待确认：表色'
    moq = option.get('moq')
    if moq and moq > 1:
        status += f'｜MOQ {moq}'
    return f"{option['platform']}｜{option['code']}｜{status}"


def validate(data):
    items = data['items']
    options = data['options']
    lines = [item['line'] for item in items]
    if len(lines) != len(set(lines)):
        raise ValueError('Duplicate item line')
    option_lines = {o['line'] for o in options}
    if not set(lines).issubset(option_lines):
        raise ValueError('Every item needs at least one option record')
    seen = set()
    by_line = defaultdict(list)
    for option in options:
        key = (option['line'], option['platform'], str(option['code']))
        if key in seen:
            raise ValueError(f'Duplicate option: {key}')
        seen.add(key)
        by_line[option['line']].append(option)
        if not option.get('remarks') or not option.get('match'):
            raise ValueError(f'Missing remarks/match: {key}')
    for item in items:
        if item.get('default'):
            platform, code = item['default'].split('｜', 1)
            if (item['line'], platform, code) not in seen:
                raise ValueError(f"Default not found: {item['line']} {item['default']}")
    return by_line


def style_table(ws, header_row, last_row, last_col, body_size=11, header_size=10):
    for cell in ws[header_row]:
        if cell.column <= last_col:
            cell.fill = HEADER_FILL
            cell.font = Font(name='微软雅黑', size=header_size, bold=True, color='FFFFFF')
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = BORDER
    for row in ws.iter_rows(min_row=header_row + 1, max_row=last_row, max_col=last_col):
        for cell in row:
            cell.font = Font(name='微软雅黑', size=body_size)
            cell.alignment = Alignment(vertical='center', wrap_text=True)
            cell.border = BORDER


def setup_page(ws, title_row, last_col, zoom=130):
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = 'A3'
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f'1:{title_row}'
    ws.sheet_view.zoomScale = zoom


def style_quote_reference(ws, last_row):
    for row in ws.iter_rows(min_row=1, max_row=last_row, max_col=24):
        for cell in row:
            cell.font = Font(name='微软雅黑', size=11, color='000000')
            cell.fill = PatternFill(fill_type=None)
            cell.border = Border()
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for cell in ws[1]:
        cell.border = REF_BORDER
    ws['A1'].font = Font(name='微软雅黑', size=11, bold=True, color='000000')
    ws['H1'].font = Font(name='微软雅黑', size=10, bold=True, color='000000')
    ws['I1'].font = Font(name='微软雅黑', size=10, color='000000')
    ws['I1'].number_format = 'yyyy/m/d;@'
    ws['K1'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    for cell in ws[2]:
        cell.font = Font(name='微软雅黑', size=10, bold=True, color='000000')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for c in range(1, 11):
        ws.cell(2, c).border = REF_BORDER
    for c in range(7, 10):
        ws.cell(2, c).fill = REF_YELLOW
    for c in range(11, 25):
        ws.cell(2, c).font = Font(name='微软雅黑', size=11, bold=c in (16, 18, 19, 20), color='000000')
        ws.cell(2, c).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    for r in range(3, last_row + 1):
        for c in range(1, 11):
            ws.cell(r, c).border = REF_BORDER
        for c in (3, 12):
            ws.cell(r, c).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        for c in (11, 12):
            ws.cell(r, c).font = Font(name='宋体', size=11, color='000000')
        for c in (16, 18, 20):
            ws.cell(r, c).fill = REF_LIGHT
    ws.row_dimensions[1].height = 16.2
    ws.row_dimensions[2].height = 16.2


def main():
    parser = argparse.ArgumentParser(description='Build two-sheet standard-stock price workbook with linked dropdowns.')
    parser.add_argument('data')
    parser.add_argument('output')
    args = parser.parse_args()
    data = json.loads(Path(args.data).read_text(encoding='utf-8'))
    by_line = validate(data)
    items = data['items']
    options = data['options']
    option_count = len(options)
    last_option_row = option_count + 1
    margin = data.get('margin', 0.5)
    vat = data.get('vat_divisor', 1.13)

    wb = Workbook()
    quote = wb.active
    quote.title = '报价汇总'
    search = wb.create_sheet('平台查询')
    helper = wb.create_sheet('选项数据')

    item_by_line = {item['line']: item for item in items}
    displays = {}
    for c, field in enumerate(FIELDS, 1):
        helper.cell(1, c, field)
    for r, option in enumerate(options, 2):
        display = display_label(option, item_by_line[option['line']])
        displays[(option['line'], option['platform'], str(option['code']))] = display
        row = {
            'lookup': f"{option['line']}|{display}",
            'display': display,
            'url': source_url(option),
        }
        for c, field in enumerate(FIELDS, 1):
            helper.cell(r, c, row.get(field, option.get(field)))

    list_start_col = 21
    for item in items:
        line = item['line']
        col = list_start_col + line - 1
        letter = get_column_letter(col)
        helper.cell(1, col, f'DV_{line}')
        values = [displays[(o['line'], o['platform'], str(o['code']))] for o in by_line[line]]
        for r, value in enumerate(values, 2):
            helper.cell(r, col, value)
        ref = f"'选项数据'!${letter}$2:${letter}${len(values)+1}"
        wb.defined_names.add(DefinedName(f'DV_{line}', attr_text=ref))
    helper.sheet_state = 'hidden'

    search.merge_cells('A1:T1')
    search['A1'] = f"{data['inquiry']} — 平台查询与条码选择（{data.get('query_date', '')}）"
    search['A1'].font = Font(name='微软雅黑', size=15, bold=True, color='1F4E78')
    search['A1'].alignment = Alignment(vertical='center')
    search_headers = ['序号','客户标准','客户规格','数量','选择条码/料号','平台','条码/料号','匹配标准/规格','材质/等级','表面处理','含税单价','单重kg/千支','MOQ（支）','库存（支）','包装','匹配情况','图纸核对','备注','报价分类','来源URL']
    for c, value in enumerate(search_headers, 1):
        search.cell(2, c, value)
    helper_map = {'F':'D','G':'E','I':'H','J':'I','K':'J','L':'K','M':'L','N':'M','O':'N','P':'O','Q':'P','R':'Q','S':'S','T':'R'}
    for item in items:
        r = item['line'] + 2
        search.cell(r, 1, item['line'])
        search.cell(r, 2, item['standard'])
        search.cell(r, 3, item['description'])
        search.cell(r, 4, item['quantity'])
        default = item.get('default', '')
        if default:
            platform, code = default.split('｜', 1)
            default = displays[(item['line'], platform, code)]
        search.cell(r, 5, default)
        search.cell(r, 5).fill = INPUT_FILL
        search.cell(r, 5).comment = Comment('从下拉列表选择；价格、单重、MOQ、图纸及备注会联动。', 'Kimi')
        dv = DataValidation(type='list', formula1=f"=DV_{item['line']}", allow_blank=True)
        dv.error = '请选择本行已保存的条码/料号'
        dv.errorTitle = '无效选项'
        dv.prompt = '仅显示本行已核查的工品一号/淘钉钉选项'
        dv.promptTitle = '选择现货'
        search.add_data_validation(dv)
        dv.add(search.cell(r, 5))
        match = f"$A{r}&\"|\"&$E{r}"
        for col, helper_col in helper_map.items():
            formula = f'=IFERROR(INDEX(选项数据!${helper_col}$2:${helper_col}${last_option_row},MATCH({match},选项数据!$A$2:$A${last_option_row},0)),"")'
            search[f'{col}{r}'] = formula
        search[f'H{r}'] = f'=IFERROR(INDEX(选项数据!$F$2:$F${last_option_row},MATCH({match},选项数据!$A$2:$A${last_option_row},0))&" / "&INDEX(选项数据!$G$2:$G${last_option_row},MATCH({match},选项数据!$A$2:$A${last_option_row},0)),"")'
        search.row_dimensions[r].height = 42
    search.auto_filter.ref = f'A2:T{len(items)+2}'
    style_table(search, 2, len(items)+2, 20, body_size=12, header_size=11)
    search.row_dimensions[2].height = 34
    widths = [7,13,30,11,55,12,17,25,25,18,13,16,12,13,23,27,45,45,19,45]
    for c, width in enumerate(widths, 1):
        search.column_dimensions[get_column_letter(c)].width = width
    for r in range(3, len(items)+3):
        for c in (1,4,11,12,13,14):
            search.cell(r, c).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        for c in (11,12): search.cell(r, c).number_format = '#,##0.00'
        for c in (4,13,14): search.cell(r, c).number_format = '#,##0'
    setup_page(search, 2, 20, zoom=90)

    quote.merge_cells('A1:G1')
    quote['A1'] = f"询价计算 — {data['inquiry']}"
    quote['H1'] = '日期'
    quote['I1'] = '=TODAY()'
    quote['K1'] = f"双平台现货｜margin {margin:.0%}（未税成本加成）"
    quote_headers = ['序号','标准/图号','描述','数量','材质','表面处理','含税价格','单重','交期','模具费','到手含税价格（元/千支）','备注','采购总价','总重','不含税成本','margin','CNY','Price to Client CNY/KPCS','MOLD TO CLIENT CNY','LEAD TIME TO CLIENT','MOQ（支）','实际采购量（支）','报价分类','客户总价']
    for c, value in enumerate(quote_headers, 1):
        quote.cell(2, c, value)
    for item in items:
        r = item['line'] + 2
        sr = r
        quote.cell(r, 1, '=ROW()-2')
        quote.cell(r, 2, item['standard'])
        quote.cell(r, 3, item['description'])
        quote.cell(r, 4, item['quantity'])
        quote.cell(r, 5, item['material'])
        quote.cell(r, 6, item['finish'])
        quote[f'G{r}'] = f'=平台查询!K{sr}'
        quote[f'H{r}'] = f'=平台查询!L{sr}'
        quote[f'I{r}'] = f'=IF(K{r}="","","现货/以选项为准")'
        quote[f'K{r}'] = f'=平台查询!K{sr}'
        note = str(item.get('note', '')).replace('"', '""')
        quote[f'L{r}'] = f'=IF(K{r}="","{note}",("{note}｜"&平台查询!Q{sr}&"｜"&平台查询!R{sr}))'
        quote[f'U{r}'] = f'=平台查询!M{sr}'
        quote[f'V{r}'] = f'=IF(K{r}="","",IF(U{r}<=1,D{r},CEILING(D{r},U{r})))'
        quote[f'M{r}'] = f'=IF(K{r}="","",K{r}*V{r}/1000)'
        quote[f'N{r}'] = f'=IF(H{r}="","",H{r}*V{r}/1000)'
        quote[f'O{r}'] = f'=IF(K{r}="","",K{r}/{vat})'
        quote[f'P{r}'] = margin
        quote[f'P{r}'].fill = INPUT_FILL
        quote[f'Q{r}'] = f'=IF(O{r}="","",O{r}*(1+P{r}))'
        quote[f'R{r}'] = f'=IF(Q{r}="","",ROUNDUP(Q{r},0))'
        quote[f'T{r}'] = f'=IF(K{r}="","","现货；下单前复核")'
        quote[f'W{r}'] = f'=平台查询!S{sr}'
        quote[f'X{r}'] = f'=IF(R{r}="","",R{r}*D{r}/1000)'
        quote.row_dimensions[r].height = 42
    quote.auto_filter.ref = f'A2:X{len(items)+2}'
    style_quote_reference(quote, len(items)+2)
    quote_widths = [7,15,32,12,14,12,14,11,12,12,20,58,15,12,15,10,14,24,20,22,12,17,20,14]
    for c, width in enumerate(quote_widths, 1):
        quote.column_dimensions[get_column_letter(c)].width = width
    for r in range(3, len(items)+3):
        for c in (1,4,7,8,9,10,13,14,15,16,17,18,19,20,21,22,23,24):
            quote.cell(r, c).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        for c in (7,8,11,13,14,15,17,24): quote.cell(r, c).number_format = '#,##0.00'
        quote.cell(r, 16).number_format = '0%'
        quote.cell(r, 18).number_format = '#,##0'
        for c in (4,21,22): quote.cell(r, c).number_format = '#,##0'
    setup_page(quote, 2, 24)

    summary_row = len(items) + 4
    first_data = 3
    last_data = len(items) + 2
    quote.merge_cells(start_row=summary_row, start_column=1, end_row=summary_row, end_column=24)
    quote.cell(summary_row, 1, '汇总（客户单价按 R 列向上取整；采购成本按MOQ后的实际采购量）')
    quote.cell(summary_row, 1).fill = INPUT_FILL
    quote.cell(summary_row, 1).font = Font(name='微软雅黑', size=11, bold=True, color='000000')
    quote.cell(summary_row, 1).border = Border(top=REF_THIN, bottom=REF_THIN)
    quote.row_dimensions[summary_row].height = 30
    summary = [
        (1, '采购成本（含税）', f'=SUM(M{first_data}:M{last_data})', '理论总价（Q列不取整）', f'=SUMPRODUCT(Q{first_data}:Q{last_data},D{first_data}:D{last_data})/1000'),
        (2, '客户总价（R列取整）', f'=SUM(X{first_data}:X{last_data})', '理论利润（Q列不取整）', None),
        (3, '预计利润（R列取整）', None, '未税成本加成', margin),
        (4, '销售毛利率（参考）', None, '成本利润率（参考）', None),
    ]
    for offset, left, left_formula, right, right_formula in summary:
        r = summary_row + offset
        quote.cell(r, 1, left)
        quote.cell(r, 2, left_formula)
        quote.cell(r, 4, right)
        quote.cell(r, 5, right_formula)
        for c in (1, 4):
            quote.cell(r, c).font = Font(name='微软雅黑', size=10, bold=True, color='000000')
        for c in (2, 5):
            quote.cell(r, c).font = Font(name='Arial', size=10, bold=offset in (2, 3), color='000000')
            quote.cell(r, c).number_format = '#,##0.00'
    quote.cell(summary_row+2, 5, f'=E{summary_row+1}-B{summary_row+1}')
    quote.cell(summary_row+3, 2, f'=B{summary_row+2}-B{summary_row+1}')
    quote.cell(summary_row+4, 2, f'=IFERROR(B{summary_row+3}/B{summary_row+2},0)')
    quote.cell(summary_row+4, 5, f'=IFERROR(B{summary_row+3}/B{summary_row+1},0)')
    quote.cell(summary_row+3, 5).number_format = '0.00%'
    quote.cell(summary_row+4, 2).number_format = '0.00%'
    quote.cell(summary_row+4, 5).number_format = '0.00%'
    for c, value in enumerate(['报价分类','行数','采购成本','客户总价','预计利润'], 7):
        cell = quote.cell(summary_row+1, c, value)
        cell.font = Font(name='微软雅黑', size=10, bold=True, color='000000')
        cell.border = Border(bottom=REF_THIN)
    for idx, group in enumerate(GROUPS, summary_row+2):
        quote.cell(idx, 7, group)
        quote.cell(idx, 8, f'=COUNTIF(W{first_data}:W{last_data},G{idx})')
        quote.cell(idx, 9, f'=SUMIF(W{first_data}:W{last_data},G{idx},M{first_data}:M{last_data})')
        quote.cell(idx, 10, f'=SUMIF(W{first_data}:W{last_data},G{idx},X{first_data}:X{last_data})')
        quote.cell(idx, 11, f'=J{idx}-I{idx}')
        for c in range(7, 12):
            quote.cell(idx, c).font = Font(name='微软雅黑', size=10, color='000000')
            quote.cell(idx, c).alignment = Alignment(vertical='center', wrap_text=True)
        for c in (9,10,11):
            quote.cell(idx, c).number_format = '#,##0.00'
    notes_row = summary_row + 8
    notes = [
        f'计算口径：P列 margin={margin:.0%}；Q=K÷{vat}×(1+P)；R=ROUNDUP(Q,0)。这里的 margin 是未税成本加成，不是销售毛利率。',
        f'="结果：采购成本 "&TEXT(B{summary_row+1},"#,##0.00")&" 元；按R列客户单价的销售总价 "&TEXT(B{summary_row+2},"#,##0.00")&" 元；预计利润 "&TEXT(B{summary_row+3},"#,##0.00")&" 元。"',
        '预计利润=客户总价−含税采购成本；未扣运费、出口费用、汇兑及其他费用；未知单重保持空白，不估算。',
        '在“平台查询”E列选择条码/料号；本页单价、单重、MOQ、实际采购量、备注、成本和客户价自动联动。“选项数据”为隐藏辅助页，不要删除。',
        '默认先选图纸和等级符合项；仅在没有符合项时保留替代等级。MOQ不作为降级或排除符合项的理由。等级替代、图纸差异和表色待确认项必须经确认；库存价格下单前复核。',
    ]
    for r, text in enumerate(notes, notes_row):
        quote.merge_cells(start_row=r, start_column=1, end_row=r, end_column=24)
        quote.cell(r, 1, text)
        quote.cell(r, 1).font = Font(name='微软雅黑', size=10, color='000000')
        quote.cell(r, 1).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        quote.row_dimensions[r].height = 30
    quote.print_area = f'A1:X{notes_row+len(notes)-1}'
    quote.conditional_formatting.add(f'W3:W{last_data}', FormulaRule(formula=['W3="原标准/等效"'], fill=GREEN_FILL))
    quote.conditional_formatting.add(f'W3:W{last_data}', FormulaRule(formula=['W3="等级替代需确认"'], fill=ALT_FILL))
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = 'auto'
    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    selected = sum(bool(item.get('default')) for item in items)
    print(f'saved {out}')
    print(f'items={len(items)}, options={option_count}, selected={selected}, sheets={wb.sheetnames}')


if __name__ == '__main__':
    main()
