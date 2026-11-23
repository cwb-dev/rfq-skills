import argparse
import json
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET

import pymupdf

DARK = (0.13, 0.13, 0.13)
GREY = (0.48, 0.48, 0.48)
ORANGE = (0.96, 0.51, 0.13)
LIGHT_ORANGE = (1.0, 0.95, 0.91)
RED = (0.75, 0.0, 0.0)
WHITE = (1, 1, 1)


def find_font():
    for p in [r"C:/Windows/Fonts/msyh.ttc", r"C:/Windows/Fonts/Deng.ttf", r"C:/Windows/Fonts/simhei.ttf"]:
        if Path(p).exists():
            return p
    raise FileNotFoundError("No Chinese font found")


def wrap(text, font, size, width):
    result = []
    for paragraph in str(text or "").splitlines() or [""]:
        line = ""
        for char in paragraph:
            if line and font.text_length(line + char, fontsize=size) > width:
                result.append(line)
                line = char
            else:
                line += char
        result.append(line)
    return result


class PdfBuilder:
    def __init__(self, spec):
        self.s = spec
        self.font_file = find_font()
        self.font = pymupdf.Font(fontfile=self.font_file)
        self.doc = pymupdf.open()
        self.page = None

    def new_page(self):
        self.page = self.doc.new_page(width=595.3, height=842.0)
        return self.page

    def text(self, x, y, value, size=9.5, color=DARK):
        self.page.insert_text((x, y), str(value), fontname="CJK", fontfile=self.font_file, fontsize=size, color=color)

    def paragraph(self, value, y, *, x=42, width=511, size=9.5, color=DARK, line_height=14, after=2, align="left"):
        lines = wrap(value, self.font, size, width)
        for line in lines:
            if align == "center":
                tx = x + max(0, (width - self.font.text_length(line, fontsize=size)) / 2)
            elif align == "right":
                tx = x + width - self.font.text_length(line, fontsize=size)
            else:
                tx = x
            self.text(tx, y, line, size, color)
            y += line_height
        return y + after

    def heading(self, value, y):
        return self.paragraph(value, y + 5, size=11, line_height=15, after=4)

    def footer(self):
        for page in self.doc:
            label = f"上海如榫工业科技有限公司  ·  上海市青浦区香大东路1318号1幢6层  ·  {self.s['ref']}"
            page.insert_text((42, 826), label, fontname="CJK", fontfile=self.font_file, fontsize=7.2, color=GREY)

    def letterhead(self):
        logo = Path(__file__).resolve().parents[1] / "assets" / "logo-runsun.png"
        self.page.insert_image(pymupdf.Rect(42, 38, 150, 86), filename=str(logo))
        self.paragraph("上海如榫工业科技有限公司", 52, x=180, width=373, size=12, line_height=15, align="right")
        self.paragraph("中国上海市青浦区香大东路1318号1幢6层", 70, x=180, width=373, size=8.2, color=GREY, line_height=12, align="right")
        self.paragraph("电话：+86 18006193111", 83, x=180, width=373, size=8.2, color=GREY, line_height=12, align="right")
        self.page.draw_line(pymupdf.Point(42, 94), pymupdf.Point(553, 94), color=ORANGE, width=2)

    def cell(self, rect, value, *, size=8.2, bold=False, fill=None, align="left", border=True):
        if fill:
            self.page.draw_rect(rect, color=None, fill=fill)
        if border:
            self.page.draw_rect(rect, color=(0.72, 0.72, 0.72), width=0.45)
        lines = wrap(value, self.font, size, rect.width - 8)
        line_height = size + 2.2
        y = rect.y0 + max(12, (rect.height - len(lines) * line_height) / 2 + size)
        for line in lines:
            if align == "center":
                x = rect.x0 + (rect.width - self.font.text_length(line, fontsize=size)) / 2
            else:
                x = rect.x0 + 4
            self.text(x, y, line, size)
            y += line_height

    def item_table(self, y):
        widths = [65, 60, 85, 85, 50, 85, 80]
        headers = ["产品名称", "图纸号", "规格", "材质", "样品数量\n（个）", "打样费包干价\n（元，未税）", "大货单价\n（元/个，含税）"]
        x0 = 42
        x = x0
        for w, value in zip(widths, headers):
            self.cell(pymupdf.Rect(x, y, x + w, y + 34), value, fill=LIGHT_ORANGE, align="center")
            x += w
        y += 34
        for item in self.s["items"]:
            values = [item["product"], item["drawing_no"], item["spec"], item["material"], str(item["quantity"]), str(item.get("sample_fee") or ""), str(item.get("mass_price") or "")]
            x = x0
            for w, value in zip(widths, values):
                self.cell(pymupdf.Rect(x, y, x + w, y + 43), value, align="center" if w in [67, 62, 82] else "left")
                x += w
            y += 43
        merged_width = sum(widths[:5])
        self.cell(pymupdf.Rect(x0, y, x0 + merged_width, y + 25), "打样费未税合计（大写）：" + str(self.s.get("total_fee_words") or "________________________"))
        self.cell(pymupdf.Rect(x0 + merged_width, y, x0 + merged_width + widths[5], y + 25), str(self.s.get("total_fee") or ""))
        self.cell(pymupdf.Rect(x0 + merged_width + widths[5], y, x0 + sum(widths), y + 25), "")
        return y + 43

    def build(self):
        self.new_page()
        self.letterhead()
        y = self.paragraph("样品试制（打样）协议", 120, width=511, size=16, line_height=20, align="center", after=6)
        parties = [
            (f"甲方（需方）：{self.s['buyer']}", f"协议编号：{self.s['ref']}"),
            (f"乙方（供方）：{self.s['supplier']}", f"签订地点：{self.s['sign_place']}"),
            ("", f"签订日期：{self.s['date']}"),
        ]
        for left, right in parties:
            self.text(42, y, left, 9.5)
            self.text(335, y, right, 9.5)
            y += 17
        y = self.paragraph("甲方委托乙方按甲方提供的图纸试制样品。根据双方沟通，经协商一致，就本次打样事宜签订本协议，双方共同遵守。", y + 3)
        y = self.heading("一、样品内容及费用", y)
        y = self.item_table(y)
        y = self.paragraph("1. 打样费为一次性包干价（未税）；大货单价为含税价，价格口径（勾选）：" + self.s["price_options"] + "。打样费已包含工装、刀具、原材料、加工、检测及包装等费用；选择含运时同时包含运费，乙方不再另行收费。", y, size=9.0, line_height=13)
        y = self.paragraph("2. 模具/工装费用：" + str(self.s.get("tooling_note") or "________________") + "；未填写或未注明另计的，视为已包含在打样费内。", y, size=9.0, line_height=13)
        y = self.paragraph("3. 表中大货单价为样品合格后大货订单的确认执行价；除双方约定的材料价格波动情形外，乙方不得拒绝接受或另行加价。", y, size=9.0, line_height=13)
        if self.s.get("material_price_note"):
            y = self.paragraph("4. " + self.s["material_price_note"], y, size=9.0, line_height=13)
        y = self.heading("二、技术与质量要求", y)
        for i, req in enumerate(self.s["technical_requirements"], 1):
            y = self.paragraph(f"{i}. {req}", y, size=9.0, line_height=13, after=1)
        y = self.heading("三、交样", y)
        for i, key in enumerate(["delivery_quantity", "delivery_date", "delivery_method"], 1):
            y = self.paragraph(f"{i}. {self.s[key]}", y, size=9.0, line_height=13, after=1)
        y = self.paragraph("4. 包装：样品清洁后按图号分别包装，防止磕碰、氧化和混料，外包装注明图纸号、数量及批次。", y, size=9.0, line_height=13)
        if y > 810:
            raise ValueError(f"Agreement page 1 overflow: y={y}")

        self.new_page()
        y = 48
        y = self.heading("四、付款", y)
        y = self.paragraph(self.s["payment_terms"], y)
        y = self.heading("五、合样、验收与退款", y)
        for i, req in enumerate(self.s["acceptance_terms"], 1):
            y = self.paragraph(f"{i}. {req}", y, line_height=14)
        y = self.heading("六、批量生产", y)
        y = self.paragraph(self.s.get("mass_production_terms") or "样品经甲方书面确认合格后，双方另行签订批量采购合同。大货订单执行本协议确认的单价；产品的材料、尺寸、外观及关键工艺须与确认样品一致，任何变更须经甲方书面同意。", y)
        y = self.heading("七、保密", y)
        y = self.paragraph("乙方对甲方提供的图纸、技术资料及本协议内容负有保密义务，未经甲方书面同意，不得向任何第三方披露，不得为甲方以外的任何方生产、销售该产品，也不得直接或间接与该产品的最终用户联系。本条义务不因本协议终止而失效。", y)
        y = self.heading("八、其他", y)
        y = self.paragraph("1. 违约责任按《中华人民共和国民法典》有关规定执行。本协议未尽事宜由双方协商解决；协商不成的，向甲方所在地人民法院起诉。", y)
        drawing_list = "、".join(i["drawing_no"] for i in self.s["items"])
        y = self.paragraph(f"2. 附件“产品图纸 {drawing_list} 及中文技术要求”为本协议不可分割的组成部分。", y)
        y = self.paragraph("3. 本协议一式两份，双方各执一份，经双方盖章后生效。盖章后的扫描件、传真件与原件具有同等效力。", y, after=10)
        left = "甲方（需方）\n单位名称（章）：上海如榫工业科技有限公司\n单位地址：上海市青浦区香大东路1318号1幢6层\n法定代表人：何阳\n委托代理人：\n电  话：18006193111\n开户银行：中国银行上海市徐泾支行\n账  号：453385565724\n税  号：91310118MACTKTHT9T\n日  期：    年  月  日"
        right = f"乙方（供方）\n单位名称（章）：{self.s['supplier']}\n单位地址：\n法定代表人：\n委托代理人：\n电  话：\n开户银行：\n账  号：\n税  号：\n日  期：    年  月  日"
        y0 = y + 5
        sign_bottom = y0 + 145
        self.page.draw_rect(pymupdf.Rect(42, y0 - 12, 553, sign_bottom), color=(0.7, 0.7, 0.7), width=0.5)
        self.page.draw_line(pymupdf.Point(297.5, y0 - 12), pymupdf.Point(297.5, sign_bottom), color=(0.7, 0.7, 0.7), width=0.5)
        self.paragraph(left, y0, x=48, width=240, size=8.5, line_height=12)
        self.paragraph(right, y0, x=304, width=240, size=8.5, line_height=12)

        for item in self.s["items"]:
            self.new_page()
            y = self.paragraph(f"附件：产品图纸 {item['drawing_no']} 及中文技术要求", 50, size=12, line_height=16)
            y = self.paragraph(f"协议编号：{self.s['ref']}", y, size=8.5, color=GREY, line_height=12, after=5)
            for note in item["annex_notes"]:
                y = self.paragraph("• " + note["text"], y, size=9.5, color=RED if note.get("review") else DARK, line_height=14)
            image = Path(item["drawing_image"])
            pix = pymupdf.Pixmap(str(image))
            max_w, max_h = 511, 590
            scale = min(max_w / pix.width, max_h / pix.height)
            w, h = pix.width * scale, pix.height * scale
            rect = pymupdf.Rect(42 + (511 - w) / 2, y + 8, 42 + (511 + w) / 2, y + 8 + h)
            self.page.insert_image(rect, filename=str(image))
        self.footer()
        return self.doc


def w_text_runs(value, size, bold=False, color="222222"):
    runs = []
    parts = str(value or "").split("\n")
    for i, part in enumerate(parts):
        if i:
            runs.append("<w:r><w:br/></w:r>")
        props = f'<w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Microsoft YaHei"/><w:sz w:val="{int(size*2)}"/><w:color w:val="{color}"/>' + ("<w:b/>" if bold else "") + "</w:rPr>"
        runs.append(f'<w:r>{props}<w:t xml:space="preserve">{escape(part)}</w:t></w:r>')
    return "".join(runs)


def w_p(value="", size=9.5, bold=False, align=None, color="222222", after=40, before=0):
    jc = f'<w:jc w:val="{align}"/>' if align else ""
    return f'<w:p><w:pPr><w:spacing w:before="{before}" w:after="{after}" w:line="270" w:lineRule="auto"/>{jc}</w:pPr>{w_text_runs(value,size,bold,color)}</w:p>'


def w_page_break():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def w_tc(content, width, fill=None, span=1):
    shd = f'<w:shd w:fill="{fill}"/>' if fill else ""
    grid = f'<w:gridSpan w:val="{span}"/>' if span > 1 else ""
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{grid}{shd}<w:vAlign w:val="center"/></w:tcPr>{content}</w:tc>'


def w_table(rows, widths, borders=True, header=False):
    border_xml = '<w:tblBorders>' + ''.join(f'<w:{x} w:val="single" w:sz="4" w:color="BFBFBF"/>' for x in ["top","left","bottom","right","insideH","insideV"]) + '</w:tblBorders>' if borders else '<w:tblBorders><w:nil/></w:tblBorders>'
    grid = ''.join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    body = []
    for ri, row in enumerate(rows):
        cells = []
        for cell in row:
            value = cell.get("content", "")
            if not value.startswith("<w:"):
                value = w_p(value, cell.get("size", 8.5), cell.get("bold", False), cell.get("align"), cell.get("color", "222222"), 0)
            cells.append(w_tc(value, cell.get("width", widths[0]), cell.get("fill"), cell.get("span", 1)))
        body.append('<w:tr>' + ''.join(cells) + '</w:tr>')
    return '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>' + border_xml + '<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>' + grid + '</w:tblGrid>' + ''.join(body) + '</w:tbl>'


def build_docx(spec, output):
    images = []
    def add_image(path):
        images.append(Path(path))
        return f"rId{len(images)+2}"

    logo = Path(__file__).resolve().parents[1] / "assets" / "logo-runsun.png"
    logo_rid = add_image(logo)
    drawing_rids = [add_image(i["drawing_image"]) for i in spec["items"]]

    def image_xml(rid, name, width_px, height_px, width_pt, docpr):
        height_pt = width_pt * height_px / width_px
        cx, cy = int(width_pt * 12700), int(height_pt * 12700)
        return f'''<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{docpr}" name="{escape(name)}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="{docpr}" name="{escape(name)}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'''

    from PIL import Image
    with Image.open(logo) as im: logo_size = im.size
    body = []
    company = w_p("上海如榫工业科技有限公司", 12, True, "right", after=0) + w_p("中国上海市青浦区香大东路1318号1幢6层", 8.5, align="right", color="777777", after=0) + w_p("电话：+86 18006193111", 8.5, align="right", color="777777", after=0)
    body.append(w_table([[{"content": image_xml(logo_rid,"Runsun logo",*logo_size,108,1)}, {"content":company}]], [2200,6800], False))
    body.append('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="18" w:color="F58220"/></w:pBdr></w:pPr></w:p>')
    body.append(w_p("样品试制（打样）协议",16,True,"center",after=140))
    party_rows = [
        [{"content":f"甲方（需方）：{spec['buyer']}"},{"content":f"协议编号：{spec['ref']}"}],
        [{"content":f"乙方（供方）：{spec['supplier']}"},{"content":f"签订地点：{spec['sign_place']}"}],
        [{"content":""},{"content":f"签订日期：{spec['date']}"}],
    ]
    body.append(w_table(party_rows,[5000,4000],False))
    body.append(w_p("甲方委托乙方按甲方提供的图纸试制样品。根据双方沟通，经协商一致，就本次打样事宜签订本协议，双方共同遵守。",before=100))
    body.append(w_p("一、样品内容及费用",11,True,before=100))
    rows = [[{"content":h,"bold":True,"align":"center","fill":"FFF1E8"} for h in ["产品名称","图纸号","规格","材质","样品数量\n（个）","打样费包干价\n（元，未税）","大货单价\n（元/个，含税）"]]]
    for item in spec["items"]:
        rows.append([{"content":str(v)} for v in [item["product"],item["drawing_no"],item["spec"],item["material"],item["quantity"],item.get("sample_fee") or "",item.get("mass_price") or ""]])
    rows.append([{"content":"打样费未税合计（大写）：" + str(spec.get("total_fee_words") or "________________________"),"span":5,"width":4900},{"content":spec.get("total_fee") or ""},{"content":""}])
    body.append(w_table(rows,[1050,950,1350,1350,800,1350,1250]))
    body.append(w_p("1. 打样费为一次性包干价（未税）；大货单价为含税价，价格口径（勾选）：" + spec["price_options"] + "。打样费已包含工装、刀具、原材料、加工、检测及包装等费用；选择含运时同时包含运费，乙方不再另行收费。",9))
    body.append(w_p("2. 模具/工装费用：" + str(spec.get("tooling_note") or "________________") + "；未填写或未注明另计的，视为已包含在打样费内。",9))
    body.append(w_p("3. 表中大货单价为样品合格后大货订单的确认执行价；除双方约定的材料价格波动情形外，乙方不得拒绝接受或另行加价。",9))
    if spec.get("material_price_note"): body.append(w_p("4. " + spec["material_price_note"],9))
    body.append(w_p("二、技术与质量要求",11,True,before=100))
    for i, req in enumerate(spec["technical_requirements"],1): body.append(w_p(f"{i}. {req}",9))
    body.append(w_p("三、交样",11,True,before=100))
    for i,key in enumerate(["delivery_quantity","delivery_date","delivery_method"],1): body.append(w_p(f"{i}. {spec[key]}",9))
    body.append(w_p("4. 包装：样品清洁后按图号分别包装，防止磕碰、氧化和混料，外包装注明图纸号、数量及批次。",9))
    body.append(w_page_break())
    body.append(w_p("四、付款",11,True)); body.append(w_p(spec["payment_terms"]))
    body.append(w_p("五、合样、验收与退款",11,True,before=100))
    for i,req in enumerate(spec["acceptance_terms"],1): body.append(w_p(f"{i}. {req}"))
    body.append(w_p("六、批量生产",11,True,before=100)); body.append(w_p(spec.get("mass_production_terms") or "样品经甲方书面确认合格后，双方另行签订批量采购合同。大货订单执行本协议确认的单价；产品的材料、尺寸、外观及关键工艺须与确认样品一致，任何变更须经甲方书面同意。"))
    body.append(w_p("七、保密",11,True,before=100)); body.append(w_p("乙方对甲方提供的图纸、技术资料及本协议内容负有保密义务，未经甲方书面同意，不得向任何第三方披露，不得为甲方以外的任何方生产、销售该产品，也不得直接或间接与该产品的最终用户联系。本条义务不因本协议终止而失效。"))
    body.append(w_p("八、其他",11,True,before=100))
    body.append(w_p("1. 违约责任按《中华人民共和国民法典》有关规定执行。本协议未尽事宜由双方协商解决；协商不成的，向甲方所在地人民法院起诉。"))
    body.append(w_p("2. 附件“产品图纸 " + "、".join(i["drawing_no"] for i in spec["items"]) + " 及中文技术要求”为本协议不可分割的组成部分。"))
    body.append(w_p("3. 本协议一式两份，双方各执一份，经双方盖章后生效。盖章后的扫描件、传真件与原件具有同等效力。"))
    left = "甲方（需方）\n单位名称（章）：上海如榫工业科技有限公司\n单位地址：上海市青浦区香大东路1318号1幢6层\n法定代表人：何阳\n委托代理人：\n电  话：18006193111\n开户银行：中国银行上海市徐泾支行\n账  号：453385565724\n税  号：91310118MACTKTHT9T\n日  期：    年  月  日"
    right = f"乙方（供方）\n单位名称（章）：{spec['supplier']}\n单位地址：\n法定代表人：\n委托代理人：\n电  话：\n开户银行：\n账  号：\n税  号：\n日  期：    年  月  日"
    body.append(w_table([[{"content":left,"size":8.5},{"content":right,"size":8.5}]],[4500,4500]))
    docpr = 2
    for item,rid in zip(spec["items"],drawing_rids):
        body.append(w_page_break())
        body.append(w_p(f"附件：产品图纸 {item['drawing_no']} 及中文技术要求",12,True))
        body.append(w_p(f"协议编号：{spec['ref']}",8.5,color="777777"))
        for note in item["annex_notes"]: body.append(w_p("• "+note["text"],9.5,color="C00000" if note.get("review") else "222222"))
        with Image.open(item["drawing_image"]) as im: size = im.size
        body.append(image_xml(rid,f"{item['drawing_no']}.png",*size,510,docpr)); docpr += 1

    sect = '<w:sectPr><w:footerReference w:type="default" r:id="rId2"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="840" w:bottom="840" w:left="840" w:header="360" w:footer="360"/></w:sectPr>'
    document = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' + ''.join(body) + sect + '</w:document>'
    styles = '<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Microsoft YaHei"/><w:sz w:val="19"/></w:rPr></w:style></w:styles>'
    footer = '<?xml version="1.0" encoding="UTF-8"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' + w_p(f"上海如榫工业科技有限公司  ·  上海市青浦区香大东路1318号1幢6层  ·  {spec['ref']}",7.5,align="center",color="888888") + '</w:ftr>'
    rels = ['<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>']
    for idx,image in enumerate(images,3): rels.append(f'<Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image{idx-2}.png"/>')
    rels.append('</Relationships>')
    content_types = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'
    root_rels = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    with ZipFile(output,"w",ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",content_types); z.writestr("_rels/.rels",root_rels)
        z.writestr("word/document.xml",document); z.writestr("word/styles.xml",styles); z.writestr("word/footer1.xml",footer); z.writestr("word/_rels/document.xml.rels",''.join(rels))
        for idx,image in enumerate(images,1): z.write(image,f"word/media/image{idx}.png")
    with ZipFile(output) as z:
        assert z.testzip() is None
        for name in z.namelist():
            if name.endswith(".xml") or name.endswith(".rels"): ET.fromstring(z.read(name))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("config")
    parser.add_argument("--out-dir")
    parser.add_argument("--force",action="store_true")
    args = parser.parse_args()
    spec = json.loads(Path(args.config).read_text(encoding="utf-8"))
    out = Path(args.out_dir or spec["out_dir"]); out.mkdir(parents=True,exist_ok=True)
    tag = "-"+spec["output_tag"] if spec.get("output_tag") else ""
    base = f"{spec['ref']}（{spec['supplier_short']}）打样协议带图纸{tag}"
    docx, pdf = out/(base+".docx"), out/(base+".pdf")
    if not args.force and (docx.exists() or pdf.exists()): raise FileExistsError("Output exists; use --force to replace a generated draft")
    document = PdfBuilder(spec).build()
    document.subset_fonts(); document.save(pdf,garbage=4,deflate=True); document.close()
    build_docx(spec,docx)
    print(docx); print(pdf)


if __name__ == "__main__":
    main()
