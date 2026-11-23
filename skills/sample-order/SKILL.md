---
name: sample-order
description: Create a concise Chinese 样品试制（打样）协议 for a factory in the Runsun Chinese letterhead style, with technical/drawing confirmation, one tax-inclusive sample-fee field with EXW or delivered price basis, delivery, fitting/acceptance and failed-sample refund terms, signatures, and one annex page per drawing. Outputs editable DOCX and combined PDF with drawings. Use for 打样协议, 样品单, 打样确认单, sample order or sampling agreement.
---

# Sample order / 样品试制（打样）协议

Default format: a concise factory sampling agreement using the Runsun letterhead visual style in Chinese. Do not turn it into a long procurement contract or a multi-column quotation workbook.

Outputs in the inquiry folder:

- `<ref>（<factory>）打样协议带图纸.docx`
- `<ref>（<factory>）打样协议带图纸.pdf`

The PDF contains the agreement and all drawing annexes. Do not send files; the user reviews and sends them.

## Layout

- A4, Chinese Microsoft YaHei/DengXian style; Runsun logo at top left.
- Top right: `上海如榫工业科技有限公司`, registered address and phone only (no email); orange divider below. Keep addresses distinct: registered/unit address is `上海市青浦区香大东路1318号1幢6层`; office/sample-delivery address is `上海市青浦区徐乐路333号9幢7层`. Use the registered address in the letterhead, footer and signature `单位地址`; use the office address only for delivery.
- Title `样品试制（打样）协议`, then buyer/factory, reference, place and date.
- Keep the eight short sections used by the reference agreement:
  1. 样品内容及费用 — one table; product, drawing, specification, material, quantity, `打样费包干价（元，未税）` and `大货单价（元/个，含税）` per item.
  2. 技术与质量要求 — actual drawing controls, material, reports and unresolved tolerances.
  3. 交样 — quantity, date, delivery and separate labelled packaging.
  4. 付款 — only confirmed wording; otherwise leave a blank.
  5. 合样、验收与退款 — inspection, necessary mating-part fitting and customer confirmation; failed samples refunded; one free remake only with buyer approval.
  6. 批量生产 — separate written order at the mass-production price confirmed in this agreement; mass production matches the approved sample.
  7. 保密.
  8. 其他 — attachment and signature effectiveness.
- Signature blocks follow the reference agreement. Do not insert a stamp unless the user provides one and asks.
- Each drawing gets one annex page: a few Chinese requirement bullets above the unchanged drawing. Avoid a separate long Chinese requirements page unless the user asks.

## Commercial and technical rules

1. Use one tax-exclusive lump-sum `打样费包干价` and one tax-inclusive `大货单价` per item; do not add parallel tooling/sample-fee columns. State a selectable price basis such as `□ 含税出厂  □ 含税含运至上海`, applying to the mass-production price. If delivered is selected, freight is included; if EXW is selected, record pickup/delivery separately.
2. Leave sample fees, mass-production prices, total, tooling, payment and lead time blank when requested. Do not copy a customer-facing sample charge or target price into the factory agreement unless the user explicitly identifies the screenshot amount as the price to use.
3. Tooling is included in the lump-sum sample fee unless a separate amount is explicitly written. The mass-production price is confirmed in this agreement and remains the accepted order price after sample approval; the factory must not reject it or add charges. State the applicable quantity and batch-delivery basis. For copper products, optionally add `material_price_note`: the factory fills the tax-inclusive copper basis in `________元/kg`; if copper fluctuates significantly when the mass order is placed, both parties may renegotiate according to the material movement. Treat this as the stated exception to the fixed-price rule.
4. State actual material and key controls in plain Chinese. Never invent missing tolerances or resolve unclear drawing variants by guesswork. Put unresolved items in red on the affected annex page and keep the document marked `初版` or `待确认`.
5. Default failed-sample wording: full refund of the affected sample fee after failed inspection/fitting; with buyer approval, one free remake; repeat shipping caused by factory failure is borne by the factory.
6. Keep client names, internal target prices and costs out of factory-facing text. Original drawing graphics stay unchanged.

## Build

Requires Python with `pymupdf` and `Pillow`. The generator creates the PDF and standard DOCX directly; Microsoft Word is not required. It uses the included `assets/logo-runsun.png`.

1. Read project/folder rules and source drawings. Render each selected PDF drawing page to PNG at 200 dpi in a temporary `<work>` directory with PyMuPDF. Do not put preview images in the inquiry folder.
2. Write a UTF-8 JSON spec using `example_config.json`. Each item's `drawing_image` points to its rendered PNG; `annex_notes` contains the short Chinese notes. Set `review: true` on an unresolved note to print it in red.
3. Run:

```bash
python "skills/sample-order/scripts/build_sample_order.py" <spec.json>
```

Use `--force` only to replace a draft generated by this skill. The script refuses to overwrite existing outputs otherwise.

## Verify

- Open the PDF and inspect every page: Chinese letterhead, orange divider, item quantities, blank/filled fees, all eight sections, signature blocks and drawing orientation.
- Check Drawing 1/2/3 or other annexes are in the intended order and no text/image is clipped.
- Confirm unresolved technical items are visible and no price was filled accidentally.
- Save the verified DOCX and combined PDF to the inquiry folder. Do not mark the order sent or production-released.
