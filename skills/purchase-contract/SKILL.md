---
name: purchase-contract
description: Create a readable Chinese Runsun purchase contract for a confirmed factory order, with validated RMB totals, XLS/PDF output and optional drawing annexes. Use for “制作采购合同”, “做工厂合同”, “更新采购合同模板”, or a supplier PO for an existing customer order.
description_zh: 根据客户 PO/PI、工厂报价或确认条款生成中文外贸产品采购合同，输出 XLS、PDF 和可选含图纸 PDF。
version: 1.1.1
author: cwb-dev
agent_created: true
---

# Purchase Contract

`<skill dir>` means this skill's folder. Read the workspace `CLAUDE.md` and relevant private client profile when present. The contract is Chinese; retain original English product names and bilingual carton-label fields where required.

## First-use guidance

On the first invocation in a conversation, tell the user in Chinese before doing the work:

> 推荐用法：把客户 PO/PI、对应询价或工厂报价，以及工厂确认的付款和交期一起给我。我会先核对产品图号、数量、含税单价和总额，再生成中文采购合同的 XLS + PDF；有图纸时可同时生成含图纸 PDF，并渲染检查版式。已有合同只作为字段和条款参考，不会未经确认改动价格、数量、付款或交期。

## Layout and template

The template is `<skill dir>/templates/PC-RUNSUN.xlsx`. To change the master layout, edit `<skill dir>/build_template.py`, rebuild the template and generate a sample before use. Do not edit the template for an individual order; order-specific content belongs in a JSON spec.

Preserve the original item columns, units, packaging, carton labels, commercial terms and complete party/bank/signature fields. The JSON format is documented at the top of `make_pc.ps1`.

## Source-of-truth rules

1. Customer PO/approved PI establishes the customer order, part number, quantity and delivery requirement. Supplier quotation, email/WeChat confirmation or an existing factory contract establishes procurement price and factory terms.
2. Never use the customer selling price as the factory purchase price. Keep customer and factory currencies separate.
3. Preserve existing contract information whenever it remains valid. Do not silently change price, quantity, payment, delivery, acceptance, liability or tolerance.
4. A customer deposit does not prove the factory received a deposit. Do not mark a supplier advance as paid without factory-payment evidence.
5. If factory lead time starts at deposit receipt, calculate the factory delivery date and flag any conflict with the customer requirement.
6. Check the destination first. Do not create an unexplained duplicate or overwrite an existing formal contract without authorization.

## Numbering and destination

Use the sales PI number as the base contract number. When one PI is split among factories, use `-P1`, `-P2`, etc.; preserve an established number when an existing order uses another convention. Read the formal archive root and client naming from the private workspace configuration. Source documents are read-only except for adding a requested new output.

## Steps

1. Collect and reconcile contract number/date/place, both parties, every item, quantities and units, weights, tax-inclusive prices, packaging, labels, terms and optional drawing paths. Resolve or explicitly flag discrepancies.
2. Write `pc_spec.json` in the working/output folder. Use forward slashes in paths, including on Windows.
   - `qty_unit`: `KPCS` or `PCS`; `price_unit` must be `CNY/KPCS` or `CNY/PCS` respectively.
   - `expected_total`: independent source-document total; generation stops on mismatch.
   - `unit_weight`: KG/KPCS for KPCS or KG/PCS for PCS; `weight_total` is checked against quantity × unit weight.
   - `labels`: up to 6 entries of `[中文标签, English label, value]`.
   - `terms`: copy confirmed terms without adding numbers; the script numbers them.
   - `drawings`: optional original PDF paths, appended in listed order.
3. Generate on Windows with desktop Excel and PowerShell:
   ```sh
   powershell -ExecutionPolicy Bypass -File <skill dir>/make_pc.ps1 -Spec <pc_spec.json>
   ```
   Outputs are `<basename>.xls` and `<basename>.pdf`; with drawings, also `<basename>-含图纸.pdf`. The script validates totals and refuses accidental overwrite.
4. Reopen the XLS or read its values, render every PDF page and inspect it. Check Chinese characters, all item text, totals, labels, terms, signatures and drawing-page count. There must be no clipping, overlap or unreadably small type.
5. Archive or email only when the user separately asks. Never overwrite prior formal versions.
6. Update the private inquiry/order tracker when the workspace uses one.

## Platform support

The generator requires Windows, PowerShell and desktop Microsoft Excel because it uses Excel COM. Mounting shared storage on macOS provides file access but does not make this generator runnable natively on macOS; use a Windows environment for generation.

## Excel troubleshooting

If Excel reports `busy` (`0x800AC472`) or hangs on a hidden dialog, inspect `Win32_Process` for `EXCEL.EXE`. Stop only processes whose command line contains `/automation -Embedding`; never close the user's interactive Excel.
