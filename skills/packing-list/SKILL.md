---
name: packing-list
description: Create a Packing List (PL) in the Runsun Industrial or CWB Industrial company format from an actual factory packing list and confirmed shipment details, as .xls plus PDF. Use when the user says "make a PL", "packing list", "装箱单", "箱单", or needs shipment packing documentation.
description_zh: 根据工厂实际装箱单和确认出货信息生成 Runsun 或 CWB 抬头的装箱单 PL，校验数量、箱数、毛净重和体积，输出 XLS 和 PDF。
description_en: Generate a Runsun or CWB packing list from actual packing evidence, with quantity, carton, weight and volume checks plus XLS/PDF output.
version: 1.2.0
author: cwb-dev
agent_created: true
---

# Packing List

`<skill dir>` means this skill's folder. Read `CLAUDE.md` and the relevant client profile first. PLs are always in **English**.

## First-use guidance

On the first invocation in a conversation, tell the user in their language before doing the work:

> 推荐用法：把工厂实际装箱单、对应客户 PI/合同和出货信息一起给我。PL 数量始终以实际箱单为准，合同只用于核对产品、规格和参考数量；若数量偏差过大，我会明确向你反馈，但不会用合同数量覆盖实际箱单。我还会逐项核对包装、箱数、毛重、净重和尺寸。生成前，我会把托盘顺序号和唛头预览发给你标记、补充并明确确认；未经你确认不会生成正式 PL。确认后生成 XLS + PDF，并渲染检查合计和版面。净重与毛重必须按工厂实际数据填写，不会自行估算。

## Sellers and templates

| `seller` | Company | Template |
|---|---|---|
| `runsun` | RUNSUN INDUSTRIAL TECHNOLOGY CO., LTD | `<skill dir>/templates/PL-RUNSUN.xls` |
| `cwb` | CWB INDUSTRIAL PTE. LTD. | `<skill dir>/templates/PL-CWB.xls` |

Use the same seller and buyer as the matching CI. Never edit a template for an individual shipment; shipment-specific data belongs in the JSON spec.

## Source-of-truth rules

- Use the factory's **actual/final packing list**, not a proforma packing plan. The corresponding purchase contract helps identify the order and products, but final shipped quantities and packing come from the actual packing list.
- Preserve the source values for quantity, unit weight, quantity per carton, carton count, gross/net weight and carton dimensions. Do not estimate missing weights or dimensions.
- `qty` and `qty_per_ctn` use `qty_unit`; `unit_weight` uses `unit_weight_unit`. For example, with `KPCS`, a quantity of 10 means 10,000 pieces and unit weight is kg per 1,000 pieces.
- `gross` and `net` are totals for the whole line. When available, item `cbm` is the total volume for that line, not one carton's volume. If the source provides only a shipment-level volume, leave item `cbm` empty and set `total_cbm`; do not distribute the total arbitrarily across lines. Calculate volume only when the source dimensions and package count support that calculation.
- Gross weight must not be below net weight. Apply client-specific rules: for CMT, gross and net weights must differ and CI/PL quantities must match the PI exactly.
- Keep CI and PL product order, descriptions, quantities and references consistent. Reconcile partial shipments and combined PIs line by line.

## Actual quantities govern; contracts are references only

- `items[].qty` and `expected_totals.qty` always come from the actual packing list. Never replace, round or balance them to match a purchase contract, customer PI or expected order quantity.
- The contract is used only to identify the order, products/specifications and reference quantities. Overproduction, under-shipment and split deliveries remain visible as differences; the PL reports what was actually packed.
- Match actual lines to contract lines by product, specification and allocation rather than row position. For repeated products or multiple factories/contracts, preserve the line-level mapping and compare subtotals as well as the shipment total.
- Put the matched reference in `contract_qty` using the same unit as `qty`; `contract_total_qty` is optional for the overall comparison. These fields are never printed as PL quantities and never participate in PL totals.
- Before generation, report any mismatch to the user. The default large-difference warning is absolute variance **over 5%** of contract quantity; set `quantity_warning_percent` when the user or contract specifies a different threshold. Contract-specific rules still apply separately, including asymmetric rules such as no short shipment.
- A large difference must be prominently fed back to the user: actual quantity, contract quantity, difference and percentage. Ask for instructions when necessary, but do not silently block or “correct” the actual packing data. If the user confirms the actual shipment, keep the actual values.
- Relay every script `WARNING` in the final response. The warning is a reconciliation aid, not permission to change the source quantity.

## Mandatory pallet-order and shipping-mark confirmation

The user, not the agent, owns the physical pallet sequence and final shipping mark.

1. Before writing the final spec, show the user a preview with each physical position, proposed pallet number, included products/cartons and the exact shipping-mark text/line breaks. If several factories supply the shipment, show every factory's packages in that preview.
2. Ask the user to mark missing pallet numbers and correct or complete the shipping mark. Source documents and previous shipments may provide a draft, but they do **not** count as user confirmation.
3. Proceed only after the user explicitly confirms both. Then set `pallet_sequence_confirmed: true` and `shipping_mark_confirmed: true`. Never set these flags preemptively, by assumption, or merely because the source packing list already contains numbers/marks.
4. Put the confirmed physical order in `pallet_sequence`, for example `["1", "2", "3"]`. Every item's `pallet_no` must belong to that list, and item rows must be grouped and sorted accordingly. For a confirmed non-palletized shipment, use a user-confirmed placeholder such as `["N/A"]` and `pallet_no: "N/A"`.
5. Put the final confirmed mark in `shipping_mark`, preserving line breaks. If marks differ by pallet, include the pallet number with each mark in this block.
6. If the user changes the sequence or mark, update the spec and regenerate both XLS and PDF; do not patch only one output. The script refuses missing confirmation, unknown pallet numbers and out-of-order rows.

## Numbering

The PL normally uses the matching CI number `<YY>RSIT-<country>-<CLIENT>-<NNNN>`. For an isolated test, use a number such as `TEST-PL-26RSIT-RU-CMT-0009` and write only to a separate test folder.

## Steps

1. Collect the buyer, matching CI number/date, PI/contract references, all actual packing lines and confirmed notes from the source documents. Draft the pallet sequence and shipping mark, then complete the mandatory user-confirmation step above before generation.
2. Independently total the **actual packing-list** quantity, cartons, gross weight, net weight and CBM. Put the actual source document's totals in `expected_totals`; the script rejects any mismatch. Separately compare `contract_qty`/`contract_total_qty` and report differences without changing the actual totals.
3. Write `<output folder>/pl_spec.json` using the format at the top of `make_pl.ps1`. Use the Write tool and forward slashes in `out_dir`.
   - `buyer`: up to 4 lines; `refs`: up to 3 label/value pairs.
   - `pallet_sequence` and `shipping_mark` must reproduce the user's explicit confirmation; both confirmation flags must be `true`.
   - Each item requires `pallet_no`, `qty`, `ctns`, `gross` and `net`. Item `cbm` is optional when `total_cbm` supplies the source's overall volume.
   - `qty` is always actual. Optional `contract_qty`, `contract_total_qty` and `quantity_warning_percent` drive reconciliation warnings only.
   - `measurement` is the carton/package dimensions as text, for example `60 X 40 X 40`.
   - `packing_summary` is optional; otherwise the script uses total cartons plus `CTNS`.
4. Generate:
   ```sh
   powershell -ExecutionPolicy Bypass -File <skill dir>/make_pl.ps1 -Spec <pl_spec.json>
   ```
   The script writes `<invoice_no>.xls` and `<invoice_no>.pdf`, fills item and summary totals, checks quantity/carton formulas, and refuses to overwrite unless `-Force` is explicitly appropriate.
5. Render every PDF page with PyMuPDF and inspect it. Check pallet numbers, product descriptions, units, gross/net order, dimensions, shipping mark, summary block, notes and text clipping. Compare all totals with the actual factory packing list.
6. Save into the formal shipment folder only when the user asks. Never overwrite existing CI/PL/customs documents.
7. If a covering email is requested, create a 263 draft only, never send. Follow `CLAUDE.md` for threading, recipients and attachments.

## Excel troubleshooting

If Excel reports `busy` (`0x800AC472`) or hangs on a hidden dialog, inspect `Win32_Process` for `EXCEL.EXE`. Stop only processes whose command line contains `/automation -Embedding`; never close the user's interactive Excel.
