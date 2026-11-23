---
name: commercial-invoice
description: Create a Commercial Invoice (CI) in the Runsun Industrial (CNY/VTB) or CWB Industrial (USD/ANEXT) company format from a confirmed shipment, as .xls plus PDF. Use when the user says "make a CI", "commercial invoice", "报关发票", or needs a customer shipment invoice.
description_zh: 根据确认的出货资料生成 Runsun 或 CWB 抬头的商业发票 CI，输出 XLS 和 PDF；适用于“做 CI”“商业发票”“报关发票”等请求。
description_en: Generate a Runsun or CWB commercial invoice from confirmed shipment evidence, with validated totals and XLS/PDF output.
version: 1.1.0
author: cwb-dev
agent_created: true
---

# Commercial Invoice

`<skill dir>` means this skill's folder. Read `CLAUDE.md` and the relevant client profile first. CIs are always in **English**.

## First-use guidance

On the first invocation in a conversation, tell the user in their language before doing the work:

> 推荐用法：把本次出货涉及的一份或多份客户 PI/销售合同、所有工厂的实际装箱单和出货信息一起给我。我会先建立“CI 行—PI—工厂箱单”的逐行对应关系，核对买方、发票号、数量、客户单价、贸易条款和金额，再生成 CI 的 XLS + PDF，并渲染 PDF 检查。工厂采购合同可用于核对订单和工厂成本，但不能作为客户销售单价来源。

## Sellers and templates

| `seller` | Company | Bank | Currency | Template |
|---|---|---|---|---|
| `runsun` | RUNSUN INDUSTRIAL TECHNOLOGY CO., LTD | VTB Bank Shanghai | CNY | `<skill dir>/templates/CI-RUNSUN.xls` |
| `cwb` | CWB INDUSTRIAL PTE. LTD. | ANEXT Bank | USD | `<skill dir>/templates/CI-CWB.xls` |

The seller must match the sales contract/PI and the payment beneficiary. Never edit a template for an individual shipment; order-specific data belongs in the JSON spec.

## Source-of-truth rules

- Final customer CI lines govern actual shipped quantities and sales amounts. Use the final factory packing list to confirm shipped quantities, but reconcile many-to-many PI/CI allocations instead of assuming a whole PI shipped.
- **Customer unit prices come from the approved customer PI/sales contract, never from a factory purchase contract.** Purchase prices are factory costs and must not leak into a customer CI.
- If the user explicitly requests a layout test using a factory packing list and purchase contract, keep every output in a separate test folder, use a test-only document number, do not overwrite shipment records, and state that the prices are procurement prices used only for the test.
- Preserve currency and units. Runsun Russian/Belarus orders normally use `MPCS`, `CNY/MPCS`, and `KG/MPCS`; per-piece orders may use `PCS`, `CNY/PC`, and `KG/PCS`. CWB orders normally use `PCS`, `USD/PC`, and price format `\"$\"#,##0.00000`.
- Do not invent buyer details, references, Incoterms, payment terms, weights or shipping information. Read them from the source documents.

## Multiple PIs and factory packing lists

A CI may combine several customer PIs and packing lists from several factories. Never assume a one-PI/one-factory shipment.

1. Read **every** approved PI and **every** actual factory packing list belonging to the shipment. Build a line-allocation worksheet before writing the spec: CI line, product/PN, PI number, factory, factory packing-list row, shipped quantity, customer unit price and amount.
2. Allocate quantities line by line. A repeated PN can appear on different PIs or at different customer prices; do not attach all candidates automatically or merge them merely because the PN matches.
3. Combine rows only when product, description, unit, customer unit price and all printed fields are identical, and the complete PI/factory allocation remains traceable in the worksheet. Otherwise keep separate CI rows.
4. Reconcile both directions: each factory packing-list line is included exactly once, and each PI line's cumulative shipped quantity is checked against its approved quantity. Flag partial shipment, over-shipment or cancellation instead of forcing totals to match.
5. Combine PIs on one CI only when seller, buyer, currency and Incoterm are compatible and the customer/shipment documents allow the combination. Otherwise split the documents.
6. The JSON `refs` area has only 3 rows. For multiple PIs, use a compact complete list only if it fits without clipping; otherwise use `PI NOS.: SEE NOTES` and list every PI number in the notes. Keep the detailed allocation worksheet with the working papers even though the CI template does not print factory names.
7. Check the CI grand total against the sum of the allocated customer PI amounts, not factory purchase totals. Check quantity/weight against the combined factory packing lists.

## Numbering

Normal CI numbers are `<YY>RSIT-<country>-<CLIENT>-<NNNN>`, one per shipment/customs declaration, for example `26RSIT-XX-ABC-0001`. Read the existing shipment archive root from the private workspace configuration and list only the relevant client folder. A PI can ship across several CIs and one CI can cover several PIs.

For an isolated test, use a number such as `TEST-CI-26RSIT-XX-ABC-0001` rather than a formal CI number.

## Steps

1. Collect from all approved customer documents and final shipment evidence: buyer legal name/address, CI number/date, every PI and contract reference, every shipped line (product, description, quantity, customer unit price, unit weight where used), Incoterm, currency, notes and country of origin. Complete the multi-PI/multi-factory allocation worksheet above first.
2. Compare line quantities with every actual packing list and the amount with `expected_total`. For combined shipments, independently calculate the sum across all source documents. Resolve discrepancies before generating; do not silently force totals to match.
3. Write `<output folder>/ci_spec.json` using the format at the top of `make_ci.ps1`. Use the Write tool and forward slashes in `out_dir`. Keep source PI/factory row identifiers in the separate allocation worksheet; do not put factory names or purchase prices on the customer CI.
   - `buyer`: up to 4 lines.
   - `refs`: up to 3 label/value pairs.
   - `expected_total`: optional independent source-document total; the script rejects a mismatch.
   - `notes`: payment, delivery/Incoterm, origin and other confirmed terms, numbered automatically.
4. Generate:
   ```sh
   powershell -ExecutionPolicy Bypass -File <skill dir>/make_ci.ps1 -Spec <ci_spec.json>
   ```
   The script writes `<no>.xls` and `<no>.pdf`, fills formulas, adds the amount in words, checks the sheet total, and refuses to overwrite unless `-Force` is explicitly appropriate.
5. Render every PDF page with PyMuPDF and inspect it. Confirm one page for ordinary shipments, correct buyer/refs/units, no clipped item or note text, and totals matching the source. Test multi-page output when a shipment has many lines.
6. Save into the formal shipment folder only when the user asks. Never overwrite existing CI/PL/customs documents.
7. If a covering email is requested, create a 263 draft only, never send. Follow `CLAUDE.md` for threading, recipients and attachments.

## Excel troubleshooting

If Excel reports `busy` (`0x800AC472`) or hangs on a hidden dialog, inspect `Win32_Process` for `EXCEL.EXE`. Stop only processes whose command line contains `/automation -Embedding`; never close the user's interactive Excel.
