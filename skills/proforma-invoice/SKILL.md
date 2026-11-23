---
name: proforma-invoice
description: Create a Proforma Invoice (PI) in the company format (Runsun Industrial for CNY/VTB orders, CWB Industrial Pte for USD orders) from a confirmed order, contract or PO, as .xls + PDF, numbered RSIT-<client code>-<YY>-<seq>. Use when the user says "make a PI", "proforma for this order", "invoice for the contract", or a client needs a PI to pay.
description_zh: 根据客户订单邮件、合同或 PO 及报价计算文件生成 Runsun 或 CWB 抬头的形式发票 PI，输出 XLS 和 PDF；适用于“做 PI”“形式发票”“客户要 PI 付款”等请求。
description_en: Generate a Runsun or CWB proforma invoice from confirmed order evidence and approved customer prices, with validated totals and XLS/PDF output.
version: 1.0.0
author: cwb-dev
agent_created: true
---

# Proforma Invoice

`<skill dir>` means this skill's folder. Read `CLAUDE.md` first. PIs are always in **English**.

## First-use guidance

On the first invocation in a conversation, tell the user in their language before doing the work:

> 推荐用法：把客户要 PI 的邮件、对应询价计算文件（Price to Client）以及客户合同/PO 一起给我。邮件用于确定本次订购的图号和数量，计算文件用于核对客户单价；买方法定信息、付款、交期和贸易条款以合同/PO或确认的客户资料为准。我会核对后生成 PI 的 XLS + PDF，并渲染 PDF 检查；合同与报价不一致时以合同为准。

## Sellers and templates
| `seller` | Company | Bank | Currency | Use for |
|---|---|---|---|---|
| `runsun` | RUNSUN INDUSTRIAL TECHNOLOGY CO., LTD (Shanghai) | VTB Bank (PJSC) Shanghai Branch, corr. VTB Moscow | CNY | Russia / Belarus clients, and any contract signed by Runsun |
| `cwb` | CWB INDUSTRIAL PTE. LTD. (Singapore) | ANEXT Bank, intermediary JP Morgan | USD | Contracts signed by CWB, normally USD orders |

The seller on the PI **must be the party named in the contract/PO**, since the bank checks that the PI, contract and beneficiary match.

Templates are `<skill dir>\templates\PI-RUNSUN.xls` and `PI-CWB.xls` (logo, letterhead and bank block are already in them). Never edit them by hand for one order; everything order-specific goes in the spec.

## Numbering
`RSIT-<CODE>-<YY>-<NNN>`, one sequence **per client** per year. Read the sales archive root and each client's existing folder/code convention from the private workspace configuration. Next number = highest existing for that client + 1; a new client starts at 001. Ask the user for a new client's three-letter code.

## Steps
1. **Collect the facts** from the contract/specification or PO, not from memory: buyer legal name + address (in English), contract/PO no. and date, every line (product, description, qty, unit price), Incoterm and place, payment terms, delivery time. If a quote and the contract differ, the **contract wins** (e.g. lots not awarded are left out). State coating/finish the way the contract does.
2. **Write the spec JSON** next to the order, `<order folder>/<PI no>/pi_spec.json` (format at the top of `make_pi.ps1`).
   - `buyer`: up to 3 lines (runsun) or 4 lines (cwb).
   - `refs`: up to 2-3 label/value pairs printed under NO/DATE (contract no., contract date, PO no.). Keep to 2 on runsun so they don't crowd the AMOUNT header; put anything else in a note.
   - Units: use the contract's units. Runsun/CNY orders (Russia/Belarus) use the Russian unit format `"qty_unit": "MPCS"`, `"price_unit": "CNY/MPCS"` (MPCS = 1000 pcs); CWB/USD orders `PCS` and `USD/PC` with `"price_format": "\"$\"#,##0.00000"`.
   - `notes`: payment, delivery (Incoterm + lead time), the contract reference, country of origin. Copy the contract's terms; don't improve them.
   - `buyer_sign_name`: the buyer's name printed above their signature line.
   - Write the JSON with the Write tool (not a bash heredoc, which can eat backslashes); use forward slashes in `out_dir`.
3. **Generate**:
   ```
   powershell -ExecutionPolicy Bypass -File <skill dir>\make_pi.ps1 -Spec <pi_spec.json> [-Force]
   ```
   It fills the template via Excel, inserts item/note rows as needed, writes ROUND(price×qty) formulas, adds "SAY TOTAL" in words, checks the sheet total against the spec, and saves `<PI no>.xls` + `<PI no>.pdf` in `out_dir`. It refuses to overwrite without `-Force`.
   - If Excel hangs or reports "busy" (0x800AC472), a hidden headless instance is stuck on a dialog: check `Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'"` and stop only processes whose command line has `/automation -Embedding` (those were started by scripts, not the user).
4. **Check the PDF**: render page 1 with PyMuPDF and look at it. One page, totals match the contract, no clipped text in item rows, refs not touching the table header.
5. **Copy to the configured sales archive** only if the user says so.
6. **Covering email**: create a **draft only** (never send) with the privately configured mail account, replying to the client's latest mail in their language and attaching the PDF by full path. Follow the private workspace rules for threading and To/CC. When the client must sign the PI back, say why.

## Notes
- The templates carry no signature stamp (removed in the 2026-11 redesign); the seller signs/stamps after printing if needed. The buyer signs and returns.
- Signed copies coming back are saved as `<PI no>-Signed.pdf` in the same folder.
