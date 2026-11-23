---
name: quote-compile
description: Compile factory quotes into a cost/margin workbook, supplier comparison, filled client RFQ form and a draft reply. Use when factory quotes have arrived or the user asks to make/compare a client quote.
description_zh: 汇总工厂报价，生成计算文件、报价对比、客户报价表及邮件草稿。
---

# Compile quote

Read the private workspace rules and relevant client profile first. Work in the inquiry folder on the configured shared drive. Keep client originals unchanged and follow the client's required form, units, quantity tiers and reply-all rules.

## Steps

1. **Read every factory reply**, including Excel files and images. Match every line to `items.md` by PN/specification and record deviations explicitly.
2. **Build the comparison** `<no>报价汇总_对比版.xlsx`: one row per client line and one column group per factory for RMB price, weight and lead time. Highlight the lowest valid price and flag no-quote/deviation lines.
3. **Ask the user** for exchange rate, margin/markup and category differences, Incoterm, tax/rebate treatment, packaging and freight assumptions. Do not guess these inputs.
4. **Write `计算文件-USD.xlsx`** with formulas: RMB cost → USD cost → margin → client price. Include unit weight and lead time, and keep rate/margin inputs in labeled cells.
5. **Fill the client form** on a copy with a date prefix. Fill only supplier columns such as unit price, weight, lead time, MOQ and deviations. Never change client columns/layout or drop quantity tiers.
6. **Write a clean internal quote summary** with the inquiry number and quote date.
7. **Create a draft reply only** when requested, in the client's language and with reply-all recipients preserved. Attach the filled form and summarize deviations/no-quote lines and lead time. Never send automatically.
8. **Update the private tracker**. Set `quoted` only after the user confirms the quote was sent.

After writing each workbook, verify that formulas have no errors and every client line has a price or an explicit no-quote/deviation result.
