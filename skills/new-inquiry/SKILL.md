---
name: new-inquiry
description: Collect a client RFQ/inquiry email and its attachments into a new shared-drive inquiry folder, extract line items and update the client's inquiry tracker. Use when the user gives an RFQ number, client or inquiry email, or asks to collect new RFQs.
description_zh: 从邮箱收集客户询价及附件，建立共享盘询价档案、提取明细并更新询价台账。
---

# New inquiry

Read the private workspace `CLAUDE.md` and the relevant `CLIENT.md` first. They supply the mail account, sender domains, shared-drive root, client folder names and inquiry naming rules. This public skill intentionally contains no private account or client configuration.

## Steps

1. **Find the email** with the configured mail MCP tools. Search by RFQ reference, sender domain or date as appropriate. Search both the inbox and the client's mail folder; also inspect sent mail when the conversation may already have replies.
2. **Check it is not already collected.** Search the client tracker and the relevant shared-drive client folder. If it exists, only append new follow-ups; do not create a duplicate.
3. **Read the complete conversation**, including later replies. Never mark mail as read as a side effect.
4. **Create the inquiry folder** under the client's shared inquiry root, following that client's existing naming convention. Keep files flat. Ask before creating a new top-level client folder.
5. **Save the email** as `email_RFQ<no>_<YYYY-MM-DD>.txt`, with From, To/CC, date, subject, provider email ID and message ID. Append later messages with clear separators.
6. **Download attachments** with their original filenames. Never overwrite an existing attachment. Record duplicate-name or failed downloads explicitly.
7. **Extract line items** into `items.md`: `# | PN | Description | Qty (EAU) | Material | Finish/Spec | Drawing | Notes`. Keep repeated PNs and quantity tiers as separate rows.
8. **Record requirements** at the top of `items.md`: weights, lead time, PPAP, REACH/RoHS/CMRT, deadline, Incoterm and reply-all requirements when present.
9. **Update the private tracker** with status `new` and the inquiry path. Add newly confirmed client requirements to the client profile.
10. **Report** the folder, item count, requirements, deadline, missing attachments and any client chasers.

Do not draft replies or contact suppliers in this skill. Never move, flag, delete or mark mail as read unless the user separately authorizes it.
