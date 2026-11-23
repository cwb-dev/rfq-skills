---
name: inquiry-status
description: Check mail against client inquiry trackers and report new RFQs, client chasers, overdue replies and recently quoted inquiries. Use for open-RFQ, client-follow-up and daily/weekly inquiry status checks.
description_zh: 对照邮箱和询价台账，检查新询价、客户催报价、长期未回复及近期已报价项目。
---

# Inquiry status

Read the private workspace rules, then each in-scope `CLIENT.md` and `INQUIRY_LOG.md`. Use the mail account configured privately for this workflow.

## Steps

1. Search the inbox, relevant client folders, sent mail and configured CC folders for the requested window (default: last 14 days). Page through all results. Search each configured sender domain separately if the provider has no OR search.
2. Match mail to tracker rows by RFQ reference, conversation subject and part number. Read ambiguous messages without changing their read state.
3. Report four groups, most urgent first:
   - **New, not collected**: client, reference, subject, date and visible item count.
   - **Client chasing**: number of chasers and days since the first unanswered message.
   - **Open, no reply after 5+ days**.
   - **Quoted recently**: who quoted and when.
4. Update tracker rows only for status changes proven by the mail. Record a quote when the sent quote is found. Ask before marking anything won, lost or declined.

Keep the report short: one line per inquiry with subject/reference and date. Offer to run `new-inquiry` for uncollected RFQs. Never send, move, flag, delete or mark mail as read as part of a status check.
