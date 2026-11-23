---
name: mail-check
description: Check configured business mail for everything received since the last successful check, start or propose collection for new RFQs, log updates on existing inquiries and write a daily digest. Use for scheduled or manual incremental mail checks.
description_zh: 增量检查业务邮箱，识别新询价、更新询价台账并生成每日摘要；默认不修改邮件状态。
---

# Mail check (since last time)

Read the private workspace `CLAUDE.md` and every in-scope client profile/tracker first. They define the account, folders, sender domains and local policies. This skill only reads mail and writes tracking files. It never sends, moves, flags, deletes or marks mail as read.

## Files

Keep these under the private workspace `Mail-Check/` directory:

- `state.json`: `last_run` (start of the last successful run), recent `seen` message IDs and new-inquiry `sessions`. Retain enough history to deduplicate overlapping windows and recent sessions.
- `<YYYY-MM-DD>.md`: daily digest. If run twice in one day, add a new run section while preserving the earlier results.

## Steps

1. Record the run start time. Use `last_run` minus a one-hour overlap; cap an excessively long first/backfill window according to workspace policy. Deduplicate by message ID.
2. Scan the inbox, sent mail, configured CC folders and client folders. Skip shipment-only, spam, archive and other folders excluded by workspace policy. Page through all results.
3. Classify mail into new inquiry, update on an existing inquiry, other client business, or ignore. Match existing inquiries by RFQ reference, subject and PN. Factory quotes can be updates when reliably matched.
4. For each new inquiry, start one `new-inquiry` session when the host supports it; otherwise add a clear proposed action to the digest. Do not start duplicates already recorded in `state.json`.
5. Log proven updates in the relevant private tracker. Never mark won/lost/declined unattended; put possible outcomes in the digest for user review.
6. Write the digest with urgent user actions, new inquiries/sessions, tracker updates and other client business. Use the local business timezone.
7. Save `last_run` only after all tracking and digest writes succeed. If the run fails part-way, leave the previous successful timestamp intact.
8. Finish with a short summary of new inquiries, chasers, updated trackers and the digest path.

A scheduler may invoke this skill daily, but scheduler configuration and mail credentials are private machine setup, not part of the public repository.
