---
name: outlook-email
description: Working rules for the outlook-web MCP tools (Outlook on the web through a signed-in browser). Use when reading, triaging, drafting, replying to or sending work email, or checking the work calendar.
---

# Outlook email through outlook-web-mcp

## Reading and triage

1. `outlook_folders` with `unread_only: true` shows where the unread mail is (including subfolders of Deleted Items and rule-sorted folders).
2. `outlook_read_unread` (no folder) reads all of it and marks it read. For a big backlog, go folder by folder.
3. To find something, use `outlook_search` (`from:name`, `subject:words`) or `outlook_list` with `match`. Pinned conversations are listed separately and don't use up `limit`.
4. Report what needs action, by date, and leave newsletters and notifications out unless asked.

## Drafting

- Write replies with `outlook_draft_reply` (`mode`: `reply`, `reply_all` or `forward`) and new mail with `outlook_draft_new`. Both only save a draft and return its id.
- Match the user's voice from their sent mail (`outlook_list` on `Sent Items`). Keep it short. Their signature is added by Outlook; don't type one.
- Use `reply` by default. Use `reply_all` only when everyone on the thread needs the answer.

## Sending

- Send only when the user has said to send **that** draft in this conversation. "Draft a reply" is not permission to send.
- Before sending, show the draft with `outlook_read_draft` (recipients, subject, text) unless the user has just read it.
- Then `outlook_send_draft` with the draft id and part of its subject. If sending is disabled, say so; the user can send it from Outlook.

## Housekeeping

- `outlook_mark` with `delete` or `archive` only when the user asks. Throw away a mistaken draft with `outlook_discard_draft`.
- If a tool says the session expired, call `outlook_login` and pass the number it returns to the user to approve in their authenticator app, then call it again with `wait_seconds: 60`.
- If something looks wrong, `browser_screenshot` shows the page.
