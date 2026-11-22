---
name: outlook-email
description: Working rules for the outlook-web MCP tools (Outlook on the web through a signed-in browser). Use when reading, triaging, drafting, replying to or sending work email, or checking the work calendar, and signing back in (Okta or Microsoft MFA) when the Outlook session expires.
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

## Calendar

- `outlook_calendar` for a day or range; `outlook_event` (date + part of the title) for the details.
- Responding to an invitation (`outlook_event_respond`) emails the organiser: only when the user has said how to answer that invitation. Never "try out" a response.
- `outlook_event_create` without attendees just adds it to the calendar. With attendees it sends invitations, so confirm the title, time, location, Teams and invite list with the user first. Set `teams: true` for online meetings.
- `outlook_event_delete` on a meeting the user organised sends a cancellation: ask first.

## Signing in

Outlook sessions expire every so often. An expired page still looks normal apart from a red "You need to sign in. Your session has expired" bar, and nothing on it works.

1. **Spotting it:** a tool says "Not signed in" or "session expired", or `outlook_status` shows `signed_in: false`. If tools time out while status says signed in, take a `browser_screenshot` and look for the red bar.
2. **Start the sign-in:** call `outlook_login` (no `wait_seconds`). It uses the password saved in the browser and handles the account picker. The picker may show the account's username form (e.g. `uqxxxxx@uni.edu.au`) rather than the email address; that is the same account. It also switches an Okta "Enter a code" page to a push, then returns a number.
3. **Relay the number at once:** tell the user to tap or enter that number in their authenticator app (Okta Verify or Microsoft Authenticator). A worker agent sends it to whoever it reports to instead.
4. **Wait for approval:** call `outlook_login` with `wait_seconds: 90`. It returns "Signed in." once approved.
5. **Don't restart the sign-in over and over:** each fresh attempt can send a new push, which voids the old number. If it times out, ask the user whether they saw the prompt before starting again.
6. **If it returns "Stopped at …":** retry once, then send the user a `browser_screenshot` and say where it stopped. Never drive the sign-in pages with general browser tools, type MFA codes, or read the saved password. The user can sign in themselves through the browser's remote viewer.
7. **After signing in:** retry the tool that failed. A compose call that failed may have left a half-written draft, so check `Drafts` and remove a mistaken one with `outlook_discard_draft`.

## Attachments and people

- `outlook_attachments` lists or downloads attachments (`folder`+`id`, or `search`+`id` for search results). Outlook does not copy attachments to OneDrive by itself; `save_to_onedrive` does it on request.
- `outlook_people` finds email addresses, titles and phone numbers before drafting to someone new.

## Housekeeping

- `outlook_mark` with `delete` or `archive` only when the user asks. Throw away a mistaken draft with `outlook_discard_draft`.
- If something looks wrong, `browser_screenshot` shows the page.

## When a tool fails

- Retry the same outlook tool once (it reloads Outlook and clears pop-ups itself), then report the exact error.
- Don't switch to general browser tools (Playwright MCP, screenshots plus clicking) to do Outlook work: they share the same browser and make things worse. Use `outlook_search` only as a different way to find mail, not as a workaround for a broken folder.
- Never answer an Outlook prompt by guessing. Prompts about time zones, sending or deleting are for the user.
