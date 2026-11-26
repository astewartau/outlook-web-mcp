# outlook-web-mcp

An [MCP](https://modelcontextprotocol.io) server that lets an AI assistant (Claude Code, Claude Desktop or any MCP client) work in **Outlook on the web** by driving a browser that you have signed in to: list and read mail, triage folders, search, draft replies, send drafts you approve, and read your calendar.

It does not use Microsoft Graph and needs no app registration or admin consent. It clicks, types and reads the page the way you would, so it can do exactly what you can do in Outlook on the web and nothing more. That makes it useful where an organisation blocks third-party mail apps but you still want an assistant to help with your own mailbox.

> **Check your organisation's rules first.** Driving your own signed-in browser is not the same as an approved integration. If your IT policy forbids automating your mailbox, don't use this.

## How it works

```
MCP client ──stdio──▶ src/server.mjs ──CDP (localhost)──▶ Chromium on a virtual display
                                                              ▲
                                    you, via noVNC ───────────┘  (sign in, approve MFA, watch)
```

- `browser/start-browser.sh` runs Chromium on a virtual display (Xvfb) with a persistent profile, so sign-ins survive restarts. It exposes:
  - the Chrome DevTools Protocol on `127.0.0.1` only, for the MCP server;
  - a password-protected [noVNC](https://novnc.com) viewer, so you can sign in, approve MFA and watch or take over from any browser (phone included). Put it on a private network such as Tailscale, not the open internet.
- `src/server.mjs` connects over CDP and drives the Outlook tab with [Playwright](https://playwright.dev). Tools run one at a time on that one tab.

It also works with an ordinary desktop Chrome started with `--remote-debugging-port=9222`.

## Tools

| Tool | What it does |
|---|---|
| `outlook_status` | Signed in or not, and the current page |
| `outlook_login` | Signs in again with the browser's saved password (autofill; never read) and returns the MFA number to approve (Okta Verify or Microsoft Authenticator number matching) |
| `outlook_folders` | Folder tree with unread and item counts |
| `outlook_list` | Conversations in a folder: pinned ones separately, then the newest `limit`; filter by unread or by text |
| `outlook_read` | A whole conversation by id, with each link's address after its text (marks it read unless asked not to) |
| `outlook_read_unread` | Every unread conversation in one folder or in all folders, marking each read |
| `outlook_search` / `outlook_read_search_result` | Outlook search (`from:`, `subject:` …) and opening a result (pass the same `query` to search again first) |
| `outlook_mark` | Read, unread, flag, unflag, pin, unpin, archive, delete |
| `outlook_draft_reply` | Reply, reply-all or forward, saved as a draft, with optional attachments. Stops if the conversation already has a draft, unless told to replace that draft's text |
| `outlook_draft_new` | New email saved as a draft, with optional attachments |
| `outlook_edit_draft` | Change a draft: replace its text (signature and quoted thread stay), recipients or subject, add attachments |
| `outlook_read_draft` | Recipients, subject, attachment names and text of a draft, to check before sending |
| `outlook_send_draft` | Sends a draft. **Off unless `OUTLOOK_ENABLE_SEND=1`**; checks the subject and that it has recipients |
| `outlook_discard_draft` | Throws a draft away |
| `outlook_attachments` | List a conversation's attachments, download them to a local folder, or Save to OneDrive |
| `outlook_people` | Look people up in contacts and the organisation directory: name, title, email, phone |
| `outlook_calendar` | Events for a range of days |
| `outlook_event` | One event in full: organiser, attendees and responses, location, Teams link, description |
| `outlook_event_respond` | Accept, tentative or decline an invitation, optionally with a note. **Needs `OUTLOOK_ENABLE_SEND=1`** (it emails the organiser) |
| `outlook_event_create` | New event: title, date, times or all day, location, Teams meeting, description; attendees get invitations (**needs `OUTLOOK_ENABLE_SEND=1`**) |
| `outlook_event_delete` | Delete an event; for a meeting you organised, sends the cancellation (**needs `OUTLOOK_ENABLE_SEND=1`**) |
| `browser_screenshot` | What the Outlook tab looks like right now |

Drafting and sending are separate on purpose: the assistant writes a draft, you read it (in Outlook or through `outlook_read_draft`), and only then is it sent. Everything else that emails other people (invitation responses, invitations, cancellations) is behind the same `OUTLOOK_ENABLE_SEND` switch.

## Setup (Linux)

Tested on Ubuntu 24.04 with Node 22.

```bash
sudo apt-get install -y xvfb x11vnc novnc websockify
git clone https://github.com/astewartau/outlook-web-mcp && cd outlook-web-mcp
npm install
npx playwright install chromium

# a password for the viewer (x11vnc uses the first 8 characters)
mkdir -p ~/.config/outlook-web-mcp && x11vnc -storepasswd ~/.config/outlook-web-mcp/vncpass

# start it (or install browser/outlook-browser.service as a systemd user unit)
VIEW_ADDR=<your Tailscale IP> browser/start-browser.sh
```

Open `http://<VIEW_ADDR>:6080/vnc.html?autoconnect=1&resize=scale`, sign in to Outlook, and let Chromium save your password if you want `outlook_login` to handle expired sessions.

**Ubuntu 23.10+:** if Chromium exits with `No usable sandbox!`, install `browser/apparmor-chromium` (instructions inside). It lets that one binary use its sandbox; don't run with `--no-sandbox` for a browser that holds your mail.

`start-browser.sh` settings (environment): `VIEW_ADDR`, `VIEW_PORT` (6080), `CDP_PORT` (9222), `DISPLAY_NUM` (77), `PROFILE_DIR`, `VNC_PASSFILE`, `CHROME`, `START_URL`, `SIZE`.

## Adding it to Claude Code

```bash
claude mcp add -s user outlook-web \
  -e OUTLOOK_ACCOUNT=you@example.edu -e OUTLOOK_TIMEZONE=Australia/Brisbane -e OUTLOOK_ENABLE_SEND=1 \
  -- node /path/to/outlook-web-mcp/src/server.mjs
```

| Variable | Default | |
|---|---|---|
| `OUTLOOK_CDP_URL` | `http://127.0.0.1:9222` | Browser to drive |
| `OUTLOOK_URL` | `https://outlook.cloud.microsoft` | Outlook address (`https://outlook.office.com` also works) |
| `OUTLOOK_ACCOUNT` | | Account to pick on the Microsoft sign-in page |
| `OUTLOOK_TIMEZONE` | system zone | Defines "today" for `outlook_calendar` |
| `OUTLOOK_ENABLE_SEND` | off | Allow tools that email other people: sending drafts, invitation responses, invitations, cancellations |
| `OUTLOOK_DOWNLOAD_DIR` | `~/Downloads/outlook-attachments` | Where attachments are downloaded |

`skills/outlook-email/SKILL.md` is a Claude Code skill with working rules (triage, drafting, when sending is allowed). Copy it to `~/.claude/skills/outlook-email/`.

Test a tool without an MCP client: `node scripts/call.mjs outlook_folders '{}'`.

## Limitations

- It depends on Outlook's page structure (ARIA roles and labels). When Microsoft changes the page, selectors may need updating. It was built against the new Outlook on the web in English (late 2026).
- One tab and one action at a time; reading a long folder takes seconds, not milliseconds.
- Events are same-day (no multi-day or recurring events yet). The calendar ignores dates in its URL, so the tools page through the mini calendar.
- `outlook_event_respond` with a note goes through Outlook's "RSVP with note" dialog.
- If the site ever publishes [WebMCP](https://github.com/webmachinelearning/webmcp) tools, those would be the better route.

## Licence

MIT
