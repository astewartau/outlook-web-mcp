#!/usr/bin/env node
// outlook-web-mcp: an MCP server that works in Outlook on the web by driving a signed-in Chromium
// over the Chrome DevTools Protocol, the same way a person would (click, type, read the page).
// No Microsoft Graph, no app registration: it only does what the signed-in user can do in the browser.
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { chromium } from 'playwright';

const CDP_URL = process.env.OUTLOOK_CDP_URL || 'http://127.0.0.1:9222';
const BASE = (process.env.OUTLOOK_URL || 'https://outlook.cloud.microsoft').replace(/\/$/, '');
const ACCOUNT = process.env.OUTLOOK_ACCOUNT || '';           // account to pick on the Microsoft sign-in page
const TIMEZONE = process.env.OUTLOOK_TIMEZONE || Intl.DateTimeFormat().resolvedOptions().timeZone;
const SEND_ENABLED = /^(1|true|yes)$/i.test(process.env.OUTLOOK_ENABLE_SEND || '');

let browser, queue = Promise.resolve();

// One browser action at a time: every tool shares the one Outlook tab.
function serial(fn) {
  const run = queue.then(fn, fn);
  queue = run.catch(() => {});
  return run;
}

async function page() {
  if (!browser || !browser.isConnected()) browser = await chromium.connectOverCDP(CDP_URL);
  const ctx = browser.contexts()[0];
  let p = ctx.pages().find(p => /outlook\.(cloud\.microsoft|office\.com|office365\.com|live\.com)/.test(p.url()));
  if (!p) { p = await ctx.newPage(); await p.goto(BASE + '/mail/'); }
  return p;
}

async function signedIn(p) {
  if (/login\.microsoftonline|login\.live|okta|adfs|\/saml/i.test(p.url())) return false;
  return (await p.locator('[role=tree]').count()) > 0 || /\/calendar/.test(p.url());
}

// Close the Reminders pop-up (without dismissing reminders): it sits over the page and swallows clicks.
// A leftover confirmation box blocks everything else too; cancel it (never confirm).
async function clearOverlays(p) {
  const cancel = p.locator('[role=dialog][aria-modal=true], [role=alertdialog]').getByRole('button', { name: /^(Cancel|No|Don't|Close)$/ }).first();
  if (await cancel.isVisible().catch(() => false)) { await cancel.click(); await p.waitForTimeout(800); }
  await p.evaluate(() => {
    for (const b of document.querySelectorAll('[data-portal-node=true] button[aria-label="Close"]')) {
      const pane = b.closest('[data-portal-node=true]');
      if (b.offsetParent && /Reminders|Dismiss all/.test(pane?.innerText || '')) b.click();
    }
  }).catch(() => {});
}

async function ensureMail(p) {
  if (!/\/mail/.test(p.url()) || !(await signedIn(p))) {
    await p.goto(BASE + '/mail/');
    await p.waitForLoadState('domcontentloaded');
  }
  await p.locator('[role=tree]').first().waitFor({ timeout: 30000 }).catch(() => {});
  if (!(await signedIn(p))) throw new Error('Not signed in to Outlook. Call outlook_login.');
  await clearOverlays(p);
}

const text = s => ({ content: [{ type: 'text', text: typeof s === 'string' ? s : JSON.stringify(s, null, 1) }] });
const tool = fn => async args => {
  try { return text(await serial(() => fn(args))); }
  catch (e) { return { ...text('Error: ' + e.message), isError: true }; }
};

async function expandFolders(p) {
  for (let i = 0; i < 6; i++) {
    const n = await p.evaluate(() => {
      const c = [...document.querySelectorAll('[role=tree] [role=treeitem][aria-expanded=false]')];
      c.forEach(e => (e.querySelector('button,i,[aria-label*="xpand"]') || e).click());
      return c.length;
    });
    if (!n) break;
    await p.waitForTimeout(700);
  }
}

// Folders under the account node (Favorites duplicates are skipped), with unread and item counts.
async function folders(p) {
  await expandFolders(p);
  return p.evaluate(() => {
    const items = [...document.querySelectorAll('[role=tree] [role=treeitem]')];
    const account = items.find(e => /@/.test(e.getAttribute('aria-label') || e.innerText));
    const out = [];
    for (const e of items) {
      if (account && items.indexOf(e) < items.indexOf(account)) continue;    // Favorites
      const raw = (e.getAttribute('aria-label') || e.innerText).replace(/[-]/g, '').replace(/\s+/g, ' ').trim();
      const m = raw.match(/^(.*?)( selected)?(?: (\d+) unread)?(?: \[(\d+)\] items?)?$/);
      if (!m || /@|^Go to Groups/.test(m[1])) continue;
      out.push({ name: m[1].trim(), unread: m[3] ? +m[3] : 0, items: m[4] ? +m[4] : 0, level: +(e.getAttribute('aria-level') || 2) });
    }
    return out;
  });
}
const fmtFolders = fs => fs.map(f => '  '.repeat(Math.max(0, f.level - 2)) + f.name + (f.unread ? ` (${f.unread} unread)` : '') + (f.items ? ` [${f.items} items]` : '')).join('\n');

async function openFolder(p, folder) {
  await ensureMail(p);
  await expandFolders(p);
  const esc = folder.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&');
  const t = p.getByRole('treeitem', { name: new RegExp('^[\\s\\uE000-\\uF8FF]*' + esc + '( selected)?( \\d+ unread)?( \\[\\d+\\] items?)?$', 'i') });
  if (!(await t.count())) throw new Error(`No folder named "${folder}". Use outlook_folders.`);
  await t.last().click();
  await p.locator('[role=listbox]').first().waitFor({ timeout: 15000 }).catch(() => {});
  await p.waitForTimeout(1500);
}

const FLAG = /^((?:Unread |Collapsed |Expanded |Flagged |Pinned |Replied |Forwarded |Has attachments |Important |Mentioned |Draft )+)/;

async function visibleItems(p) {
  const raw = await p.evaluate(() => [...document.querySelectorAll('[role=listbox] [role=option][data-convid]')]
    .map(e => ({ id: e.getAttribute('data-convid'), label: (e.getAttribute('aria-label') || '').replace(/\s+/g, ' ') })));
  return raw.map(({ id, label }) => {
    const flags = (label.match(FLAG) || ['', ''])[1].trim();
    return { id, unread: /^Unread\b/.test(label), flags, summary: label.slice(flags.length).trim().slice(0, 220) };
  });
}

async function scrollList(p) {
  const lb = p.locator('[role=listbox]').first();
  if (!(await lb.count())) return false;
  const bb = await lb.boundingBox();
  if (!bb) return false;
  const before = (await visibleItems(p)).map(i => i.id).join();
  await p.mouse.move(bb.x + bb.width / 2, bb.y + Math.min(bb.height, 600) / 2);
  await p.mouse.wheel(0, 600);
  await p.waitForTimeout(900);
  return (await visibleItems(p)).map(i => i.id).join() !== before;
}

async function scrollTop(p) {
  if (!(await p.locator('[role=listbox]').count())) return;
  const bb = await p.locator('[role=listbox]').first().boundingBox();
  if (!bb) return;
  await p.mouse.move(bb.x + bb.width / 2, bb.y + 100);
  for (let i = 0; i < 30; i++) await p.mouse.wheel(0, -3000);
  await p.waitForTimeout(800);
}

// The message list is virtualised: scroll through it to collect items, then scroll back.
// Pinned conversations sit at the top of a folder whatever their date, so they don't count
// towards `limit` (otherwise a handful of pins hides every new message).
async function listFolder(p, { unread_only = false, limit = 50, match = '' }) {
  const seen = new Map();
  const want = i => (!unread_only || i.unread) && (!match || i.summary.toLowerCase().includes(match.toLowerCase()));
  const counted = () => [...seen.values()].filter(i => !/Pinned/.test(i.flags)).length;
  for (let i = 0; i < 80 && counted() < limit; i++) {
    for (const it of await visibleItems(p)) if (!seen.has(it.id) && want(it)) seen.set(it.id, it);
    if (!(await scrollList(p))) break;
  }
  await scrollTop(p);
  const all = [...seen.values()];
  return [...all.filter(i => /Pinned/.test(i.flags)), ...all.filter(i => !/Pinned/.test(i.flags)).slice(0, limit)];
}
const fmtItem = i => {
  const f = i.flags.replace(/Unread ?|Collapsed ?|Expanded ?|Pinned ?/g, '').trim();
  return `${i.unread ? '● ' : ''}${f ? '[' + f + '] ' : ''}${i.summary}\n  id: ${i.id}`;
};
const fmtList = l => {
  if (!l.length) return 'Nothing.';
  const pinned = l.filter(i => /Pinned/.test(i.flags)), rest = l.filter(i => !/Pinned/.test(i.flags));
  return [pinned.length ? `Pinned (${pinned.length}):\n` + pinned.map(fmtItem).join('\n') : '',
          `Newest first (${rest.length}):\n` + (rest.map(fmtItem).join('\n') || 'Nothing.')].filter(Boolean).join('\n\n');
};

async function findItem(p, id) {
  for (let pass = 0; pass < 2; pass++) {
    for (let i = 0; i < 60; i++) {
      const it = p.locator(`[role=listbox] [role=option][data-convid="${id}"]`).first();
      if (await it.count()) return it;
      if (!(await scrollList(p))) break;
    }
    await scrollTop(p);
  }
  throw new Error('Conversation not found in this folder: ' + id);
}

async function readingPaneText(p, max) {
  await p.waitForTimeout(2000);
  // expand collapsed earlier messages so the whole thread is captured
  await p.evaluate(() => document.querySelectorAll('[aria-label="Reading Pane"] [aria-expanded=false][role=button]').forEach(b => b.click())).catch(() => {});
  await p.waitForTimeout(800);
  const t = await p.evaluate(() => (document.querySelector('[aria-label="Reading Pane"]') || document.querySelector('[role=main]'))?.innerText || '');
  return t.replace(/[͏‌­]+/g, '').replace(/[ \t ]+\n/g, '\n').replace(/\n{2,}/g, '\n')
    .replace(/^(Summarize this email|Reply|Reply all|Forward)\n/gm, '').slice(0, max);
}

async function openConversation(p, folder, id, mark_read) {
  await openFolder(p, folder);
  const it = await findItem(p, id);
  const wasUnread = /^Unread\b/.test((await it.getAttribute('aria-label')) || '');
  await it.click();
  if (wasUnread && mark_read) { await p.waitForTimeout(500); await p.keyboard.press('Control+q').catch(() => {}); }
  if (wasUnread && !mark_read) { await p.waitForTimeout(1500); await p.keyboard.press('Control+u').catch(() => {}); }
  return wasUnread;
}

async function fillBody(p, body) {
  await clearOverlays(p);
  const ed = p.locator('[contenteditable=true][aria-label*="Message body"], div[role=textbox][aria-label*="Message body"]').first();
  await ed.waitFor({ timeout: 15000 });
  await ed.click();
  await p.keyboard.press('Control+Home');
  for (const [i, line] of body.split('\n').entries()) {
    if (i) await p.keyboard.press('Enter');
    if (line) await p.keyboard.insertText(line);
  }
}

async function addRecipients(p, field, list) {
  if (!list.length) return;
  const box = p.locator(`[contenteditable=true][aria-label="${field}"], div[role=textbox][aria-label="${field}"], input[aria-label="${field}"]`).first();
  if (!(await box.isVisible().catch(() => false))) await p.getByRole('button', { name: new RegExp('^' + field + '$') }).first().click().catch(() => {});
  await box.click();
  for (const a of list) { await p.keyboard.insertText(a); await p.waitForTimeout(800); await p.keyboard.press(';'); await p.waitForTimeout(400); }
}

// Save the open compose form as a draft, then reload the mail view so the page is clean.
async function saveDraft(p) {
  await p.keyboard.press('Control+s');
  await p.waitForTimeout(2500);
  let ok = await p.getByText(/Draft saved/i).first().isVisible().catch(() => false);
  if (!ok) { await p.keyboard.press('Control+s'); await p.waitForTimeout(3000); ok = await p.getByText(/Draft saved/i).first().isVisible().catch(() => false); }
  if (!ok) throw new Error('Could not confirm the draft saved; it is still open in the browser. Use browser_screenshot.');
  await p.goto(BASE + '/mail/');
  await p.locator('[role=tree]').first().waitFor({ timeout: 30000 }).catch(() => {});
}

// Find a draft's id in Drafts by a piece of its text (subject or body).
async function draftIdFor(p, needle) {
  await openFolder(p, 'Drafts');
  const hit = (await listFolder(p, { limit: 50, match: needle }))[0];
  return hit?.id;
}

// Open a draft from Drafts in the compose form. New-mail drafts open straight into it;
// reply drafts open as a preview with an Edit (pencil) button.
async function openDraftForEdit(p, id) {
  await openFolder(p, 'Drafts');
  const it = await findItem(p, id);
  await it.click();
  await clearOverlays(p);
  const editor = p.locator('[contenteditable=true][aria-label*="Message body"]').first();
  const edit = p.getByRole('button', { name: /^(Edit|Continue editing)/ }).filter({ visible: true }).first();
  for (let i = 0; i < 20 && !(await editor.isVisible().catch(() => false)); i++) {
    if (await edit.isVisible().catch(() => false)) await edit.click();
    await p.waitForTimeout(700);
  }
  await editor.waitFor({ timeout: 10000 });
}

// What the open compose form holds: recipients, subject and body text.
async function composeSummary(p, max = 6000) {
  return p.evaluate(max => {
    const field = f => {
      const el = document.querySelector(`[contenteditable=true][aria-label="${f}"]`);
      let box = el; for (let i = 0; i < 3 && box && !box.innerText.trim(); i++) box = box.parentElement;
      const t = (box?.innerText || '').replace(/[\u200b\s]+/g, ' ').trim();
      return t ? `${f}: ${t}` : '';
    };
    const subject = document.querySelector('input[aria-label="Subject"]')?.value || '';
    const body = document.querySelector('[contenteditable=true][aria-label*="Message body"]')?.innerText.replace(/\n{3,}/g, '\n\n').trim() || '';
    return [field('To'), field('Cc'), field('Bcc'), 'Subject: ' + subject, '', body].filter((x, i) => x || i === 4).join('\n').slice(0, max);
  }, max);
}

async function leaveCompose(p) {
  await p.keyboard.press('Control+s').catch(() => {});
  await p.waitForTimeout(1500);
  await p.goto(BASE + '/mail/');
  await p.locator('[role=tree]').first().waitFor({ timeout: 30000 }).catch(() => {});
}

const server = new McpServer({ name: 'outlook-web', version: '0.1.0' }, {
  instructions: [
    'Outlook on the web through a signed-in browser. Folder names come from outlook_folders; conversation ids from outlook_list or outlook_search.',
    'Replies and new mail are saved as drafts. Send a draft only with outlook_send_draft, and only after the user has explicitly approved sending that specific draft.',
    'If a tool says not signed in, call outlook_login and pass any number it returns on to the user to approve in their authenticator app.',
  ].join(' '),
});

server.registerTool('outlook_status', { description: 'Whether the browser is signed in to Outlook, and the page it is on.', inputSchema: {} },
  tool(async () => { const p = await page(); return { signed_in: await signedIn(p), url: p.url().slice(0, 120), title: await p.title(), send_enabled: SEND_ENABLED }; }));

server.registerTool('outlook_login', {
  description: 'Sign in again after the session expires, using the password saved in the browser (Chrome autofill; the password is never read). Handles the Microsoft account picker and Okta or Microsoft number-matching MFA: returns the number for the user to approve in their authenticator app. Call again with wait_seconds to wait for approval.',
  inputSchema: { wait_seconds: z.number().int().min(0).max(180).default(0) },
}, tool(async ({ wait_seconds }) => {
  const p = await page();
  if (await signedIn(p)) return 'Already signed in.';
  if (!/login\.microsoftonline|login\.live|okta/i.test(p.url())) { await p.goto(BASE + '/mail/'); await p.waitForTimeout(5000); }
  const deadline = Date.now() + Math.max(wait_seconds, 25) * 1000;
  while (Date.now() < deadline) {
    if (await signedIn(p)) return 'Signed in.';
    if (ACCOUNT) {
      const acct = p.getByText(ACCOUNT).first();
      if (/login\.microsoftonline/.test(p.url()) && await acct.isVisible().catch(() => false)) { await acct.click(); await p.waitForTimeout(3000); continue; }
    }
    const body = await p.evaluate(() => document.body.innerText).catch(() => '');
    const num = body.match(/(?:tap|select|enter|number)[^\n]{0,80}?\b(\d{2,3})\b/i) || body.match(/\n\s*(\d{2,3})\s*\n/);
    if (num && /okta|login\.microsoftonline/i.test(p.url()) && /verify|approve|authenticator|okta/i.test(body)) {
      if (!wait_seconds) return `Approval requested: the user should tap or enter ${num[1]} in their authenticator app. Then call outlook_login with wait_seconds=60.`;
      await p.waitForTimeout(3000); continue;
    }
    const push = p.getByRole('link', { name: /push notification|okta verify/i }).or(p.getByRole('button', { name: /send push|push notification/i })).first();
    if (await push.isVisible().catch(() => false)) { await push.click(); await p.waitForTimeout(3000); continue; }
    const pw = p.locator('input[type=password]').first();
    if (await pw.isVisible().catch(() => false)) {
      await pw.click(); await p.waitForTimeout(1500);   // give Chrome a moment to autofill
      if (!(await pw.inputValue())) return 'The password field is empty (no saved password was autofilled). The user needs to sign in in the browser themselves.';
      await p.keyboard.press('Enter'); await p.waitForTimeout(4000); continue;
    }
    const user = p.locator('input[name=identifier], input[name=username], input[name=loginfmt], input[type=email]').first();
    if (await user.isVisible().catch(() => false)) {
      if (!(await user.inputValue())) { if (!ACCOUNT) return 'The sign-in page asks for a username. Set OUTLOOK_ACCOUNT or sign in in the browser.'; await user.fill(ACCOUNT); }
      await p.keyboard.press('Enter'); await p.waitForTimeout(4000); continue;
    }
    const next = p.getByRole('button', { name: /^(Yes|Next|Verify|Sign in|Continue)$/ }).first();
    if (await next.isVisible().catch(() => false)) { await next.click(); await p.waitForTimeout(3000); continue; }
    await p.waitForTimeout(2000);
  }
  return `Stopped at ${p.url().slice(0, 100)} (${await p.title()}). Use browser_screenshot to see it.`;
}));

server.registerTool('outlook_folders', {
  description: 'Mail folders (indented by nesting) with unread and item counts.',
  inputSchema: { unread_only: z.boolean().default(false) },
}, tool(async ({ unread_only }) => {
  const p = await page(); await ensureMail(p);
  const f = (await folders(p)).filter(x => !unread_only || x.unread);
  return f.length ? fmtFolders(f) : 'No unread mail.';
}));

server.registerTool('outlook_list', {
  description: 'List conversations in a folder: pinned ones first (in their own section, not counted in limit), then the newest `limit` others. Each has an id, unread (●), flags and a one-line summary (sender, subject, time, preview). `match` keeps only items whose summary contains that text (sender, subject or preview).',
  inputSchema: { folder: z.string().default('Inbox'), unread_only: z.boolean().default(false), limit: z.number().int().min(1).max(200).default(40), match: z.string().default(''), include_pinned: z.boolean().default(true) },
}, tool(async ({ folder, unread_only, limit, match, include_pinned }) => {
  const p = await page(); await openFolder(p, folder);
  const l = await listFolder(p, { unread_only, limit, match });
  return fmtList(include_pinned ? l : l.filter(i => !/Pinned/.test(i.flags)));
}));

server.registerTool('outlook_read', {
  description: 'Read a conversation (every message in the thread) by id. Marks it read unless mark_read is false.',
  inputSchema: { folder: z.string().default('Inbox'), id: z.string(), mark_read: z.boolean().default(true), max_chars: z.number().int().default(12000) },
}, tool(async ({ folder, id, mark_read, max_chars }) => { const p = await page(); await openConversation(p, folder, id, mark_read); return readingPaneText(p, max_chars); }));

server.registerTool('outlook_read_unread', {
  description: 'Read every unread conversation in one folder, or in every folder with unread mail if folder is omitted, marking each read. Returns their text.',
  inputSchema: { folder: z.string().optional(), max_chars_each: z.number().int().default(5000) },
}, tool(async ({ folder, max_chars_each }) => {
  const p = await page(); await ensureMail(p);
  const targets = folder ? [folder] : [...new Set((await folders(p)).filter(f => f.unread).map(f => f.name))];
  const out = [];
  for (const f of targets) {
    await openFolder(p, f);
    for (const { id } of await listFolder(p, { unread_only: true, limit: 200 })) {
      try {
        const it = await findItem(p, id); await it.click();
        await p.waitForTimeout(500); await p.keyboard.press('Control+q').catch(() => {});
        out.push(`### [${f}] ${id}\n` + await readingPaneText(p, max_chars_each));
      } catch (e) { out.push(`### [${f}] ${id}\n(could not open: ${e.message})`); }
    }
  }
  return out.length ? out.join('\n\n=========\n\n') : 'No unread mail.';
}));

server.registerTool('outlook_mark', {
  description: 'Mark a conversation read or unread, flag or unflag, pin or unpin, archive it, or delete it (moves it to Deleted Items; only when the user asks). To throw away a draft use outlook_discard_draft.',
  inputSchema: { folder: z.string().default('Inbox'), id: z.string(), action: z.enum(['read', 'unread', 'flag', 'unflag', 'pin', 'unpin', 'archive', 'delete']) },
}, tool(async ({ folder, id, action }) => {
  const p = await page(); await openFolder(p, folder);
  const it = await findItem(p, id);
  const menu = async name => {
    await it.click({ button: 'right' }); await p.waitForTimeout(800);
    await p.getByRole('menuitem').filter({ hasText: new RegExp('^[\\s\\uE000-\\uF8FF]*' + name + '\\s*$', 'i') }).first().click();
    await p.waitForTimeout(1500);
    const dlg = p.locator('[role=dialog][aria-modal=true], [role=alertdialog]').filter({ visible: true });
    if (await dlg.count()) { const msg = (await dlg.first().innerText()).replace(/\s+/g, ' ').slice(0, 200); await clearOverlays(p); throw new Error(`Outlook asked for confirmation and it was cancelled: "${msg}"`); }
  };
  const label = async () => (await it.getAttribute('aria-label').catch(() => '')) || '';
  if (action === 'delete' || action === 'archive') {
    await menu(action);
    return (await p.locator(`[role=option][data-convid="${id}"]`).count()) ? `Tried to ${action}, but it is still listed.` : `Done: ${action}.`;
  }
  if (action === 'pin' || action === 'unpin') { await menu(action); return `Done: ${action}.`; }
  if (action === 'flag' || action === 'unflag') {
    if ((action === 'flag') === /Flagged/.test(await label())) return `Already ${action}ged.`;
    await it.click(); await p.waitForTimeout(1000); await p.keyboard.press('Insert'); await p.waitForTimeout(1000);
    return `Done: ${action}.`;
  }
  await it.click(); await p.waitForTimeout(1200);
  await p.keyboard.press(action === 'read' ? 'Control+q' : 'Control+u'); await p.waitForTimeout(1000);
  return `Done: ${action}.`;
}));

server.registerTool('outlook_search', {
  description: 'Search all folders (Sent Items included) with Outlook search syntax. Combine terms: from:name, to:name, cc:name, subject:words, hasattachments:yes, attachment:filename, received:2026-11-18, received>=2026-11-01, received:this week, plus plain words. Use the user\'s own address rather than "me" to find mail they sent. Returns conversations with ids; open one with outlook_read_search_result.',
  inputSchema: { query: z.string(), limit: z.number().int().min(1).max(100).default(25) },
}, tool(async ({ query, limit }) => {
  const p = await page(); await ensureMail(p);
  const box = p.getByRole('combobox', { name: /search/i }).or(p.locator('input[aria-label*="Search"]')).first();
  await box.click(); await box.fill(query); await p.keyboard.press('Enter');
  await p.waitForTimeout(4000);
  return fmtList(await listFolder(p, { limit }));
}));

server.registerTool('outlook_read_search_result', {
  description: 'Open a conversation from the most recent outlook_search by id. Marks it read.',
  inputSchema: { id: z.string(), max_chars: z.number().int().default(12000) },
}, tool(async ({ id, max_chars }) => { const p = await page(); const it = await findItem(p, id); await it.click(); return readingPaneText(p, max_chars); }));

server.registerTool('outlook_draft_reply', {
  description: 'Write a reply, reply-all or forward for a conversation and save it as a draft (does not send). The text goes above the quoted thread. Forward needs `to`. Returns the draft id for outlook_send_draft.',
  inputSchema: {
    folder: z.string().default('Inbox'), id: z.string(), body: z.string(),
    mode: z.enum(['reply', 'reply_all', 'forward']).default('reply'),
    to: z.array(z.string()).default([]), cc: z.array(z.string()).default([]),
  },
}, tool(async ({ folder, id, body, mode, to, cc }) => {
  if (mode === 'forward' && !to.length) throw new Error('A forward needs at least one address in `to`.');
  const p = await page(); await openConversation(p, folder, id, true);
  await p.waitForTimeout(1500); await clearOverlays(p);
  const name = { reply: /^Reply$/, reply_all: /^Reply all$/, forward: /^Forward$/ }[mode];
  await p.getByRole('button', { name }).filter({ visible: true }).last().click();
  await p.waitForTimeout(2000);
  await addRecipients(p, 'To', to);
  await addRecipients(p, 'Cc', cc);
  await fillBody(p, body);
  await saveDraft(p);
  const draft = await draftIdFor(p, body.split('\n').find(Boolean).slice(0, 40));
  return `Draft saved (not sent). Draft id: ${draft || '(not found in Drafts yet; list Drafts)'}`;
}));

server.registerTool('outlook_draft_new', {
  description: 'Write a new email and save it as a draft (does not send). to/cc/bcc are lists of email addresses. Returns the draft id for outlook_send_draft.',
  inputSchema: { to: z.array(z.string()).default([]), cc: z.array(z.string()).default([]), bcc: z.array(z.string()).default([]), subject: z.string(), body: z.string() },
}, tool(async ({ to, cc, bcc, subject, body }) => {
  const p = await page(); await ensureMail(p);
  await p.getByRole('button', { name: /^New mail$/ }).first().click();
  await p.waitForTimeout(2500);
  await addRecipients(p, 'To', to);
  await addRecipients(p, 'Cc', cc);
  await addRecipients(p, 'Bcc', bcc);
  const subj = p.locator('input[aria-label="Subject"], input[placeholder="Add a subject"]').first();
  await subj.click(); await subj.fill(subject);
  await fillBody(p, body);
  await saveDraft(p);
  const draft = await draftIdFor(p, subject);
  return `Draft saved (not sent). Draft id: ${draft || '(not found in Drafts yet; list Drafts)'}`;
}));

server.registerTool('outlook_read_draft', {
  description: 'Show a draft from Drafts (recipients, subject and text) so the user can check it before sending.',
  inputSchema: { id: z.string(), max_chars: z.number().int().default(6000) },
}, tool(async ({ id, max_chars }) => {
  const p = await page(); await openDraftForEdit(p, id);
  const t = await composeSummary(p, max_chars);
  await leaveCompose(p);
  return t;
}));

server.registerTool('outlook_send_draft', {
  description: 'Send a draft from Drafts. Only call this after the user has explicitly approved sending this specific draft (show it with outlook_read_draft first). Disabled unless the server runs with OUTLOOK_ENABLE_SEND=1.',
  inputSchema: { id: z.string(), expected_subject: z.string().describe('Part of the draft subject, as a check that the right draft is sent.') },
}, tool(async ({ id, expected_subject }) => {
  if (!SEND_ENABLED) throw new Error('Sending is disabled. Start the server with OUTLOOK_ENABLE_SEND=1 to allow it.');
  const p = await page(); await openFolder(p, 'Drafts');
  const it = await findItem(p, id);
  const label = (await it.getAttribute('aria-label')) || '';
  if (!label.toLowerCase().includes(expected_subject.toLowerCase())) throw new Error(`Draft ${id} does not match "${expected_subject}": ${label.slice(0, 150)}`);
  await openDraftForEdit(p, id);
  const summary = await composeSummary(p, 400);
  const who = summary.split('\n').filter(l => /^(To|Cc|Bcc):/.test(l)).join('; ');
  if (!/^To: \S/m.test(summary)) { await leaveCompose(p); throw new Error('The draft has no recipients; not sent.'); }
  await p.getByRole('button', { name: /^Send$/ }).filter({ visible: true }).first().click();
  await p.waitForTimeout(4000);
  const dlg = p.locator('[role=dialog][aria-modal=true], [role=alertdialog]').filter({ visible: true });
  if (await dlg.count()) { const msg = (await dlg.first().innerText()).replace(/\s+/g, ' ').slice(0, 300); await clearOverlays(p); throw new Error(`Outlook asked something before sending, and it was cancelled (not sent): "${msg}"`); }
  await openFolder(p, 'Drafts');
  const still = await p.locator(`[role=option][data-convid="${id}"]`).filter({ hasText: expected_subject }).count();
  return still ? 'The draft is still in Drafts: sending may have failed. Check with browser_screenshot.' : `Sent. ${who}`;
}));

server.registerTool('outlook_discard_draft', {
  description: 'Throw away a draft from Drafts (for example one made by mistake).',
  inputSchema: { id: z.string(), expected_subject: z.string().describe('Part of the draft subject or text, as a check.') },
}, tool(async ({ id, expected_subject }) => {
  const p = await page(); await openFolder(p, 'Drafts');
  const it = await findItem(p, id);
  if (!((await it.getAttribute('aria-label')) || '').toLowerCase().includes(expected_subject.toLowerCase())) throw new Error('That draft does not match expected_subject.');
  await it.click(); await p.waitForTimeout(2500); await clearOverlays(p);
  const bin = p.locator('[aria-label="Reading Pane"], [role=main]').getByRole('button', { name: /^(Discard|Delete)/ }).filter({ visible: true }).first();
  await bin.click(); await p.waitForTimeout(1500);
  const ok = p.locator('[role=dialog], [role=alertdialog]').getByRole('button', { name: /^(OK|Discard|Delete)$/ }).filter({ visible: true });
  if (await ok.count()) await ok.first().click();
  await p.waitForTimeout(2500);
  await openFolder(p, 'Drafts');
  return (await p.locator(`[role=option][data-convid="${id}"]`).filter({ hasText: expected_subject }).count()) ? 'Still in Drafts.' : 'Discarded.';
}));

server.registerTool('outlook_calendar', {
  description: 'Calendar events from start_date (YYYY-MM-DD, default today) for `days` days, grouped by day: title, time, location, organiser and response as Outlook shows them.',
  inputSchema: { start_date: z.string().optional(), days: z.number().int().min(1).max(31).default(1) },
}, tool(async ({ start_date, days }) => {
  const today = new Intl.DateTimeFormat('en-CA', { timeZone: TIMEZONE }).format(new Date());
  const start = new Date((start_date || today) + 'T00:00:00');
  const fmt = d => d.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' });
  const want = [];
  for (let i = 0; i < days; i++) { const d = new Date(start); d.setDate(d.getDate() + i); want.push(fmt(d)); }
  const seen = new Map();
  const cp = await (await page()).context().newPage();
  try {
    for (let i = 0; i < days; i += 7) {
      const d = new Date(start); d.setDate(d.getDate() + i);
      await cp.goto(`${BASE}/calendar/view/week/${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`);
      await cp.waitForTimeout(7000);
      for (const l of await cp.evaluate(() => [...document.querySelectorAll('[role=button][aria-label]')].map(e => e.getAttribute('aria-label')))) {
        const day = want.find(w => l.includes(w));
        if (day && /\d{1,2}:\d\d|all day/i.test(l)) seen.set(l, day);
      }
    }
  } finally { await cp.close(); }
  const byDay = {};
  for (const [l, d] of seen) (byDay[d] ||= []).push(l.replace(', ' + d, ''));
  return Object.keys(byDay).length ? byDay : 'No events found.';
}));

server.registerTool('browser_screenshot', {
  description: 'Screenshot the Outlook tab, to see what state it is in.',
  inputSchema: {},
}, async () => serial(async () => {
  const p = await page();
  return { content: [{ type: 'image', data: (await p.screenshot()).toString('base64'), mimeType: 'image/png' }] };
}));

await server.connect(new StdioServerTransport());
