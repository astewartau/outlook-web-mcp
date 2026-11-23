#!/usr/bin/env node
// outlook-web-mcp: an MCP server that works in Outlook on the web by driving a signed-in Chromium
// over the Chrome DevTools Protocol, the same way a person would (click, type, read the page).
// No Microsoft Graph, no app registration: it only does what the signed-in user can do in the browser.
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { chromium } from 'playwright';
import fs from 'fs';
import path from 'path';

const CDP_URL = process.env.OUTLOOK_CDP_URL || 'http://127.0.0.1:9222';
const BASE = (process.env.OUTLOOK_URL || 'https://outlook.cloud.microsoft').replace(/\/$/, '');
const ACCOUNT = process.env.OUTLOOK_ACCOUNT || '';           // account to pick on the Microsoft sign-in page
const TIMEZONE = process.env.OUTLOOK_TIMEZONE || Intl.DateTimeFormat().resolvedOptions().timeZone;
const SEND_ENABLED = /^(1|true|yes)$/i.test(process.env.OUTLOOK_ENABLE_SEND || '');
const DOWNLOAD_DIR = process.env.OUTLOOK_DOWNLOAD_DIR || `${process.env.HOME}/Downloads/outlook-attachments`;

const TAB_NAME = 'outlook-web-mcp';
const LOCK = process.env.OUTLOOK_LOCK_FILE || `/tmp/outlook-web-mcp-${process.getuid?.() ?? 'user'}.lock`;
let browser, tab, queue = Promise.resolve();

// Every client (each Claude session starts its own server) shares one browser and one tab,
// so calls are serialised within this process (queue) and across processes (lock file).
function serial(fn) {
  const run = queue.then(() => withLock(fn), () => withLock(fn));
  queue = run.catch(() => {});
  return run;
}

async function withLock(fn) {
  const deadline = Date.now() + 10 * 60 * 1000;
  for (;;) {
    try { fs.writeFileSync(LOCK, String(process.pid), { flag: 'wx' }); break; }
    catch (e) {
      if (e.code !== 'EEXIST') throw e;
      const holder = +fs.readFileSync(LOCK, 'utf8') || 0;
      let alive = false; try { process.kill(holder, 0); alive = true; } catch {}
      if (!alive) { try { fs.unlinkSync(LOCK); } catch {} continue; }       // stale lock
      if (Date.now() > deadline) throw new Error('Another session has been using Outlook for 10 minutes; try again later.');
      await new Promise(r => setTimeout(r, 500));
    }
  }
  try { return await fn(); } finally { try { if (fs.readFileSync(LOCK, 'utf8') === String(process.pid)) fs.unlinkSync(LOCK); } catch {} }
}

// The server's own tab, marked with window.name so it never adopts a tab someone else opened.
async function page() {
  if (!browser || !browser.isConnected()) { browser = await chromium.connectOverCDP(CDP_URL); tab = null; }
  if (tab && !tab.isClosed()) return tab;
  const ctx = browser.contexts()[0];
  for (const p of ctx.pages()) {
    // window.name is cleared by the cross-site trip through the sign-in pages; sessionStorage survives it.
    if (await p.evaluate(n => window.name === n || sessionStorage.getItem(n) === '1', TAB_NAME).catch(() => false)) { tab = p; break; }
  }
  if (!tab) {
    tab = await ctx.newPage();
    await tab.goto(BASE + '/mail/');
  }
  await tab.evaluate(n => { window.name = n; sessionStorage.setItem(n, '1'); }, TAB_NAME).catch(() => {});
  // Outlook compares the browser's time zone with the mailbox's: report the user's zone, so it
  // never offers to move the calendar to the server's zone.
  const cdp = await ctx.newCDPSession(tab);
  await cdp.send('Emulation.setTimezoneOverride', { timezoneId: TIMEZONE }).catch(() => {});
  await tab.bringToFront().catch(() => {});
  return tab;
}

// When the session expires Outlook keeps the old page (folder tree and all) and only shows a
// "You need to sign in / Your session has expired" bar; nothing on the page works after that.
async function sessionExpired(p) {
  return p.getByText(/^(You need to sign in\.?|Your session has expired\b.*)$/).first().isVisible().catch(() => false);
}

async function signedIn(p) {
  if (/login\.microsoftonline|login\.live|okta|adfs|\/saml/i.test(p.url())) return false;
  if (await sessionExpired(p)) return false;
  return (await p.locator('[role=tree]').count()) > 0 || /\/calendar/.test(p.url());
}

async function ensureCalendar(cp) {
  if (/login\.microsoftonline|login\.live|okta|adfs|\/saml/i.test(cp.url()) || await sessionExpired(cp))
    throw new Error('Not signed in to Outlook (session expired). Call outlook_login.');
}

// Close the Reminders pop-up (without dismissing reminders): it sits over the page and swallows clicks.
// A leftover confirmation box blocks everything else too; cancel it (never confirm).
async function clearOverlays(p) {
  const cancel = p.locator('[role=dialog][aria-modal=true], [role=alertdialog]').getByRole('button', { name: /^(Cancel|No|Don't|Close)$/ }).first();
  if (await cancel.isVisible().catch(() => false)) { await cancel.click(); await p.waitForTimeout(800); }
  // Known prompts, answered with the option that changes nothing ("You're on the move" time zone
  // prompt: never "Yes", which would move the mailbox to the browser's zone).
  for (const name of ['Not now', 'Skip', 'Maybe later', 'Got it']) {
    const b = p.locator('[role=dialog], [role=alertdialog]').getByRole('button', { name, exact: true }).first();
    if (await b.isVisible().catch(() => false)) { await b.click(); await p.waitForTimeout(800); }
  }
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
    await p.evaluate(n => { window.name = n; sessionStorage.setItem(n, '1'); }, TAB_NAME).catch(() => {});
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
  if (!(await t.count())) {
    // the page may be mid-way through something else: reload the mail view and look again
    await p.goto(BASE + '/mail/'); await p.locator('[role=tree]').first().waitFor({ timeout: 30000 }).catch(() => {});
    await clearOverlays(p); await expandFolders(p);
    if (!(await t.count())) throw new Error(`No folder named "${folder}". Folders are:\n` + fmtFolders(await folders(p)));
  }
  await t.last().scrollIntoViewIfNeeded().catch(() => {});
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
// Drafts in one conversation share its id, so in Drafts (by_item) each draft is kept, not just the first.
async function listFolder(p, { unread_only = false, limit = 50, match = '', by_item = false }) {
  const seen = new Map();
  const want = i => (!unread_only || i.unread) && (!match || i.summary.toLowerCase().includes(match.toLowerCase()));
  const counted = () => [...seen.values()].filter(i => !/Pinned/.test(i.flags)).length;
  for (let i = 0; i < 80 && counted() < limit; i++) {
    for (const it of await visibleItems(p)) { const k = by_item ? it.id + '|' + it.summary : it.id; if (!seen.has(k) && want(it)) seen.set(k, it); }
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

// `needle` (part of the item's subject or text) picks one item when several share the id (drafts).
async function findItem(p, id, needle = '') {
  const sel = `[role=listbox] [role=option][data-convid="${id}"]` + (needle ? `[aria-label*="${needle.replace(/["\\]/g, '\\$&')}" i]` : '');
  for (let pass = 0; pass < 2; pass++) {
    for (let i = 0; i < 60; i++) {
      const it = p.locator(sel).first();
      if (await it.count()) return it;
      if (!(await scrollList(p))) break;
    }
    await scrollTop(p);
  }
  throw new Error(needle ? `Nothing with id ${id} and text "${needle}" in this folder.` : 'Conversation not found in this folder: ' + id);
}

async function readingPaneText(p, max) {
  await p.waitForTimeout(2000);
  // expand collapsed earlier messages so the whole thread is captured
  await p.evaluate(() => document.querySelectorAll('[aria-label="Reading Pane"] [aria-expanded=false][role=button]').forEach(b => b.click())).catch(() => {});
  await p.waitForTimeout(800);
  const t = await p.evaluate(() => {
    const pane = document.querySelector('[aria-label="Reading Pane"]') || document.querySelector('[role=main]');
    if (!pane) return '';
    // Show each link's address after its text ("Verify <https://...>"), unwrapping Safe Links.
    // Skip mailto/tel, links whose text already is the address, and all but the last of a run of
    // pieces of one link (a signature can split a word into one link per letter).
    const added = [];
    const as = [...pane.querySelectorAll('[aria-label="Message body"] a[href], [id^=UniqueMessageBody] a[href]')];
    as.forEach((a, i) => {
      let url = a.getAttribute('href') || '';
      if (!/^https?:/i.test(url) || as[i + 1]?.getAttribute('href') === url) return;
      try { const u = new URL(url); if (/safelinks\.protection\.outlook\.com$/i.test(u.hostname) && u.searchParams.get('url')) url = u.searchParams.get('url'); } catch {}
      const txt = a.innerText.replace(/\s+/g, ' ').trim();
      const bare = s => s.replace(/^https?:\/\/(www\.)?/i, '').replace(/\/$/, '').toLowerCase();
      if (!txt || bare(txt) === bare(url) || bare(url).startsWith(bare(txt)) && /\.\w/.test(txt)) return;
      const s = document.createElement('span'); s.setAttribute('data-mcp-link', ''); s.textContent = ` <${url}>`;
      a.after(s); added.push(s);
    });
    const out = pane.innerText;
    added.forEach(s => s.remove());
    return out;
  });
  return t.replace(/[͏‌­]+/g, '').replace(/[ \t ]+\n/g, '\n').replace(/\n{2,}/g, '\n')
    .replace(/^(Summarize this email|Reply|Reply all|Forward)\n/gm, '').slice(0, max);
}

async function openConversation(p, folder, id, mark_read) {
  await openFolder(p, folder);
  const it = await findItem(p, id);
  const label = (await it.getAttribute('aria-label')) || '';
  const wasUnread = /^Unread\b/.test(label);
  await it.click();
  if (wasUnread && mark_read) { await p.waitForTimeout(500); await p.keyboard.press('Control+q').catch(() => {}); }
  if (wasUnread && !mark_read) { await p.waitForTimeout(1500); await p.keyboard.press('Control+u').catch(() => {}); }
  return label;      // the list shows "[Draft]" in it while the conversation holds an unsent draft
}

// Click Reply, Reply all or Forward for an open conversation. The ribbon's buttons act on the
// newest message but are disabled while the conversation holds a draft (even a discarded one,
// still shown from Deleted Items); then use the reading pane's own (menu item) buttons.
async function clickReply(p, mode) {
  const label = { reply: 'Reply', reply_all: 'Reply all', forward: 'Forward' }[mode];
  const enabled = `[aria-label="${label}"]:not([aria-disabled="true"]):not([disabled])`;
  const ribbon = p.locator(`button${enabled}:not([aria-label="Reading Pane"] *)`).filter({ visible: true });
  const pane = p.locator('[aria-label="Reading Pane"]').locator(`button${enabled}, [role=button]${enabled}, [role=menuitem]${enabled}`).filter({ visible: true });
  for (let i = 0; i < 20; i++) {
    for (const b of [ribbon.first(), pane.first()]) if (await b.isVisible().catch(() => false)) { await b.click(); return; }
    await p.waitForTimeout(500);
  }
  throw new Error(`No enabled "${label}" button for this conversation. Use browser_screenshot.`);
}

// Type `body` at the top of the message. With replace, first remove the text already written there
// (everything above the signature or the quoted thread, which stay).
async function fillBody(p, body, { replace = false } = {}) {
  await clearOverlays(p);
  const ed = p.locator('[contenteditable=true][aria-label*="Message body"], div[role=textbox][aria-label*="Message body"]').first();
  await ed.waitFor({ timeout: 15000 });
  // Outlook adds the signature (and quoted text) a moment after the form opens, re-rendering the
  // body; text typed before that is lost. Wait until the body has stopped changing.
  let last = null;
  for (let i = 0, same = 0; i < 20 && same < 3; i++) {
    const now = await ed.innerText().catch(() => '');
    same = now === last ? same + 1 : 0; last = now;
    await p.waitForTimeout(500);
  }
  await ed.click();
  if (replace) {
    const sel = await ed.evaluate(e => {
      const kids = [...e.children];
      const stop = kids.findIndex(c => /^(Signature|appendonsend|divRplyFwdMsg)$/.test(c.id) || c.tagName === 'HR' || c.querySelector('#Signature, #divRplyFwdMsg'));
      const r = document.createRange();
      r.setStart(e, 0);
      if (stop < 0) r.setEnd(e, e.childNodes.length); else r.setEndBefore(kids[stop]);
      const s = getSelection(); s.removeAllRanges(); s.addRange(r);
      return !r.collapsed;
    });
    if (sel) { await p.waitForTimeout(200); await p.keyboard.press('Delete'); await p.waitForTimeout(300); }
  }
  await p.keyboard.press('Control+Home');
  for (const [i, line] of body.split('\n').entries()) {
    if (i) await p.keyboard.press('Enter');
    if (line) await p.keyboard.insertText(line);
  }
  await p.waitForTimeout(500);
  const squash = t => t.replace(/[\s\u200b]+/g, ' ').trim();
  const got = squash(await ed.innerText().catch(() => ''));
  if (!got.includes(squash(body))) throw new Error('The text did not all reach the message body; the form is left open (Outlook may have autosaved it to Drafts). Use browser_screenshot.');
}

// A recipient box (To, Cc, Bcc). Outlook has labelled it both "To" and "To recipients".
const recipientBox = (p, field) => p.locator([`[contenteditable=true]`, `div[role=textbox]`, `input`]
  .flatMap(t => [`${t}[aria-label="${field}"]`, `${t}[aria-label="${field} recipients"]`]).join(', ')).filter({ visible: true }).first();

// Wait until an opened compose form can take input; explain why if it never does.
async function waitForCompose(p) {
  const ready = p.locator('[contenteditable=true][aria-label*="Message body"], div[role=textbox][aria-label*="Message body"]').filter({ visible: true }).first();
  for (let i = 0; i < 40; i++) {
    if (await ready.isVisible().catch(() => false)) return;
    if (await sessionExpired(p)) throw new Error('Not signed in to Outlook (session expired). Call outlook_login.');
    await p.waitForTimeout(500);
  }
  throw new Error('The compose form did not open (no message body field). Use browser_screenshot.');
}

// Open a new, empty compose form. Starts from a fresh mail view so a half-open compose pane left by
// an earlier failed call (or the compose ribbon, which has no New mail button) can't get in the way.
async function openNewCompose(p) {
  await p.goto(BASE + '/mail/');
  await p.evaluate(n => { window.name = n; sessionStorage.setItem(n, '1'); }, TAB_NAME).catch(() => {});
  await ensureMail(p);
  const home = p.getByRole('tab', { name: 'Home', exact: true }).first();
  if (await home.isVisible().catch(() => false) && (await home.getAttribute('aria-selected')) !== 'true') { await home.click(); await p.waitForTimeout(500); }
  const btn = p.getByRole('button', { name: /^New (mail|message|email)$/ }).filter({ visible: true }).first();
  await btn.waitFor({ timeout: 15000 }).catch(() => { throw new Error('No "New mail" button on the Home tab. Use browser_screenshot.'); });
  await btn.click();
  await waitForCompose(p);
}

async function addRecipients(p, field, list) {
  if (!list.length) return;
  const box = recipientBox(p, field);
  if (!(await box.isVisible().catch(() => false))) await p.getByRole('button', { name: new RegExp('^(Show )?' + field + '$') }).filter({ visible: true }).first().click().catch(() => {});
  await box.waitFor({ timeout: 10000 }).catch(() => { throw new Error(`The ${field} field of the compose form was not found. Use browser_screenshot.`); });
  await box.click();
  for (const a of list) { await p.keyboard.insertText(a); await p.waitForTimeout(800); await p.keyboard.press(';'); await p.waitForTimeout(400); }
}

// Remove every recipient from a field of the open compose form.
async function clearRecipients(p, field) {
  const box = recipientBox(p, field);
  if (!(await box.isVisible().catch(() => false))) return;
  for (let i = 0; i < 50; i++) {
    const x = box.locator('#removeButton').first();
    if (!(await x.count())) return;
    await x.click(); await p.waitForTimeout(300);
  }
}

// Check local files before opening any form, so a bad path fails without leaving a half-made draft.
function localFiles(list) {
  return list.map(f => {
    const abs = path.resolve(f.replace(/^~(?=\/)/, process.env.HOME));
    let st; try { st = fs.statSync(abs); } catch { throw new Error(`No such file: ${f}`); }
    if (!st.isFile()) throw new Error(`Not a file: ${f}`);
    return abs;
  });
}

// Names (and sizes) of the files attached to the open compose form.
async function composeAttachments(p) {
  return p.evaluate(() => {
    const ed = document.querySelector('[contenteditable=true][aria-label*="Message body"]');
    let root = ed;
    while (root && !root.querySelector('input[aria-label="Subject"], input[aria-label="Add a subject"], input[placeholder="Add a subject"]')) root = root.parentElement;
    return [...(root || document).querySelectorAll('[role=listbox][aria-label="file attachments"] [role=option]')].map(e => {
      const name = (e.querySelector('[title]')?.getAttribute('title') || e.innerText.split('\n')[0]).trim();
      const size = (e.innerText.match(/[\d.,]+ (?:bytes|KB|MB|GB)/) || [''])[0];
      return { name, size, busy: /upload|%/i.test(e.innerText) || !!e.querySelector('[role=progressbar]') };
    });
  });
}

// Attach local files to the open compose form through Attach file > Browse this computer.
async function attachFiles(p, files) {
  if (!files.length) return;
  await clearOverlays(p);
  const before = (await composeAttachments(p)).length;
  const msgTab = p.getByRole('tab', { name: 'Message', exact: true }).first();
  if (await msgTab.isVisible().catch(() => false) && (await msgTab.getAttribute('aria-selected')) !== 'true') { await msgTab.click(); await p.waitForTimeout(500); }
  const btn = p.locator('button[aria-label="Attach file"], button[label="Attach file"]').filter({ visible: true }).first();
  await btn.waitFor({ timeout: 10000 }).catch(() => { throw new Error('No "Attach file" button on the compose ribbon. Use browser_screenshot.'); });
  await btn.click(); await p.waitForTimeout(800);
  const chooser = p.waitForEvent('filechooser', { timeout: 10000 });
  await p.getByRole('menuitem', { name: /Browse this computer/i }).first().click();
  await (await chooser).setFiles(files);
  const names = files.map(f => path.basename(f));
  for (let i = 0; i < 240; i++) {      // up to 2 minutes for uploads
    await p.waitForTimeout(500);
    // Outlook offers to share large files as a OneDrive link instead; that is the user's call.
    const ask = p.locator('[role=dialog][aria-modal=true], [role=alertdialog]').filter({ visible: true });
    if (await ask.count()) { const msg = (await ask.first().innerText()).replace(/\s+/g, ' ').slice(0, 200); await clearOverlays(p); throw new Error(`Outlook asked something while attaching, and it was cancelled: "${msg}". The draft may have been autosaved.`); }
    const now = await composeAttachments(p);
    if (now.length >= before + files.length && names.every(n => now.some(a => a.name === n)) && !now.some(a => a.busy)) return;
  }
  throw new Error(`The attachments did not finish uploading (${names.join(', ')}). Use browser_screenshot.`);
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
  const hit = (await listFolder(p, { limit: 50, match: needle, by_item: true }))[0];
  return hit?.id;
}

// The drafts in Drafts with this id (a conversation id: drafts replying in one conversation share it).
async function draftsWithId(p, id) {
  await openFolder(p, 'Drafts');
  return (await listFolder(p, { limit: 200, by_item: true })).filter(d => d.id === id);
}

// Open a draft from Drafts in the compose form. New-mail drafts open straight into it;
// reply drafts may open as a preview with an Edit (pencil) button. `match` (part of its subject
// or text) picks one when several drafts share the id.
async function openDraftForEdit(p, id, match = '') {
  await openFolder(p, 'Drafts');
  if (!match) {
    const same = await p.locator(`[role=listbox] [role=option][data-convid="${id}"]`).count();
    if (same > 1) throw new Error(`${same} drafts share id ${id} (replies in one conversation). Pass part of the one you mean as \`match\`:\n` + (await draftsWithId(p, id)).map(d => '  ' + d.summary).join('\n'));
  }
  const it = await findItem(p, id, match);
  await it.click();
  await clearOverlays(p);
  const editor = p.locator('[contenteditable=true][aria-label*="Message body"]').first();
  const edit = p.getByRole('button', { name: /^(Edit|Continue editing)/ }).filter({ visible: true }).first();
  for (let i = 0; i < 20 && !(await editor.isVisible().catch(() => false)); i++) {
    if (await sessionExpired(p)) throw new Error('Not signed in to Outlook (session expired). Call outlook_login.');
    if (await edit.isVisible().catch(() => false)) await edit.click();
    await p.waitForTimeout(700);
  }
  await editor.waitFor({ timeout: 10000 });
}

// What the open compose form holds: recipients, subject and body text.
async function composeSummary(p, max = 6000) {
  const c = await p.evaluate(() => {
    const field = f => {
      const el = document.querySelector(`[contenteditable=true][aria-label="${f}"], [contenteditable=true][aria-label="${f} recipients"]`);
      let box = el; for (let i = 0; i < 3 && box && !box.innerText.trim(); i++) box = box.parentElement;
      const t = (box?.innerText || '').replace(/[\u200b\s]+/g, ' ').trim();
      return t ? `${f}: ${t}` : '';
    };
    const subject = document.querySelector('input[aria-label="Subject"]')?.value || '';
    const body = document.querySelector('[contenteditable=true][aria-label*="Message body"]')?.innerText.replace(/\n{3,}/g, '\n\n').trim() || '';
    return { head: [field('To'), field('Cc'), field('Bcc'), 'Subject: ' + subject], body };
  });
  const att = await composeAttachments(p);
  if (att.length) c.head.push('Attachments: ' + att.map(a => `${a.name} (${a.size})`).join(', '));
  return [...c.head.filter(Boolean), '', c.body].join('\n').slice(0, max);
}

async function leaveCompose(p) {
  await p.keyboard.press('Control+s').catch(() => {});
  await p.waitForTimeout(1500);
  await p.goto(BASE + '/mail/');
  await p.locator('[role=tree]').first().waitFor({ timeout: 30000 }).catch(() => {});
}


// Open a conversation either from a folder or from a search (search results are not a folder).
async function openTarget(p, { folder, id, search }) {
  if (search) {
    await ensureMail(p);
    const box = p.getByRole('combobox', { name: /search/i }).or(p.locator('input[aria-label*="Search"]')).first();
    await box.click(); await box.fill(search); await p.keyboard.press('Enter');
    await p.waitForTimeout(4000);
    const it = await findItem(p, id); await it.click();
  } else {
    await openConversation(p, folder || 'Inbox', id, true);
  }
  await p.waitForTimeout(2000);
  await p.evaluate(() => document.querySelectorAll('[aria-label="Reading Pane"] [aria-expanded=false][role=button]').forEach(b => b.click())).catch(() => {});
  await p.waitForTimeout(1500);
}

const ATTACHMENTS = '[aria-label="Reading Pane"] [role=listbox][aria-label="file attachments"] [role=option]';

async function attachmentNames(p) {
  return p.locator(ATTACHMENTS).evaluateAll(es => es.map(e => {
    const name = e.querySelector('[title]')?.getAttribute('title') || e.innerText.split('\n')[0];
    const size = (e.innerText.match(/[\d.,]+ (?:bytes|KB|MB|GB)/) || [''])[0];
    return { name: name.trim(), size };
  }));
}

async function attachmentMenu(p, index, item) {
  const opt = p.locator(ATTACHMENTS).nth(index);
  await opt.scrollIntoViewIfNeeded().catch(() => {});
  await opt.hover();
  await opt.locator('button[aria-label="More actions"]').click();
  await p.waitForTimeout(900);
  await p.getByRole('menuitem').filter({ hasText: new RegExp('^[\\s\\uE000-\\uF8FF]*' + item, 'i') }).first().click();
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
    // Account picker: take OUTLOOK_ACCOUNT, or the only account listed (the picker may show the
    // username form, e.g. uqxxxxx@uq.edu.au, rather than the email address).
    if (/login\.microsoftonline/.test(p.url()) && /Pick an account/i.test(await p.evaluate(() => document.body.innerText).catch(() => ''))) {
      const tiles = p.getByText(/^[^\s@]+@[^\s@]+\.[^\s@]+$/).filter({ visible: true });
      const want = ACCOUNT ? tiles.filter({ hasText: ACCOUNT }) : tiles;
      const pick = (await want.count()) ? want.first() : (await tiles.count()) === 1 ? tiles.first() : null;
      if (pick) { await pick.click(); await p.waitForTimeout(3000); continue; }
      return 'The Microsoft sign-in page asks which account to use and none matches OUTLOOK_ACCOUNT. The user needs to pick it in the browser.';
    }
    const body = await p.evaluate(() => document.body.innerText).catch(() => '');
    const num = body.match(/(?:tap|select|enter|number)[^\n]{0,80}?\b(\d{2,3})\b/i) || body.match(/\n\s*(\d{2,3})\s*\n/);
    // (Okta can be on the organisation's own domain, e.g. signin.uq.edu.au, so don't rely on the URL.)
    if (num && !/outlook\./i.test(p.url()) && /verify|approve|authenticator|okta/i.test(body)) {
      if (!wait_seconds) return `Approval requested: the user should tap or enter ${num[1]} in their authenticator app. Then call outlook_login with wait_seconds=60.`;
      await p.waitForTimeout(3000); continue;
    }
    // Okta may default to typing a code from Okta Verify: switch to a push (number matching) instead.
    const codeBox = p.locator('input[name="credentials.totp"], input[name="credentials.passcode"], input[name=passcode]').filter({ visible: true }).first();
    if (await codeBox.count() && /Enter a code/i.test(body)) {
      const other = p.getByRole('link', { name: /Verify with something else/i }).or(p.getByRole('button', { name: /Verify with something else/i })).first();
      if (await other.isVisible().catch(() => false)) { await other.click(); await p.waitForTimeout(3000); continue; }
      return 'Okta asks for a code from Okta Verify and offers no push option. The user needs to sign in in the browser.';
    }
    const pushOption = p.getByRole('link', { name: /^Select Get a push notification/i }).or(p.getByRole('button', { name: /^Select Get a push notification/i })).first();
    if (await pushOption.isVisible().catch(() => false)) { await pushOption.click(); await p.waitForTimeout(3000); continue; }
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
      // Submit from the field itself (Enter with focus elsewhere does nothing), or its Next button.
      await user.focus(); await user.press('Enter'); await p.waitForTimeout(4000);
      if (await user.isVisible().catch(() => false)) { const nx = p.locator('input[type=submit][value="Next"], button[type=submit]').filter({ visible: true }).first(); if (await nx.count()) { await nx.click(); await p.waitForTimeout(4000); } }
      continue;
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
  const l = await listFolder(p, { unread_only, limit, match, by_item: /^drafts$/i.test(folder) });
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

const ATTACH = z.array(z.string()).default([]).describe('Local file paths to attach');
const LEAD = body => body.split('\n').find(l => l.trim())?.trim().slice(0, 40) || '';

// Change the open compose form (a draft being edited) and save it.
async function editOpenDraft(p, { body, subject, to, cc, bcc, attachments }) {
  for (const [field, list] of [['To', to], ['Cc', cc], ['Bcc', bcc]]) if (list) { await clearRecipients(p, field); await addRecipients(p, field, list); }
  if (subject !== undefined) {
    const subj = p.locator('input[aria-label="Subject"], input[aria-label="Add a subject"], input[placeholder="Add a subject"]').filter({ visible: true }).first();
    await subj.click(); await subj.fill(subject);
  }
  await attachFiles(p, attachments || []);
  if (body !== undefined) await fillBody(p, body, { replace: true });
  const summary = await composeSummary(p, 3000);
  await saveDraft(p);
  return summary;
}

server.registerTool('outlook_draft_reply', {
  description: 'Write a reply, reply-all or forward for a conversation and save it as a draft (does not send). The text goes above the quoted thread. Forward needs `to`. `attachments` are local file paths. If the conversation already has an unsent draft it stops and says so, unless replace_existing is true: then the existing draft gets this text (and any to/cc/attachments) instead of a second draft being made. Returns the draft id for outlook_send_draft.',
  inputSchema: {
    folder: z.string().default('Inbox'), id: z.string(), body: z.string(),
    mode: z.enum(['reply', 'reply_all', 'forward']).default('reply'),
    to: z.array(z.string()).default([]), cc: z.array(z.string()).default([]),
    attachments: ATTACH,
    replace_existing: z.boolean().default(false).describe('If the conversation already has a draft, replace that draft\'s text instead of stopping'),
  },
}, tool(async ({ folder, id, body, mode, to, cc, attachments, replace_existing }) => {
  if (mode === 'forward' && !to.length) throw new Error('A forward needs at least one address in `to`.');
  const files = localFiles(attachments);
  const p = await page();
  // Drafts replying in a conversation carry its id. (The list's "[Draft]" mark can lag, so look in Drafts.)
  const drafts = await draftsWithId(p, id);
  if (drafts.length) {
    const list = drafts.map(d => '  ' + d.summary).join('\n');
    const what = drafts.length > 1 ? `${drafts.length} unsent drafts` : 'an unsent draft';
    if (!replace_existing) throw new Error(`This conversation already has ${what} (draft id ${id}):\n${list}\nNothing was changed. To put this text in that draft instead, call again with replace_existing: true (or use outlook_edit_draft). To start afresh, outlook_discard_draft it first.`);
    if (drafts.length > 1) throw new Error(`This conversation has ${what} (draft id ${id}):\n${list}\nUse outlook_edit_draft with \`match\` to pick one, or discard the extras.`);
    await openDraftForEdit(p, id);
    await editOpenDraft(p, { body, to: to.length ? to : undefined, cc: cc.length ? cc : undefined, attachments: files });
    return `Replaced the text of the existing draft (not sent). Draft id: ${id}`;
  }
  await openConversation(p, folder, id, true);
  await p.waitForTimeout(1500); await clearOverlays(p);
  await clickReply(p, mode);
  await waitForCompose(p);
  await addRecipients(p, 'To', to);
  await addRecipients(p, 'Cc', cc);
  await attachFiles(p, files);
  await fillBody(p, body);
  await saveDraft(p);
  const draft = await draftIdFor(p, LEAD(body));
  return `Draft saved (not sent). Draft id: ${draft || '(not found in Drafts yet; list Drafts)'}`;
}));

server.registerTool('outlook_draft_new', {
  description: 'Write a new email and save it as a draft (does not send). to/cc/bcc are lists of email addresses; `attachments` are local file paths. Returns the draft id for outlook_send_draft.',
  inputSchema: { to: z.array(z.string()).default([]), cc: z.array(z.string()).default([]), bcc: z.array(z.string()).default([]), subject: z.string(), body: z.string(), attachments: ATTACH },
}, tool(async ({ to, cc, bcc, subject, body, attachments }) => {
  const files = localFiles(attachments);
  const p = await page(); await openNewCompose(p);
  await addRecipients(p, 'To', to);
  await addRecipients(p, 'Cc', cc);
  await addRecipients(p, 'Bcc', bcc);
  const subj = p.locator('input[aria-label="Subject"], input[aria-label="Add a subject"], input[placeholder="Add a subject"]').filter({ visible: true }).first();
  await subj.click(); await subj.fill(subject);
  await attachFiles(p, files);
  await fillBody(p, body);
  await saveDraft(p);
  const draft = await draftIdFor(p, subject);
  return `Draft saved (not sent). Draft id: ${draft || '(not found in Drafts yet; list Drafts)'}`;
}));

server.registerTool('outlook_edit_draft', {
  description: 'Change a draft in Drafts and save it (does not send). Only the fields given change: `body` replaces the text written above the signature and quoted thread (both kept); to/cc/bcc replace that field\'s recipients; `subject` replaces the subject; `attachments` (local file paths) are added. `match` (part of its subject or text) picks the draft when several share one id, as replies in one conversation do.',
  inputSchema: {
    id: z.string(), match: z.string().default(''),
    body: z.string().optional(), subject: z.string().optional(),
    to: z.array(z.string()).optional(), cc: z.array(z.string()).optional(), bcc: z.array(z.string()).optional(),
    attachments: ATTACH,
  },
}, tool(async ({ id, match, body, subject, to, cc, bcc, attachments }) => {
  const files = localFiles(attachments);
  if ([body, subject, to, cc, bcc].every(x => x === undefined) && !files.length) throw new Error('Nothing to change: give body, subject, to, cc, bcc or attachments.');
  const p = await page(); await openDraftForEdit(p, id, match);
  const summary = await editOpenDraft(p, { body, subject, to, cc, bcc, attachments: files });
  return `Draft updated (not sent). Draft id: ${id}\n\n${summary}`;
}));

server.registerTool('outlook_read_draft', {
  description: 'Show a draft from Drafts (recipients, subject, attachment names and text) so the user can check it before sending. `match` (part of its subject or text) picks the draft when several share one id.',
  inputSchema: { id: z.string(), match: z.string().default(''), max_chars: z.number().int().default(6000) },
}, tool(async ({ id, match, max_chars }) => {
  const p = await page(); await openDraftForEdit(p, id, match);
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
  const it = await findItem(p, id).catch(() => null);
  if (!it) throw new Error(`No draft with id ${id} in Drafts.`);
  if (!(await p.locator(`[role=listbox] [role=option][data-convid="${id}"]`).filter({ hasText: expected_subject }).count()) &&
      !((await it.getAttribute('aria-label')) || '').toLowerCase().includes(expected_subject.toLowerCase()))
    throw new Error(`Draft ${id} does not match "${expected_subject}": ${((await it.getAttribute('aria-label')) || '').slice(0, 150)}`);
  if ((await p.locator(`[role=listbox] [role=option][data-convid="${id}"]`).count()) > 1)
    throw new Error(`Several drafts share id ${id} (replies in one conversation); not sent. Discard the extras with outlook_discard_draft first.`);
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
  description: 'Throw away a draft from Drafts (for example one made by mistake). When several drafts share the id (replies in one conversation), the one whose subject or text contains expected_subject goes.',
  inputSchema: { id: z.string(), expected_subject: z.string().describe('Part of the draft subject or text, as a check.') },
}, tool(async ({ id, expected_subject }) => {
  const p = await page(); await openFolder(p, 'Drafts');
  const rows = () => p.locator(`[role=listbox] [role=option][data-convid="${id}"]`).filter({ hasText: expected_subject });
  const it = await findItem(p, id, expected_subject).catch(async () => {
    // the text may be in the preview rather than the label
    const r = rows().first(); if (await r.count()) return r;
    throw new Error('No draft with that id matches expected_subject.');
  });
  const before = await rows().count();
  await it.click(); await p.waitForTimeout(2500); await clearOverlays(p);
  const bin = p.locator('[aria-label="Reading Pane"], [role=main]').getByRole('button', { name: /^(Discard|Delete)/ }).filter({ visible: true }).first();
  await bin.click(); await p.waitForTimeout(1500);
  const ok = p.locator('[role=dialog], [role=alertdialog]').getByRole('button', { name: /^(OK|Discard|Delete)$/ }).filter({ visible: true });
  if (await ok.count()) await ok.first().click();
  await p.waitForTimeout(2500);
  // Reload so the list and reading pane show the change, not a stale view.
  await p.goto(BASE + '/mail/'); await p.locator('[role=tree]').first().waitFor({ timeout: 30000 }).catch(() => {});
  await openFolder(p, 'Drafts');
  // The list can lag behind the discard; give it a few seconds before saying it failed.
  for (let i = 0; i < 5 && await rows().count() >= before; i++) await p.waitForTimeout(2000);
  const left = await rows().count();
  return left >= before ? 'Still in Drafts.' : left ? `Discarded one; ${left} more draft(s) with this id and text remain.` : 'Discarded.';
}));


server.registerTool('outlook_attachments', {
  description: 'List the attachments in a conversation (every message in the thread), or download them. Give folder+id, or search+id for a conversation found with outlook_search (pass the same query). action "download" saves files to a local folder; "save_to_onedrive" uses Outlook\'s Save to OneDrive (it lands in OneDrive/Attachments). `names` limits it to attachments whose name contains any of the given strings.',
  inputSchema: {
    id: z.string(), folder: z.string().optional(), search: z.string().optional(),
    action: z.enum(['list', 'download', 'save_to_onedrive']).default('list'),
    names: z.array(z.string()).default([]),
    dest: z.string().optional().describe('Download folder (default: OUTLOOK_DOWNLOAD_DIR)'),
  },
}, tool(async ({ id, folder, search, action, names, dest }) => {
  const p = await page(); await openTarget(p, { folder, id, search });
  const all = await attachmentNames(p);
  if (!all.length) return 'No attachments in this conversation.';
  const pick = all.map((a, i) => ({ ...a, i })).filter(a => !names.length || names.some(n => a.name.toLowerCase().includes(n.toLowerCase())));
  if (action === 'list') return all.map(a => `${a.name} (${a.size})`).join('\n');
  if (!pick.length) return 'None of the attachments match: ' + all.map(a => a.name).join(', ');
  const out = [];
  if (action === 'save_to_onedrive') {
    for (const a of pick) { await attachmentMenu(p, a.i, 'Save to OneDrive'); await p.waitForTimeout(3000); await clearOverlays(p); out.push(`${a.name}: sent to OneDrive (Attachments folder)`); }
    return out.join('\n');
  }
  const dir = path.resolve(dest || DOWNLOAD_DIR);
  fs.mkdirSync(dir, { recursive: true });
  const cdp = await p.context().newCDPSession(p);
  await cdp.send('Browser.setDownloadBehavior', { behavior: 'allow', downloadPath: dir, eventsEnabled: true });
  try {
    for (const a of pick) {
      const done = new Promise(r => { const h = e => { if (e.state !== 'inProgress') { cdp.off('Browser.downloadProgress', h); r(e.state); } }; cdp.on('Browser.downloadProgress', h); });
      const named = new Promise(r => cdp.once('Browser.downloadWillBegin', e => r(e.suggestedFilename)));
      await attachmentMenu(p, a.i, 'Download');
      const state = await Promise.race([done, new Promise(r => setTimeout(() => r('timed out'), 120000))]);
      const file = await Promise.race([named, new Promise(r => setTimeout(() => r(a.name), 100))]);
      out.push(state === 'completed' ? path.join(dir, file) : `${a.name}: download ${state}`);
    }
  } finally {
    await cdp.send('Browser.setDownloadBehavior', { behavior: 'default' }).catch(() => {});
  }
  return out.join('\n');
}));

server.registerTool('outlook_people', {
  description: 'Look people up in your contacts and the organisation directory (name, email or part of either). Returns name, job title, email addresses and phone.',
  inputSchema: { query: z.string(), limit: z.number().int().min(1).max(50).default(10) },
}, tool(async ({ query, limit }) => {
  const p = await page();
  const pp = await p.context().newPage();
  try {
    await pp.goto(BASE + '/people/');
    await pp.locator('#topSearchInput').waitFor({ timeout: 30000 });
    await pp.waitForTimeout(1500); await clearOverlays(pp);
    const box = pp.locator('#topSearchInput');
    await box.click(); await box.fill(query); await pp.keyboard.press('Enter');
    await pp.locator('[data-tid=contacts-list-grid-row]').first().waitFor({ timeout: 15000 }).catch(() => {});
    await pp.waitForTimeout(1500);
    const rows = await pp.locator('[data-tid=contacts-list-grid-row]').evaluateAll(es => es.map(e =>
      e.innerText.replace(/[\uE000-\uF8FF]/g, '').split('\n').map(x => x.trim()).filter(x => x && x !== '•')));
    if (!rows.length) return 'No one found.';
    const isPhone = x => /^[+\d][\d ()-]{3,}$/.test(x);
    return rows.slice(0, limit).map(r => {
      if (/^[A-Z]{1,3}$/.test(r[0])) r = r.slice(1);          // avatar initials
      const [name, ...rest] = r;
      return [name, rest.filter(x => !/@/.test(x) && !isPhone(x)).join('; '), rest.filter(x => /@/.test(x)).join(', '), rest.filter(isPhone).join(', ')].filter(Boolean).join(' | ');
    }).join('\n');
  } finally { await pp.close().catch(() => {}); }
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
  const cp = await calendarTab();
  try {
    for (let i = 0; i < days; i += 7) {
      const d = new Date(start); d.setDate(d.getDate() + i);
      await gotoDate(cp, `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`);
      for (const l of await cp.evaluate(() => [...document.querySelectorAll('[role=button][aria-label]')].map(e => e.getAttribute('aria-label').replace(/ 0(\d), (\d{4})/, ' $1, $2')))) {
        const day = want.find(w => l.includes(w));
        if (day && /\d{1,2}:\d\d|all day/i.test(l)) seen.set(l, day);
      }
    }
    // a range that crosses into a week not visited (e.g. starts on Sunday) needs the next week too
    const last = new Date(start); last.setDate(last.getDate() + days - 1);
    if (!want.every(w => [...seen.values()].includes(w)) && days > 1) {
      await gotoDate(cp, `${last.getFullYear()}-${String(last.getMonth() + 1).padStart(2, '0')}-${String(last.getDate()).padStart(2, '0')}`);
      for (const l of await cp.evaluate(() => [...document.querySelectorAll('[role=button][aria-label]')].map(e => e.getAttribute('aria-label').replace(/ 0(\d), (\d{4})/, ' $1, $2')))) {
        const day = want.find(w => l.includes(w));
        if (day && /\d{1,2}:\d\d|all day/i.test(l)) seen.set(l, day);
      }
    }
  } finally { await cp.close(); }
  const byDay = {};
  for (const [l, d] of seen) (byDay[d] ||= []).push(l.replace(', ' + d, ''));
  return Object.keys(byDay).length ? byDay : 'No events found.';
}));


// ---- Calendar: events are found by date + title in the day view, then opened in the event view.

async function calendarTab() {
  const cp = await (await page()).context().newPage();
  await (await cp.context().newCDPSession(cp)).send('Emulation.setTimezoneOverride', { timezoneId: TIMEZONE }).catch(() => {});
  return cp;
}

async function openEvent(cp, date, title) {
  const [y, m, d] = date.split('-').map(Number);
  await gotoDate(cp, date);
  const day = new Date(y, m - 1, d).toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' });
  const items = cp.locator('[role=button][aria-label]').filter({ hasText: new RegExp(title.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&'), 'i') });
  const labels = await items.evaluateAll((es, day) => es.map(e => e.getAttribute('aria-label')).filter(l => l.replace(/ 0(\d), (\d{4})/, ' $1, $2').includes(day)), day);
  if (!labels.length) throw new Error(`No event matching "${title}" on ${date}. Use outlook_calendar to see that day.`);
  if (labels.length > 1) throw new Error(`Several events match "${title}" on ${date}: ${labels.join(' / ')}. Use a more specific title.`);
  await cp.locator(`[role=button][aria-label="${labels[0].replace(/"/g, '\\"')}"]`).first().click();
  await cp.waitForTimeout(1500);
  await cp.getByRole('button', { name: 'View event', exact: true }).click();
  // the event window's dialog wrapper reports as hidden; wait for one of its buttons instead
  await cp.getByRole('button', { name: 'Open in new window' }).last().waitFor({ timeout: 15000 });
  const dlg = cp.locator('[role=dialog]').filter({ has: cp.getByRole('button', { name: 'Open in new window' }) }).last();
  await cp.waitForTimeout(2000);
  return dlg;
}


// The calendar ignores dates in the URL, so navigate with the mini calendar: page to the month,
// then click the day; the main view (week) moves to the week holding that day.
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
async function gotoDate(cp, date) {
  const [y, m, d] = date.split('-').map(Number);
  if (!/\/calendar\/view\/week/.test(cp.url())) await cp.goto(`${BASE}/calendar/view/week`);
  await cp.waitForTimeout(1000); await ensureCalendar(cp);
  const header = cp.getByRole('button', { name: /select to change the month/ }).first();
  await header.waitFor({ timeout: 30000 });
  await cp.waitForTimeout(1500); await clearOverlays(cp);
  const dayBtn = cp.getByRole('button', { name: `${d}, ${MONTHS[m - 1]}, ${y}`, exact: true }).first();
  for (let i = 0; i < 48 && !(await dayBtn.isVisible().catch(() => false)); i++) {
    const shown = (await header.getAttribute('aria-label')).match(/(\d{4}) (\w+)/);
    const diff = (y - +shown[1]) * 12 + (m - 1 - MONTHS.indexOf(shown[2]));
    // the two unlabelled buttons after the month title are previous and next month
    const arrows = header.locator('xpath=following::button[normalize-space(@aria-label)="" or not(@aria-label)]');
    await arrows.nth(diff < 0 ? 0 : 1).click();
    await cp.waitForTimeout(500);
  }
  await dayBtn.click();
  await cp.waitForTimeout(2500);
}

const RESPONSE = { accept: 'Accept', tentative: 'Tentative', decline: 'Decline' };

server.registerTool('outlook_event', {
  description: 'Open a calendar event by date (YYYY-MM-DD) and part of its title: organiser, attendees and their responses, time, location, Teams link, your response and the description.',
  inputSchema: { date: z.string(), title: z.string(), max_chars: z.number().int().default(5000) },
}, tool(async ({ date, title, max_chars }) => {
  const cp = await calendarTab();
  try {
    const dlg = await openEvent(cp, date, title);
    const t = await dlg.innerText();
    return t.replace(/[\uE000-\uF8FF]/g, '').replace(/[ \t\u00a0]+\n/g, '\n').replace(/\n{2,}/g, '\n').slice(0, max_chars);
  } finally { await cp.close().catch(() => {}); }
}));

server.registerTool('outlook_event_respond', {
  description: 'Accept, tentatively accept or decline a meeting invitation, optionally with a note to the organiser. This emails the organiser: only call it after the user has said how to respond to this specific invitation. Disabled unless OUTLOOK_ENABLE_SEND=1.',
  inputSchema: { date: z.string(), title: z.string(), response: z.enum(['accept', 'tentative', 'decline']), note: z.string().default('') },
}, tool(async ({ date, title, response, note }) => {
  if (!SEND_ENABLED) throw new Error('Responding to invitations is disabled (it emails the organiser). Start the server with OUTLOOK_ENABLE_SEND=1.');
  const cp = await calendarTab();
  try {
    const dlg = await openEvent(cp, date, title);
    const word = RESPONSE[response];
    // Already responded: the response buttons are folded into a menu named after the current response.
    const menuBtn = dlg.locator('button[aria-haspopup=menu]').filter({ hasText: /Accepted|Tentative|Declined/ }).first();
    const hasMenu = await menuBtn.isVisible().catch(() => false);
    if (note) {
      const direct = dlg.getByRole('button', { name: 'RSVP with note', exact: true });
      if (await direct.isVisible().catch(() => false)) await direct.click();
      else { await menuBtn.click(); await cp.waitForTimeout(800); await cp.getByRole('menuitem', { name: 'RSVP with note', exact: true }).click(); }
      const rsvp = cp.getByRole('dialog').filter({ hasText: 'RSVP to' }).last();
      await rsvp.waitFor({ timeout: 10000 });
      await rsvp.locator('textarea, [contenteditable=true]').first().fill(note);
      await rsvp.getByRole('combobox', { name: 'Select an option' }).click(); await cp.waitForTimeout(700);
      await cp.getByRole('option', { name: word, exact: true }).click(); await cp.waitForTimeout(500);
      await rsvp.getByRole('button', { name: 'Confirm', exact: true }).click();
    } else if (hasMenu) {
      if ((await menuBtn.innerText()).includes(word === 'Accept' ? 'Accepted' : word === 'Decline' ? 'Declined' : 'Tentative')) return `Already ${response === 'accept' ? 'accepted' : response === 'decline' ? 'declined' : 'tentative'}; nothing sent.`;
      await menuBtn.click(); await cp.waitForTimeout(800);
      await cp.getByRole('menuitem', { name: word, exact: true }).click();
    } else {
      await dlg.getByRole('button', { name: word, exact: true }).click();
    }
    await cp.waitForTimeout(3000);
    const ask = cp.locator('[role=alertdialog], [role=dialog][aria-modal=true]').filter({ hasNotText: title }).filter({ visible: true });
    if (await ask.count()) { const msg = (await ask.first().innerText()).replace(/\s+/g, ' ').slice(0, 300); await clearOverlays(cp); throw new Error(`Outlook asked something and it was cancelled: "${msg}". Check the event with outlook_event.`); }
    return `Responded: ${word}${note ? ' (with note)' : ''}.`;
  } finally { await cp.close().catch(() => {}); }
}));

server.registerTool('outlook_event_create', {
  description: 'Create a calendar event: title, date (YYYY-MM-DD), start and end (HH:MM, 24-hour) or all_day, location, optional Microsoft Teams meeting, description. With no attendees it is just saved in the calendar. With attendees it sends them invitations: only after the user has approved this event and its invite list (needs OUTLOOK_ENABLE_SEND=1). Same-day events only.',
  inputSchema: {
    title: z.string(), date: z.string(), start: z.string().optional(), end: z.string().optional(), all_day: z.boolean().default(false),
    location: z.string().default(''), teams: z.boolean().default(false), attendees: z.array(z.string()).default([]), body: z.string().default(''),
  },
}, tool(async ({ title, date, start, end, all_day, location, teams, attendees, body }) => {
  if (attendees.length && !SEND_ENABLED) throw new Error('Inviting attendees sends email, which is disabled. Start the server with OUTLOOK_ENABLE_SEND=1, or create the event without attendees.');
  if (!all_day && !(start && end)) throw new Error('Give start and end (HH:MM), or all_day: true.');
  const cp = await calendarTab();
  try {
    await cp.goto(`${BASE}/calendar/view/week`);
    await cp.waitForTimeout(1000); await ensureCalendar(cp);
    await cp.getByRole('button', { name: 'New event', exact: true }).first().waitFor({ timeout: 30000 });
    await cp.waitForTimeout(2000); await clearOverlays(cp);
    await cp.getByRole('button', { name: 'New event', exact: true }).first().click();
    // the form's dialog wrapper reports as hidden, so work from the page (only one form is open)
    const form = cp.locator('body');
    await cp.locator('input[placeholder="Add title"]').waitFor({ timeout: 15000 });
    await form.locator('input[placeholder="Add title"]').fill(title);
    if (attendees.length) {
      await form.locator('[aria-label="Invite required attendees"]').click();
      for (const a of attendees) { await cp.keyboard.insertText(a); await cp.waitForTimeout(900); await cp.keyboard.press(';'); await cp.waitForTimeout(400); }
    }
    // the date/time summary opens the start date, start time, end time and all-day fields
    await form.getByText(/\d{4}-\d\d-\d\d \d\d:\d\d - /).first().click();
    const setField = async (name, value) => {
      const f = form.getByRole('combobox', { name, exact: true });
      await f.click(); await f.fill(value); await cp.keyboard.press('Enter'); await cp.waitForTimeout(500);
    };
    await setField('Start date', date);
    if (all_day) { const ad = form.getByRole('checkbox', { name: 'All day' }); if (!(await ad.isChecked())) await ad.check(); }
    else { await setField('Start time', start); await setField('End time', end); }
    await cp.locator('input[placeholder="Add title"]').click();     // closes the date/time panel
    await cp.waitForTimeout(800);
    if (location) {
      const loc = cp.locator('input[aria-label="Add a room or location"]');
      await loc.click(); await loc.fill(location); await cp.waitForTimeout(800); await cp.keyboard.press('Tab');
    }
    if (teams) { const sw = form.getByRole('switch', { name: /Microsoft Teams meeting/ }); if (!(await sw.isChecked())) await sw.click(); }
    if (body) { const b = form.getByRole('textbox', { name: 'Event body' }); await b.click(); await cp.keyboard.insertText(body); }
    await cp.waitForTimeout(1000);
    const summary = (await cp.locator('[role=dialog]').filter({ has: cp.locator('input[placeholder="Add title"]') }).last().innerText().catch(() => '')).replace(/[\uE000-\uF8FF]/g, '').replace(/\s+/g, ' ').slice(0, 600);
    const saveName = attendees.length ? 'Send' : 'Save';
    await form.getByRole('button', { name: saveName, exact: true }).click();
    await cp.waitForTimeout(4000);
    const ask = cp.locator('[role=alertdialog], [role=dialog][aria-modal=true]').filter({ visible: true }).filter({ hasNotText: 'Add title' });
    if (await ask.count()) { const msg = (await ask.first().innerText()).replace(/\s+/g, ' ').slice(0, 300); await clearOverlays(cp); throw new Error(`Outlook asked something and it was cancelled: "${msg}"`); }
    return `${attendees.length ? 'Sent invitations for' : 'Saved'} "${title}" on ${date}. Form before saving: ${summary}`;
  } finally { await cp.close().catch(() => {}); }
}));

server.registerTool('outlook_event_delete', {
  description: 'Delete an event from the calendar (date + part of its title). For a meeting you organised with attendees this sends them a cancellation: only after the user approves it (needs OUTLOOK_ENABLE_SEND=1). To turn down someone else\'s meeting use outlook_event_respond with decline.',
  inputSchema: { date: z.string(), title: z.string() },
}, tool(async ({ date, title }) => {
  const cp = await calendarTab();
  try {
    const dlg = await openEvent(cp, date, title);
    await cp.getByRole('button', { name: 'Delete', exact: true }).filter({ visible: true }).first().click();
    await cp.waitForTimeout(1500);
    // confirmation: "Delete event" (no attendees) or a cancellation to send (meeting you organised)
    const box = cp.locator('[role=dialog], [role=alertdialog]')
      .filter({ hasNot: cp.getByRole('button', { name: 'Open in new window' }) })     // not the event window itself
      .filter({ hasText: /Delete event|Are you sure|cancellation|Cancel meeting/i }).last();
    const sendBtn = box.getByRole('button', { name: /^Send( cancellation)?$/ });
    const delBtn = box.getByRole('button', { name: 'Delete', exact: true });
    if (await sendBtn.count()) {
      if (!SEND_ENABLED) { await clearOverlays(cp); throw new Error('Deleting this meeting sends a cancellation to its attendees, which is disabled.'); }
      await sendBtn.last().click();
    } else if (await delBtn.count()) {
      await delBtn.last().click();
    } else {
      throw new Error('Outlook did not ask to confirm the deletion; check the calendar with outlook_calendar.');
    }
    await cp.waitForTimeout(2500);
    await gotoDate(cp, date);
    const left = await cp.locator('[role=button][aria-label]').filter({ hasText: new RegExp(title.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&'), 'i') }).count();
    if (left) throw new Error('Outlook still shows the event after deleting; check it with outlook_event.');
    return 'Deleted.';
  } finally { await cp.close().catch(() => {}); }
}));

server.registerTool('browser_screenshot', {
  description: 'Screenshot the Outlook tab, to see what state it is in.',
  inputSchema: {},
}, async () => serial(async () => {
  const p = await page();
  return { content: [{ type: 'image', data: (await p.screenshot()).toString('base64'), mimeType: 'image/png' }] };
}));

await server.connect(new StdioServerTransport());
