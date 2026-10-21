/**
 * RSVP backend for Anushri & Bharath's wedding site.
 * Lives inside your Google Sheet (Extensions → Apps Script). See README.md for setup.
 *
 * Guests tab: one row per person. People who share a Household value share one link.
 * Menu "Wedding RSVP → Create links" fills Code, Link and a ready-to-send Message.
 */

// ── Settings ───────────────────────────────────────────────────────────────
const SITE_URL = 'https://bharathsatheesh.github.io/goa-wedding/';

// Event ids must match CONFIG.events in index.html.
const EVENTS = [
  { id: 'sangeet', header: 'Sangeet' },
  { id: 'kerala', header: 'Kerala ceremony' },
  { id: 'wedding', header: 'Wedding' },
];

// Invite wording. {names} = the names on that link exactly as written in the Name column.
// Households starting "Anushri's friends" / "Bharath's friends" get the friends version
// (written as that person); everyone else gets the family version.
const MESSAGE_FAMILY =
  "Hi {names}! Anushri & Bharath are getting married in Goa 30–31 March 2027. " +
  "They're doing a very intimate ceremony and would love to have you join in the celebration! " +
  "Please RSVP here by Nov 15th: {link}\n\n" +
  "Bring your best outfits, your dancing shoes, and all your north-south rivalry myths.";
const MESSAGE_FRIENDS =
  "Hi {names}! {partner} & I are getting married in Goa 30–31 March 2027. " +
  "We're doing a very intimate ceremony and would love to have you join in the celebration! " +
  "Please RSVP here by Nov 15th: {link}\n\n" +
  "Bring your best outfits, your dancing shoes, and all your north-south rivalry myths.";

function messageFor_(household, names, link) {
  const hh = String(household).trim().toLowerCase();
  let partner = '';
  if (hh.indexOf("anushri's friends") === 0) partner = 'Bharath';
  if (hh.indexOf("bharath's friends") === 0) partner = 'Anushri';
  const tpl = partner ? MESSAGE_FRIENDS.replace('{partner}', partner) : MESSAGE_FAMILY;
  return tpl.replace('{names}', joinNames_(names)).replace('{link}', link);
}

// ── Sheet layout (don't rename these headers) ──────────────────────────────
const SHEET = 'Guests';
const H = {
  household: 'Household',
  name: 'Name',
  code: 'Code',
  link: 'Link',
  message: 'Message',
  attending: 'Attending',
  perform: 'Sangeet act',
  dietary: 'Dietary',
  email: 'Email',
  accom: 'Accommodation',
  note: 'Note',
  updated: 'Last updated',
};
const HEADERS = [H.household, H.name, H.code, H.link, H.message, H.attending]
  .concat(EVENTS.map(function (e) { return e.header; }))
  .concat([H.perform, H.dietary, H.email, H.accom, H.note, H.updated]);
// Columns added after the first version; table_ creates them if an older Sheet lacks them.
const ADDED_LATER = ['email', 'accom'];

// ── Menu & one-time setup ──────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Wedding RSVP')
    .addItem('Create links', 'createLinks')
    .addItem('Set up sheet (first time only)', 'setup')
    .addToUi();
}

/** Run once: creates the Guests and Summary tabs with a test row for the two of you. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) sh = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0
    ? ss.getSheets()[0].setName(SHEET)
    : ss.insertSheet(SHEET);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.getRange(2, 1, 2, 2).setValues([['Us (test)', 'Anushri'], ['Us (test)', 'Bharath']]);
    sh.setFrozenRows(1);
    sh.setColumnWidth(HEADERS.indexOf(H.message) + 1, 360);
    sh.setColumnWidth(HEADERS.indexOf(H.link) + 1, 260);
  }

  let sum = ss.getSheetByName('Summary');
  if (!sum) sum = ss.insertSheet('Summary');
  const col = function (header) { return colLetter_(HEADERS.indexOf(header) + 1); };
  const range = function (header) { return SHEET + '!' + col(header) + '2:' + col(header); };
  const rows = [
    ['Invited', '=COUNTA(' + range(H.name) + ')'],
    ['Replied', '=COUNTA(' + range(H.attending) + ')'],
    ['Coming', '=COUNTIF(' + range(H.attending) + ',"Yes")'],
    ["Can't make it", '=COUNTIF(' + range(H.attending) + ',"No")'],
    ['Waiting to hear', '=B1-B2'],
  ]
    .concat(EVENTS.map(function (e) { return [e.header, '=COUNTIF(' + range(e.header) + ',"Yes")']; }))
    .concat([['Up for a sangeet act', '=COUNTIF(' + range(H.perform) + ',"Yes")']]);
  sum.clear();
  sum.getRange(1, 1, rows.length, 2).setValues(rows);
  sum.getRange(1, 1, rows.length, 1).setFontWeight('bold');
  onOpen();
}

/**
 * Gives every named guest without a code a code and a link, and writes a message on the first row
 * of each household whose Message cell is empty. Existing codes and messages are never changed,
 * so hand-edited messages are safe. To regenerate a message, clear its cell and run this again.
 */
function createLinks() {
  const sh = sheet_();
  const t = table_(sh);
  const codes = t.rows.map(function (r) { return normCode_(r[t.c.code]); });
  const taken = {};
  const byHousehold = {};
  t.rows.forEach(function (r, i) {
    const hh = String(r[t.c.household]).trim();
    if (codes[i]) { taken[codes[i]] = true; if (hh && !byHousehold[hh]) byHousehold[hh] = codes[i]; }
  });
  t.rows.forEach(function (r, i) {
    if (!String(r[t.c.name]).trim() || codes[i]) return;
    const hh = String(r[t.c.household]).trim();
    let code = hh && byHousehold[hh];
    if (!code) {
      do { code = newCode_(); } while (taken[code]);
      taken[code] = true;
      if (hh) byHousehold[hh] = code;
    }
    codes[i] = code;
  });

  const base = SITE_URL.replace(/\/?$/, '/');
  const seen = {};
  const links = [];
  const messages = [];
  t.rows.forEach(function (r, i) {
    const code = codes[i];
    if (!code) { links.push(['']); messages.push(['']); return; }
    const link = base + '?g=' + code;
    links.push([link]);
    const current = String(r[t.c.message] == null ? '' : r[t.c.message]);
    if (seen[code]) { messages.push([current]); return; }
    seen[code] = true;
    if (current.trim()) { messages.push([current]); return; }
    const names = t.rows
      .filter(function (_, j) { return codes[j] === code; })
      .map(function (x) { return String(x[t.c.name]).trim(); });
    messages.push([messageFor_(r[t.c.household], names, link)]);
  });
  if (!t.rows.length) return;
  sh.getRange(2, t.c.code + 1, t.rows.length, 1).setValues(codes.map(function (c) { return [c]; }));
  sh.getRange(2, t.c.link + 1, t.rows.length, 1).setValues(links);
  sh.getRange(2, t.c.message + 1, t.rows.length, 1).setValues(messages);
}

// ── Web app: the site calls these ──────────────────────────────────────────
/** GET ?code=ABC123 → the names on that link and any answers already saved. */
function doGet(e) {
  const code = normCode_(e && e.parameter && e.parameter.code);
  if (!code) return json_({ ok: true, service: 'rsvp' });
  const t = table_(sheet_());
  const guests = [];
  let note = '';
  let accom = '';
  t.rows.forEach(function (r) {
    if (normCode_(r[t.c.code]) !== code || !String(r[t.c.name]).trim()) return;
    const events = {};
    EVENTS.forEach(function (ev) { events[ev.id] = r[t.c[ev.id]] === 'Yes'; });
    guests.push({
      name: String(r[t.c.name]).trim(),
      attending: r[t.c.attending] === 'Yes' ? 'yes' : r[t.c.attending] === 'No' ? 'no' : '',
      events: events,
      perform: r[t.c.perform] === 'Yes',
      dietary: String(r[t.c.dietary] || ''),
      email: String(r[t.c.email] || ''),
    });
    if (r[t.c.note]) note = String(r[t.c.note]);
    if (r[t.c.accom]) accom = String(r[t.c.accom]);
  });
  if (!guests.length) return json_({ ok: false, error: 'not_found' });
  return json_({ ok: true, version: 2, guests: guests, note: note, accom: accom });
}

/** POST {code, guests:[{name, attending, events, perform, dietary?}], note} → saves answers. */
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  const code = normCode_(body && body.code);
  if (!code || !Array.isArray(body.guests)) return json_({ ok: false, error: 'bad_request' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const t = table_(sh);
    const note = clean_(body.note, 1000);
    const now = new Date();
    let found = 0;
    let saved = 0;
    t.rows.forEach(function (r, i) {
      if (normCode_(r[t.c.code]) !== code) return;
      const name = String(r[t.c.name]).trim();
      if (!name) return;
      found++;
      // Only names already on this link can be answered for: no one can add a guest.
      const g = body.guests.filter(function (x) { return x && String(x.name).trim() === name; })[0];
      if (!g || (g.attending !== 'yes' && g.attending !== 'no')) return;
      const yes = g.attending === 'yes';
      const row = i + 2;
      const set = function (key, value) { sh.getRange(row, t.c[key] + 1).setValue(value); };
      set('attending', yes ? 'Yes' : 'No');
      EVENTS.forEach(function (ev) { set(ev.id, yes ? (g.events && g.events[ev.id] === true ? 'Yes' : 'No') : ''); });
      if ('perform' in g) set('perform', yes && g.perform === true ? 'Yes' : '');
      if ('dietary' in g) set('dietary', clean_(g.dietary, 200));
      if ('email' in g) set('email', clean_(g.email, 200));
      if ('accom' in body) set('accom', clean_(body.accom, 500));
      set('note', note);
      set('updated', now);
      saved++;
    });
    if (!found) return json_({ ok: false, error: 'not_found' });
    if (!saved) return json_({ ok: false, error: 'nothing_saved' });
    return json_({ ok: true, saved: saved });
  } finally {
    lock.releaseLock();
  }
}

// ── Helpers ────────────────────────────────────────────────────────────────
function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET);
  if (!sh) throw new Error('No "' + SHEET + '" tab. Run setup first.');
  return sh;
}

/** Reads the Guests tab; c maps keys (and event ids) to 0-based column indexes. */
function table_(sh) {
  let values = sh.getDataRange().getValues();
  const header = values[0].map(function (v) { return String(v).trim(); });
  let added = false;
  ADDED_LATER.forEach(function (k) {
    if (header.indexOf(H[k]) < 0) {
      sh.getRange(1, header.length + 1).setValue(H[k]).setFontWeight('bold');
      header.push(H[k]);
      added = true;
    }
  });
  if (added) values = sh.getDataRange().getValues();
  const c = {};
  Object.keys(H).forEach(function (k) { c[k] = header.indexOf(H[k]); });
  EVENTS.forEach(function (ev) { c[ev.id] = header.indexOf(ev.header); });
  const missing = Object.keys(c).filter(function (k) { return c[k] < 0; });
  if (missing.length) throw new Error('Guests tab is missing columns: ' + missing.join(', '));
  return { c: c, rows: values.slice(1) };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function normCode_(v) {
  const s = String(v == null ? '' : v).trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
  return s.length >= 4 && s.length <= 12 ? s : '';
}

/** 6 characters without look-alikes (no 0/O, 1/I/L). */
function newCode_() {
  const alphabet = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  const hex = Utilities.getUuid().replace(/-/g, '');
  let out = '';
  for (let i = 0; i < 6; i++) out += alphabet[parseInt(hex.substr(i * 4, 4), 16) % alphabet.length];
  return out;
}

/** Trims, caps length, and stops text being read as a spreadsheet formula. */
function clean_(v, max) {
  let s = String(v == null ? '' : v).trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function firstName_(n) { return String(n).trim().split(/\s+/)[0]; }

function joinNames_(list) {
  if (list.length <= 1) return list[0] || '';
  return list.slice(0, -1).join(', ') + ' & ' + list[list.length - 1];
}

function colLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
