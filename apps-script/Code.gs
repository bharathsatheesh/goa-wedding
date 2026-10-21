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
  if (/^anushri.s friends/.test(hh)) partner = 'Bharath';
  if (/^bharath.s friends/.test(hh)) partner = 'Anushri';
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
    .addItem('Fix links & messages', 'createLinks')
    .addItem('Set up sheet (first time only)', 'setup')
    .addToUi();
  try { if (PropertiesService.getDocumentProperties().getProperty('layout') !== 'formulas-v1') createLinks(); }
  catch (err) { /* a viewer without edit access opened the Sheet */ }
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

// ── Links & messages are formulas ─────────────────────────────────────────
// Every row's Link and Message cell holds the same formula, which reads its own row. Deleting,
// moving or sorting rows can never put a link or message on the wrong person, and nothing has
// to run for a rename to show up. Typing over a Message cell keeps your own wording for that
// invite; clearing it brings the standard message back.
function cols_(t) {
  return {
    hh: colLetter_(t.c.household + 1), name: colLetter_(t.c.name + 1), code: colLetter_(t.c.code + 1),
  };
}
function linkFormula_(t) {
  const L = cols_(t);
  return '=LET(cd,INDEX($' + L.code + ':$' + L.code + ',ROW()),IF(cd="","","' +
    SITE_URL.replace(/\/?$/, '/') + '?g="&cd))';
}
function messageFormula_(t) {
  const L = cols_(t);
  const col = function (x) { return '$' + x + ':$' + x; };
  const q = function (str) { return '"' + str.replace(/"/g, '""') + '"'; };
  const parts = MESSAGE_FAMILY.split('{link}');
  const fam = parts[0].split('{names}');
  const fr = MESSAGE_FRIENDS.split('{link}')[0].split('{names}');
  const frAfter = fr[1].split('{partner}');
  const tail = parts[1].replace(/\n/g, '');
  return '=LET(nm,INDEX(' + col(L.name) + ',ROW()),cd,INDEX(' + col(L.code) + ',ROW()),' +
    'IF(OR(nm="",cd=""),"",IF(MATCH(cd,' + col(L.code) + ',0)<>ROW(),"",' +
    'LET(n,REGEXREPLACE(TEXTJOIN(", ",TRUE,FILTER(' + col(L.name) + ',' + col(L.code) + '=cd)),", ([^,]*)$"," & $1"),' +
    'h,LOWER(INDEX(' + col(L.hh) + ',ROW())),' +
    'p,IF(REGEXMATCH(h,"^anushri.s friends"),"Bharath",IF(REGEXMATCH(h,"^bharath.s friends"),"Anushri","")),' +
    'IF(p="",' + q(fam[0]) + '&n&' + q(fam[1]) + ',' + q(fr[0]) + '&n&' + q(frAfter[0]) + '&p&' + q(frAfter[1]) + ')' +
    '&"' + SITE_URL.replace(/\/?$/, '/') + '?g="&cd&CHAR(10)&CHAR(10)&' + q(tail) + '))))';
}

/**
 * Menu "Fix links & messages" (also runs once by itself the first time the Sheet opens):
 * gives every named guest without a code a code (shared with their Household), puts the link
 * formula in every row, and the message formula in every Message cell that's empty or still
 * holds the standard wording. Hand-written messages are left exactly as they are.
 */
function createLinks() {
  const sh = sheet_();
  const t = table_(sh);
  const n = t.rows.length;
  if (!n) return;
  const codes = t.rows.map(function (r) { return normCode_(r[t.c.code]); });
  const hh = t.rows.map(function (r) { return String(r[t.c.household]).trim(); });
  const hk = hh.map(hhKey_);
  const names = t.rows.map(function (r) { return String(r[t.c.name]).trim(); });
  const taken = {};
  const byHousehold = {};
  codes.forEach(function (c, i) { if (c) { taken[c] = true; if (hk[i] && !byHousehold[hk[i]]) byHousehold[hk[i]] = c; } });
  codes.forEach(function (c, i) {
    if (!names[i] || c) return;
    let code = hk[i] && byHousehold[hk[i]];
    if (!code) {
      do { code = newCode_(); } while (taken[code]);
      taken[code] = true;
      if (hk[i]) byHousehold[hk[i]] = code;
    }
    codes[i] = code;
    sh.getRange(i + 2, t.c.code + 1).setValue(code);
  });

  const base = SITE_URL.replace(/\/?$/, '/');
  const linkF = linkFormula_(t), msgF = messageFormula_(t);
  const msgFormulas = sh.getRange(2, t.c.message + 1, n, 1).getFormulas();
  const firstRow = {};
  codes.forEach(function (c, i) { if (c && names[i] && !(c in firstRow)) firstRow[c] = i; });
  const messages = t.rows.map(function (r, i) {
    const current = String(r[t.c.message] == null ? '' : r[t.c.message]);
    if (msgFormulas[i][0] || !current.trim()) return [msgF];
    const c = codes[i];
    if (!c || firstRow[c] !== i) return [{ keep: true }];
    const group = names.filter(function (_, j) { return codes[j] === c && names[j]; });
    return current === messageFor_(hh[i], group, base + '?g=' + c) ? [msgF] : [{ keep: true }];
  });
  sh.getRange(2, t.c.link + 1, n, 1).setFormulas(t.rows.map(function () { return [linkF]; }));
  messages.forEach(function (m, i) { if (!m[0].keep) sh.getRange(i + 2, t.c.message + 1).setFormula(m[0]); });
  PropertiesService.getDocumentProperties().setProperty('layout', 'formulas-v1');
}

// ── Automatic upkeep: runs whenever someone edits the Guests tab by hand ────
/**
 * Type a new Name     → they get a code (their Household's, or a new one) plus the link and
 *                       message formulas.
 * Change a Household  → that person moves to that household's link, or gets their own.
 * Clear a Message     → the standard message comes back.
 * Only the Code cell of the edited row is ever written, after re-checking that row.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEET || e.range.getLastRow() < 2) return;
  const t = table_(sh);
  const c0 = e.range.getColumn() - 1, c1 = c0 + e.range.getNumColumns() - 1;
  const hits = function (k) { return t.c[k] >= c0 && t.c[k] <= c1; };
  if (!hits('household') && !hits('name') && !hits('message')) return;
  const single = e.range.getNumRows() === 1 && e.range.getNumColumns() === 1;
  const notes = [];
  for (let row = Math.max(2, e.range.getRow()); row <= e.range.getLastRow(); row++) {
    const note = fixRow_(sh, t, row, single && c0 === t.c.household ? String(e.oldValue == null ? '' : e.oldValue) : null);
    if (note) notes.push(note);
  }
  if (notes.length) SpreadsheetApp.getActiveSpreadsheet().toast(notes.join('\n'), 'Wedding RSVP', 8);
}

/** Brings one row up to date. movedFrom = old Household if that cell was just changed by hand. */
function fixRow_(sh, t, row, movedFrom) {
  const i = row - 2;
  const r = t.rows[i];
  if (!r) return '';
  const name = String(r[t.c.name]).trim();
  const hh = String(r[t.c.household]).trim();
  const code = normCode_(r[t.c.code]);
  let note = '';

  if (name) {
    let want = code;
    const others = function (pred) { for (let j = 0; j < t.rows.length; j++) if (j !== i && pred(t.rows[j])) return t.rows[j]; return null; };
    const sameHh = hh ? others(function (x) { return hhKey_(x[t.c.household]) === hhKey_(hh) && normCode_(x[t.c.code]); }) : null;
    if (!code) {
      want = sameHh ? normCode_(sameHh[t.c.code]) : freshCode_(t);
    } else if (movedFrom !== null && hhKey_(movedFrom) !== hhKey_(hh)) {
      if (sameHh) want = normCode_(sameHh[t.c.code]);
      else if (others(function (x) { return normCode_(x[t.c.code]) === code; })) want = freshCode_(t);
    }
    if (want !== code) {
      // Re-read the row just before writing, so a row moved or deleted meanwhile is never overwritten.
      const now = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
      if (String(now[t.c.name]).trim() !== name || String(now[t.c.household]).trim() !== hh ||
          normCode_(now[t.c.code]) !== code) return '';
      sh.getRange(row, t.c.code + 1).setValue(want);
      r[t.c.code] = want;   // so the next row in a paste sees it
      if (code) {
        const sentCol = t.header.indexOf('Invite sent');
        const sent = sentCol >= 0 && r[sentCol] === true;
        note = sameHh ? name + ' now shares a link with ' + hh + '.' : name + ' now has their own link.';
        if (sent) note += ' Their invite was already sent with the old link, so send them the new one.';
        const msgNow = sh.getRange(row, t.c.message + 1);
        if (!msgNow.getFormula() && String(msgNow.getValue()).trim()) {
          note += ' Their message was written by hand and still has the old link: update it, or clear it to get the standard one.';
        }
      }
    }
  }
  // Formulas are the same in every row, so writing them can never land on the wrong person.
  const linkCell = sh.getRange(row, t.c.link + 1), msgCell = sh.getRange(row, t.c.message + 1);
  if (!linkCell.getFormula()) linkCell.setFormula(linkFormula_(t));
  if (!msgCell.getFormula() && !String(msgCell.getValue()).trim()) msgCell.setFormula(messageFormula_(t));
  return note;
}

/** Households match ignoring capitals, extra spaces and curly vs straight apostrophes. */
function hhKey_(v) {
  return String(v == null ? '' : v).replace(/[\u2018\u2019`]/g, "'").replace(/\s+/g, ' ').trim().toLowerCase();
}

function freshCode_(t) {
  const taken = {};
  t.rows.forEach(function (r) { const c = normCode_(r[t.c.code]); if (c) taken[c] = true; });
  let c;
  do { c = newCode_(); } while (taken[c]);
  return c;
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
  return { c: c, rows: values.slice(1), header: header };
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
