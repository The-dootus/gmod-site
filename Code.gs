// V&B Event Ledger backend. Paste into Extensions > Apps Script in your Google Sheet.
// 1) Set the three passwords below  2) Run setup() once  3) DELETE the passwords from this file  4) Deploy as web app.
const USERS = [
  { name: 'viewer',  role: 'viewer', password: 'CHANGE-ME-12+chars' },
  { name: 'you',     role: 'editor', password: 'CHANGE-ME-12+chars' },
  { name: 'manager', role: 'editor', password: 'CHANGE-ME-12+chars' }
];
const SHEETS = {
  Entries: ['id', 'date', 'venue', 'type', 'amount', 'reason', 'recognized', 'note', 'voided', 'by', 'at'],
  Payouts: ['id', 'date', 'amount', 'note', 'voided', 'by', 'at'],
  Tax: ['id', 'name', 'amount_per_period', 'start', 'end', 'by', 'at'],
  Venues: ['venue', 'split_pct'],
  Settings: ['key', 'value'],
  Users: ['name', 'role', 'salt', 'hash'],
  Audit: ['at', 'user', 'action', 'detail']
};
const TEXT_COLS = { Entries: ['B'], Payouts: ['B'], Tax: ['D', 'E'], Settings: ['B'] };

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const n in SHEETS) {
    const s = ss.getSheetByName(n) || ss.insertSheet(n);
    (TEXT_COLS[n] || []).forEach(c => s.getRange(c + ':' + c).setNumberFormat('@'));
    if (!s.getLastRow()) {
      s.appendRow(SHEETS[n]); s.setFrozenRows(1);
      s.getRange(1, 1, 1, SHEETS[n].length).setFontWeight('bold').setBackground('#1f6b4f').setFontColor('#ffffff');
    }
  }
  const add = (n, row) => ss.getSheetByName(n).appendRow(row);
  if (ss.getSheetByName('Venues').getLastRow() < 2) { add('Venues', ['Motel', 30]); add('Venues', ['Townhouse 4', 0]); }
  if (ss.getSheetByName('Settings').getLastRow() < 2) { add('Settings', ['tax_period_days', '14']); add('Settings', ['tax_anchor', '2026-11-08']); }
  if (ss.getSheetByName('Tax').getLastRow() < 2) add('Tax', [1, 'Base license', 20000, '2026-10-26', '', 'setup', new Date().toISOString()]);
  const us = ss.getSheetByName('Users');
  USERS.forEach(u => {
    if (u.password.length < 12 || u.password.indexOf('CHANGE-ME') === 0) throw Error('Set a real 12+ character password for ' + u.name);
    const salt = Utilities.getUuid(), vals = us.getDataRange().getValues(), i = vals.findIndex(r => r[0] === u.name);
    const row = [u.name, u.role, salt, hash(salt, u.password)];
    if (i > 0) us.getRange(i + 1, 1, 1, 4).setValues([row]); else us.appendRow(row);
  });
  Logger.log('Setup done. Now delete the passwords from the USERS list at the top of this file.');
}

function hash(salt, pw) {
  let h = salt + pw;
  for (let i = 0; i < 2000; i++) h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt + pw));
  return h;
}

function doPost(e) {
  let out;
  try { out = handle(JSON.parse(e.postData.contents)); } catch (x) { out = { error: x.message }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function handle(r) {
  const cache = CacheService.getScriptCache();
  if (r.action === 'login') {
    const uname = String(r.username || '').toLowerCase(), key = 'f_' + uname, f = +cache.get(key) || 0;
    if (f >= 5) throw Error('Too many attempts. Try again in 10 minutes.');
    const u = rows('Users').find(x => String(x.name).toLowerCase() === uname);
    if (!u || hash(u.salt, String(r.password || '')) !== u.hash) { cache.put(key, String(f + 1), 600); throw Error('Wrong username or password'); }
    cache.remove(key);
    const tok = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
    cache.put('s_' + tok, JSON.stringify({ name: u.name, role: u.role }), 21600);
    audit(u.name, 'login', '');
    return { token: tok };
  }
  const raw = cache.get('s_' + r.token);
  if (!raw) throw Error('SESSION');
  const me = JSON.parse(raw);
  if (r.action === 'logout') { cache.remove('s_' + r.token); return { ok: 1 }; }
  if (r.action !== 'data') {
    if (me.role !== 'editor') throw Error('Viewer accounts are read-only');
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try { write(r, me); } finally { lock.releaseLock(); }
  }
  const entries = rows('Entries'), payouts = rows('Payouts'), tax = rows('Tax'), venues = rows('Venues');
  const st = {}; rows('Settings').forEach(x => st[x.key] = x.value);
  return { me, venues, entries, payouts, tax, report: compute(entries, payouts, tax, venues, st, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')),
           audit: rows('Audit').reverse().slice(0, 300) };
}

function ds(v) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v); }
function rows(name) {
  const v = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).getDataRange().getValues(), h = v.shift();
  return v.filter(r => r.some(c => c !== '')).map(r => { const o = {}; h.forEach((k, i) => o[k] = (k === 'date' || k === 'start' || k === 'end') && r[i] !== '' ? ds(r[i]) : r[i]); return o; });
}
function audit(user, action, detail) { SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Audit').appendRow([new Date().toISOString(), user, action, detail]); }
function nextId(name) { return rows(name).reduce((m, r) => Math.max(m, +r.id || 0), 0) + 1; }
function setCell(name, id, col, val) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name), v = s.getDataRange().getValues(), c = v[0].indexOf(col), i = v.findIndex((r, k) => k > 0 && +r[0] === +id);
  if (i < 1) throw Error('No such record'); s.getRange(i + 1, c + 1).setValue(val);
}
const isDate = x => /^\d{4}-\d{2}-\d{2}$/.test(String(x));

function write(r, me) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), now = new Date().toISOString(), a = r.entry || {};
  if (r.action === 'addEntry') {
    const venues = rows('Venues').map(v => v.venue), exp = a.type === 'Expense';
    if (!isDate(a.date) || venues.indexOf(a.venue) < 0 || ['Revenue', 'Expense'].indexOf(a.type) < 0) throw Error('Check the date, venue and type');
    if (!(Number(a.amount) > 0) || Math.round(a.amount * 100) !== a.amount * 100) throw Error('Amount must be positive with at most 2 decimals');
    if (exp && !String(a.reason || '').trim()) throw Error('Every expense needs a reason');
    const id = nextId('Entries');
    ss.getSheetByName('Entries').appendRow([id, a.date, a.venue, a.type, Number(a.amount), String(a.reason || '').slice(0, 120), exp ? !!a.recognized : '', String(a.note || '').slice(0, 500), false, me.name, now]);
    audit(me.name, 'entry.add', JSON.stringify(Object.assign({ id }, a)));
  } else if (r.action === 'setRecognized') {
    setCell('Entries', r.id, 'recognized', !!r.value); audit(me.name, 'entry.recognized', JSON.stringify({ id: r.id, value: !!r.value }));
  } else if (r.action === 'voidEntry') { setCell('Entries', r.id, 'voided', true); audit(me.name, 'entry.void', JSON.stringify({ id: r.id }));
  } else if (r.action === 'addPayout') {
    if (!isDate(r.date) || !(Number(r.amount) > 0) || Math.round(r.amount * 100) !== r.amount * 100) throw Error('Check the date and amount');
    const id = nextId('Payouts'); ss.getSheetByName('Payouts').appendRow([id, r.date, Number(r.amount), String(r.note || '').slice(0, 300), false, me.name, now]);
    audit(me.name, 'payout.add', JSON.stringify({ id, date: r.date, amount: r.amount }));
  } else if (r.action === 'voidPayout') { setCell('Payouts', r.id, 'voided', true); audit(me.name, 'payout.void', JSON.stringify({ id: r.id }));
  } else if (r.action === 'addTax') {
    if (!isDate(r.start) || !(Number(r.amount) > 0) || !String(r.name || '').trim()) throw Error('Tax item needs a name, amount and start date');
    const id = nextId('Tax'); ss.getSheetByName('Tax').appendRow([id, String(r.name).slice(0, 80), Number(r.amount), r.start, '', me.name, now]);
    audit(me.name, 'tax.add', JSON.stringify({ id, name: r.name, amount: r.amount, start: r.start }));
  } else if (r.action === 'endTax') {
    if (!isDate(r.end)) throw Error('End date must be YYYY-MM-DD'); setCell('Tax', r.id, 'end', r.end); audit(me.name, 'tax.end', JSON.stringify({ id: r.id, end: r.end }));
  } else throw Error('Unknown action');
}

// ---- profit math (all in cents) ----
const cents = v => Math.round(Number(v) * 100);
const dnum = s => { const p = String(s).split('-'); return Date.UTC(+p[0], +p[1] - 1, +p[2]) / 86400000; };
const dstr = n => new Date(n * 86400000).toISOString().slice(0, 10);
const truthy = v => v === true || String(v).toUpperCase() === 'TRUE';

function compute(entries, payouts, tax, venues, st, today) {
  const split = {}; venues.forEach(v => split[v.venue] = Number(v.split_pct) || 0);
  const days = Number(st.tax_period_days) || 14, anchor = dnum(st.tax_anchor) - days + 1, E = {};   // tax_anchor = a day tax was paid (a weekend); each period ends on a payment day
  entries.filter(e => !truthy(e.voided)).forEach(e => {
    const o = E[e.venue + '|' + e.date] || (E[e.venue + '|' + e.date] = { venue: e.venue, date: e.date, revenue: 0, costs: 0, pending: 0, tax: 0 }), c = cents(e.amount);
    if (e.type === 'Revenue') o.revenue += c; else if (truthy(e.recognized)) o.costs += c; else o.pending += c;
  });
  const evs = Object.values(E), byP = {};
  evs.forEach(o => { o.p = Math.floor((dnum(o.date) - anchor) / days); (byP[o.p] = byP[o.p] || []).push(o); });
  const lo = Math.min(0, ...evs.map(o => o.p)), hi = Math.max(Math.floor((dnum(today) - anchor) / days), ...evs.map(o => o.p)), periods = [];
  for (let p = lo; p <= hi; p++) {
    const s = anchor + p * days, en = s + days - 1;
    const pool = tax.filter(t => dnum(t.start) <= en && (!t.end || dnum(t.end) >= s)).reduce((a, t) => a + cents(t.amount_per_period), 0);
    // tax is carried only by venues under the split agreement (the Motel)
    const g = (byP[p] || []).filter(o => split[o.venue] > 0), tot = g.reduce((a, o) => a + o.revenue, 0); let used = 0;
    if (tot > 0) {
      g.forEach(o => { o.tax = Math.floor(pool * o.revenue / tot); used += o.tax; });
      g.slice().sort((a, b) => b.revenue - a.revenue)[0].tax += pool - used; used = pool;   // remainder cents to the biggest night
    }
    periods.push({ start: dstr(s), end: dstr(en), due: pool, allocated: used, unallocated: pool - used });
  }
  let owed = 0, T = { revenue: 0, costs: 0, pending: 0, tax: 0 };
  evs.forEach(o => {
    o.net = o.revenue - o.costs - o.tax; o.loss = o.net < 0; o.pct = split[o.venue] || 0;
    o.share = o.net > 0 ? Math.round(o.net * o.pct / 100) : 0; owed += o.share;
    T.revenue += o.revenue; T.costs += o.costs; T.pending += o.pending; T.tax += o.tax;
  });
  const paid = payouts.filter(p => !truthy(p.voided)).reduce((a, p) => a + cents(p.amount), 0);
  T.net = T.revenue - T.costs - T.tax; T.earned = owed; T.paid = paid; T.balance = owed - paid;
  return { events: evs.sort((a, b) => b.date.localeCompare(a.date)), periods: periods.reverse(), totals: T };
}
