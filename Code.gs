// Ledger sync (Google Apps Script).
// Easiest way: in the Ledger app open Sync, tap "Copy script", then paste it into a new project at script.google.com.
// Then Deploy > New deployment > Web app > Who has access: Anyone > Deploy, and paste the Web app URL back into the app.
// The script creates a Google Sheet called "Ledger" in a Drive folder called "Ledger": one tab per month, newest first,
// tabs colored by year, and a Summary tab with each year as a collapsible group.

const TOKEN = '__TOKEN__';
const FOLDER = 'Ledger';
const FILE_NAME = '__NAME__';
const EMAIL = ''; // leave empty to email your own Google account (bill reminders)
const SHEET = 'Entries';
const SETTINGS = 'Settings';
const MON = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const YEAR_COLORS = ['#0B8F7A', '#D98A2B', '#4A7BD0', '#C4547A', '#8A6FD1', '#6B7A2E'];

function ss_() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SSID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  ss = SpreadsheetApp.create(FILE_NAME);
  props.setProperty('SSID', ss.getId());
  try { ss.getSheets()[0].setName('Summary'); } catch (e) {}
  try {
    const it = DriveApp.getFoldersByName(FOLDER);
    const f = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER);
    DriveApp.getFileById(ss.getId()).moveTo(f);
  } catch (e) {}
  return ss;
}

function entries_(ss) {
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(['id', 'date', 'type', 'category', 'amount', 'note', 'deleted', 'updated', 'sv', 'payment', 'goal']);
    sh.getRange('A:B').setNumberFormat('@');
    sh.getRange('F:F').setNumberFormat('@');
    sh.setFrozenRows(1);
    try { sh.hideSheet(); } catch (e) {}
  }
  return sh;
}

// Bills, budgets, goals, plans, favorites, categories and repeating entries live here as one JSON blob (last write wins).
function cfgSheet_(ss) {
  let sh = ss.getSheetByName(SETTINGS);
  if (!sh) {
    sh = ss.insertSheet(SETTINGS);
    sh.getRange('A1:A2').setValues([['updated'], ['data']]);
    sh.getRange('B1:B2').setNumberFormat('@');
    try { sh.hideSheet(); } catch (e) {}
  }
  return sh;
}

function owner_() {
  try { return Session.getEffectiveUser().getEmail(); } catch (e) { return ''; }
}

function fmtDate_(v, tz) {
  return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v);
}
function mk_(r) { return String(r[1]).slice(0, 7); }

function tot_(rows) {
  let inc = 0, sp = 0, sv = 0;
  rows.forEach(r => {
    const a = Number(r[4]);
    if (r[2] === 'income') inc += a; else if (Number(r[10])) sv += a; else sp += a;
  });
  return [inc, sp, sv, inc - sp - sv, rows.length];
}

function monthSheet_(ss, m, rows) {
  let sh = ss.getSheetByName(m);
  if (!sh) sh = ss.insertSheet(m);
  sh.clear();
  const y = Number(m.slice(0, 4)), mo = Number(m.slice(5));
  rows = rows.slice().sort((a, b) => String(a[1]).localeCompare(String(b[1])) || Number(a[0]) - Number(b[0]));
  const t = tot_(rows);
  const body = rows.map(r => {
    const a = Number(r[4]), inc = r[2] === 'income';
    return [r[1], inc ? 'Income' : (Number(r[10]) ? 'Savings' : 'Expense'), r[3], r[5], inc ? a : -a, r[9] || ''];
  });
  sh.getRange(1, 1).setValue(MON[mo - 1] + ' ' + y).setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1, 2, 5).setValues([['Income', 'Spent', 'Saved to goals', 'Left', 'Entries'], t]);
  sh.getRange(2, 1, 1, 5).setFontWeight('bold');
  sh.getRange(3, 1, 1, 4).setNumberFormat('#,##0.00');
  sh.getRange(5, 1, 1, 6).setValues([['Date', 'Type', 'Category', 'Note', 'Amount (RM)', 'Paid by']]).setFontWeight('bold').setBackground('#E8F3F0');
  if (body.length) {
    sh.getRange(6, 1, body.length, 1).setNumberFormat('@');
    sh.getRange(6, 1, body.length, 6).setValues(body);
    sh.getRange(6, 5, body.length, 1).setNumberFormat('#,##0.00;[Red]-#,##0.00');
  }
  try { sh.setFrozenRows(5); sh.setTabColor(YEAR_COLORS[y % YEAR_COLORS.length]); sh.setColumnWidth(1, 95); sh.setColumnWidth(4, 220); } catch (e) {}
}

function order_(ss) {
  const names = ss.getSheets().map(s => s.getName()).filter(n => /^\d{4}-\d{2}$/.test(n)).sort().reverse();
  names.forEach((n, i) => {
    try { ss.setActiveSheet(ss.getSheetByName(n)); ss.moveActiveSheet(i + 2); } catch (e) {}
  });
}

function summary_(ss, by) {
  let sh = ss.getSheetByName('Summary');
  if (!sh) sh = ss.insertSheet('Summary', 0);
  try { for (let i = 0; i < 2; i++) sh.getRange(1, 1, sh.getMaxRows(), 1).shiftRowGroupDepth(-1); } catch (e) {}
  sh.clear();
  const months = Object.keys(by).sort().reverse(), years = {};
  months.forEach(m => { (years[m.slice(0, 4)] = years[m.slice(0, 4)] || []).push(m); });
  const yrs = Object.keys(years).sort().reverse();
  const curYear = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy');
  const out = [['Year / Month', 'Income', 'Spent', 'Saved to goals', 'Left', 'Entries']];
  const yearRows = [], monthRows = [], groups = [];
  yrs.forEach(y => {
    const yt = [0, 0, 0, 0, 0], list = years[y].map(m => ({ m: m, t: tot_(by[m]) }));
    list.forEach(x => x.t.forEach((v, i) => { yt[i] += v; }));
    out.push([y].concat(yt));
    yearRows.push(out.length);
    const start = out.length + 1;
    list.forEach(x => { out.push([MON[Number(x.m.slice(5)) - 1]].concat(x.t)); monthRows.push({ row: out.length, m: x.m }); });
    if (list.length) groups.push({ start: start, count: list.length, open: y === curYear });
  });
  sh.getRange(1, 1, out.length, 6).setValues(out);
  sh.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#E8F3F0');
  yearRows.forEach(r => sh.getRange(r, 1, 1, 6).setFontWeight('bold').setBackground('#F3F4F1'));
  sh.getRange(2, 2, out.length, 4).setNumberFormat('#,##0.00');
  monthRows.forEach(x => {
    try { sh.getRange(x.row, 1).setFormula('=HYPERLINK("#gid=' + ss.getSheetByName(x.m).getSheetId() + '","' + out[x.row - 1][0] + '")'); } catch (e) {}
  });
  try { sh.setFrozenRows(1); sh.setColumnWidth(1, 150); } catch (e) {}
  groups.forEach(g => {
    try {
      sh.getRange(g.start, 1, g.count, 1).shiftRowGroupDepth(1);
      if (!g.open) sh.getRowGroup(g.start, 1).collapse();
    } catch (e) {}
  });
}

function rebuild_(ss, data, months) {
  const by = {};
  data.forEach(r => { if (!Number(r[6]) && /^\d{4}-\d{2}/.test(String(r[1]))) (by[mk_(r)] = by[mk_(r)] || []).push(r); });
  const todo = {};
  months.forEach(m => { todo[m] = 1; });
  Object.keys(by).forEach(m => { if (!ss.getSheetByName(m)) todo[m] = 1; });
  Object.keys(todo).forEach(m => monthSheet_(ss, m, by[m] || []));
  order_(ss);
  summary_(ss, by);
}

function doGet() {
  return ContentService.createTextOutput('Ledger sync is running');
}

function doPost(e) {
  const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: 'Bad request' }); }
  if (req.token !== TOKEN) return out({ ok: false, error: 'Wrong token' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = ss_(), sh = entries_(ss), now = Date.now(), tz = Session.getScriptTimeZone();
    if (req.peek) {
      // read-only check used when an app connects: lets it decide whether to pull, push or ask first
      const lr = sh.getLastRow();
      const flags = lr > 1 ? sh.getRange(2, 7, lr - 1, 1).getValues() : [];
      const pv = cfgSheet_(ss).getRange('B1').getValue();
      return out({ ok: true, peek: true, sid: ss.getId(), name: ss.getName(), owner: owner_(), entries: flags.filter(f => !Number(f[0])).length, hasCfg: (Number(pv) || 0) > 0 });
    }
    const last = sh.getLastRow();
    const data = last > 1 ? sh.getRange(2, 1, last - 1, 11).getValues() : [];
    data.forEach(r => { r[0] = String(r[0]); r[1] = fmtDate_(r[1], tz); });
    const idx = {}, touched = {};
    data.forEach((r, i) => { idx[r[0]] = i; });
    (req.rows || []).forEach(t => {
      const row = [String(t.id), t.date, t.type, t.cat, t.amt / 100, t.note || '', t.del ? 1 : 0, t.upd, now, t.pay || '', t.goal || 0];
      const i = idx[String(t.id)];
      if (i === undefined) { data.push(row); idx[String(t.id)] = data.length - 1; touched[mk_(row)] = 1; }
      else if (t.upd >= Number(data[i][7])) { touched[mk_(data[i])] = 1; data[i] = row; touched[mk_(row)] = 1; }
    });
    if (data.length) sh.getRange(2, 1, data.length, 11).setValues(data);
    const months = Object.keys(touched).filter(m => /^\d{4}-\d{2}$/.test(m));
    if (months.length || !ss.getSheetByName('Summary')) rebuild_(ss, data, months);
    const since = Number(req.since) || 0;
    const rows = data.filter(r => Number(r[8]) >= since).map(r => ({
      id: Number(r[0]), date: r[1], type: r[2], cat: r[3],
      amt: Math.round(Number(r[4]) * 100), note: String(r[5]),
      del: Number(r[6]) ? 1 : 0, upd: Number(r[7]), pay: String(r[9] || ''), goal: Number(r[10]) || 0
    }));

    const cs = cfgSheet_(ss), cv = cs.getRange('B1:B2').getValues(), cu = Number(cv[0][0]) || 0;
    let cfg = null, cfgVer = null;
    // A device may only overwrite the settings if it has seen the latest version (base). A blank or out-of-date device can never wipe them.
    const accept = req.cfg && (cu === 0 || (req.cfg.base === undefined ? req.cfg.upd > cu : Number(req.cfg.base) === cu));
    if (accept) {
      const json = JSON.stringify(req.cfg.data || {});
      if (json.length > 45000) return out({ ok: false, error: 'Settings too large' });
      cfgVer = Math.max(Date.now(), cu + 1);
      cs.getRange('B1:B2').setValues([[String(cfgVer)], [json]]);
    } else if (cu > 0 && (req.cfg || cu > (Number(req.cfgKnown) || 0))) {
      cfg = { upd: cu, data: JSON.parse(cv[1][0] || '{}') };
    }
    return out({ ok: true, rows: rows, cursor: now, cfg: cfg, cfgVer: cfgVer, sid: ss.getId(), url: ss.getUrl(), name: ss.getName(), owner: owner_() });
  } finally {
    lock.releaseLock();
  }
}

// Daily email for bills that are overdue or due within 3 days. Add a time-driven trigger for this function.
function sendBillReminders() {
  const raw = cfgSheet_(ss_()).getRange('B2').getValue();
  if (!raw || raw === 'data') return;
  const bills = (JSON.parse(raw).bills) || [];
  const tz = Session.getScriptTimeZone(), now = new Date();
  const k = Utilities.formatDate(now, tz, 'yyyy-MM');
  const td = Number(Utilities.formatDate(now, tz, 'd'));
  const dim = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  const lines = [];
  bills.forEach(b => {
    if (b.paid && b.paid[k]) return;
    if ((b.from && k < b.from) || (b.to && k > b.to)) return;
    const d = Math.min(b.day, dim) - td;
    if (d > 3) return;
    lines.push('- ' + b.name + ': RM ' + (b.amt / 100).toFixed(2) + (d < 0 ? ' (overdue)' : d === 0 ? ' (due today)' : ' (due in ' + d + ' days)'));
  });
  if (!lines.length) return;
  MailApp.sendEmail(EMAIL || Session.getEffectiveUser().getEmail(), 'Bills due soon', lines.join('\n'));
}
