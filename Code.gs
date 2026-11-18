/**
 * PCash Vs IP – Stations. MRS petty cash for Head Office and the stations
 * (Carletonville, Evander, Welkom, Steelpoort) in one app.
 * Google Apps Script backend:  phone app (GitHub Pages) --POST JSON--> this web app --> Google Sheet + Drive.
 *
 * Built from the single-user Head Office app (PCash Vs IP v1.0). Same money rules, plus:
 *   - Stations sheet: one row per station, with its own expense categories (Head Office keeps PC / CF).
 *   - Users sheet: one row per person, with their own PIN, station (or ALL) and role.
 *     MANAGER sees and captures only their own station. ADMIN (finance) can open every station.
 *   - Every entry and month row carries a Station code. Vouchers are numbered per station.
 *   - Slip photos are filed per station: Petty Cash Stations / <Station> / <yyyy-mm>.
 *
 * ONE-TIME SETUP (from the Apps Script editor, signed in as the MRS apps account):
 *   1. Paste this file into a new Apps Script project.
 *   2. Select the function "setup" and press Run. Approve the permissions.
 *   3. Open the Execution log – it shows the Sheet link, Drive folder and the first admin PIN.
 *   4. Deploy > New deployment > Web app: Execute as "Me", Who has access "Anyone".
 *   5. Put the /exec URL into config.json in the GitHub repo.
 *   6. Add people on the Users sheet (Name, PIN, Station code or ALL, Role MANAGER or ADMIN, Active Y).
 *
 * Money rules (how every entry moves the two pots):
 *   SLIP  Card      -> bank  - amount   (counts in the month it shows on the statement: ReconMonth)
 *   SLIP  Cash      -> cash  - amount
 *   SLIP  EFT       -> no pot (paid from the main MRS account; kept for cost only)
 *   SLIP  Own card  -> no pot until reimbursed (reimbursement is a Cash out entry)
 *   IN    Cash exchange (EFT in) -> bank + amount, cash - amount
 *   IN    Cash exchange (cash in) -> cash + amount, bank - amount
 *   IN    others    -> bank or cash + amount (where it landed)
 *   OUT   all       -> cash - amount (or bank - amount for bank fees / bank-side outs)
 */

const APP_NAME = 'PCash Vs IP';
const APP_FULL = 'PCash Vs IP – Stations';

const ENTRY_HEADERS = [
  'ID', 'Created', 'Date', 'Kind', 'Type', 'Name', 'Voucher', 'Line', 'Alloc',
  'Amount', 'VAT', 'Side', 'Direction', 'ReconMonth', 'OnStatement', 'LinkID',
  'Status', 'Note', 'FileId', 'FileUrl', 'Station', 'By'
];
const MONTH_HEADERS = ['Station', 'Month', 'OpenBank', 'OpenCash', 'StmtClose', 'CashCount', 'PdfUrl', 'Updated'];
const STATION_HEADERS = ['Code', 'Name', 'Categories', 'Active'];
const USER_HEADERS = ['Name', 'PIN', 'Station', 'Role', 'Active'];
const TEXT_COLS = ['ID', 'Created', 'Date', 'Kind', 'Type', 'Name', 'Voucher', 'Line', 'Alloc',
  'Side', 'Direction', 'ReconMonth', 'OnStatement', 'LinkID', 'Status', 'Note', 'FileId', 'FileUrl',
  'Station', 'By', 'Month', 'Code', 'Categories', 'Active', 'PIN', 'Role', 'PdfUrl', 'Updated'];

const DEFAULT_STATIONS = [
  ['HO', 'Head Office', 'PC, CF', 'Y'],
  ['CTV', 'Carletonville', 'Groceries, Cleaning, Water, Fuel, Maintenance, Stationery, Other', 'Y'],
  ['EVD', 'Evander', 'Groceries, Cleaning, Water, Fuel, Maintenance, Stationery, Other', 'Y'],
  ['WLK', 'Welkom', 'Groceries, Cleaning, Water, Fuel, Maintenance, Stationery, Other', 'Y'],
  ['STP', 'Steelpoort', 'Groceries, Cleaning, Water, Fuel, Maintenance, Stationery, Other', 'Y']
];

const KINDS = ['SLIP', 'IN', 'OUT'];
const IN_TYPES = ['KITCHEN_SALES', 'STORE_SALES', 'CASH_EXCHANGE', 'REPAYMENT', 'TRANSFER', 'OTHER_IN'];
const OUT_TYPES = ['STAFF_WITHDRAWAL', 'CASH_DRAW', 'REIMBURSE', 'BANK_FEE', 'OTHER_OUT'];
const SLIP_SIDES = ['CARD', 'CASH', 'EFT', 'OWN'];
const EDITABLE = ['Date', 'Name', 'Alloc', 'Amount', 'VAT', 'Side', 'Direction', 'ReconMonth',
  'OnStatement', 'LinkID', 'Note', 'Type'];

/* ------------------------------------------------------------------ setup */

function setup() {
  const props = PropertiesService.getScriptProperties();
  let ss;
  const sheetId = props.getProperty('SHEET_ID');
  if (sheetId) {
    ss = SpreadsheetApp.openById(sheetId);
  } else {
    ss = SpreadsheetApp.create(APP_FULL);
    props.setProperty('SHEET_ID', ss.getId());
  }
  ensureSheet_(ss, 'Entries', ENTRY_HEADERS, ['Amount', 'VAT']);
  ensureSheet_(ss, 'Months', MONTH_HEADERS, ['OpenBank', 'OpenCash', 'StmtClose', 'CashCount']);
  const st = ensureSheet_(ss, 'Stations', STATION_HEADERS, []);
  if (st.getLastRow() < 2) st.getRange(2, 1, DEFAULT_STATIONS.length, STATION_HEADERS.length).setValues(DEFAULT_STATIONS);
  const us = ensureSheet_(ss, 'Users', USER_HEADERS, []);
  let firstPin = '';
  if (us.getLastRow() < 2) {
    firstPin = newPin_([]);
    us.getRange(2, 1, 1, USER_HEADERS.length).setValues([['Bertus', firstPin, 'ALL', 'ADMIN', 'Y']]);
  }
  ensureSheet_(ss, 'Config', ['Key', 'Value'], []);
  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1) ss.deleteSheet(sheet1);

  let root;
  const rootId = props.getProperty('ROOT_FOLDER_ID');
  if (rootId) {
    root = DriveApp.getFolderById(rootId);
  } else {
    root = DriveApp.createFolder('Petty Cash Stations');
    props.setProperty('ROOT_FOLDER_ID', root.getId());
    DriveApp.getFileById(ss.getId()).moveTo(root);
  }

  Logger.log('Sheet:  ' + ss.getUrl());
  Logger.log('Folder: ' + root.getUrl());
  if (firstPin) Logger.log('First admin user "Bertus", PIN ' + firstPin + '  (change it in the app under Settings)');
  checkUsers();
}

/** Run any time after editing the Users sheet: lists problems in the Execution log. */
function checkUsers() {
  const ss = sheet_();
  const codes = stations_(ss).map(function (s) { return s.Code; });
  const seen = {};
  readTable_(ss.getSheetByName('Users')).forEach(function (u, i) {
    const row = i + 2;
    if (!/^\d{4,8}$/.test(u.PIN)) Logger.log('Row ' + row + ' (' + u.Name + '): PIN must be 4 to 8 digits');
    if (seen[u.PIN]) Logger.log('Row ' + row + ' (' + u.Name + '): same PIN as ' + seen[u.PIN] + ' – every person needs their own PIN');
    seen[u.PIN] = u.Name;
    if (u.Station !== 'ALL' && codes.indexOf(u.Station) < 0) Logger.log('Row ' + row + ' (' + u.Name + '): station "' + u.Station + '" is not on the Stations sheet');
    if (['MANAGER', 'ADMIN'].indexOf(u.Role) < 0) Logger.log('Row ' + row + ' (' + u.Name + '): role must be MANAGER or ADMIN');
  });
  Logger.log('Users checked.');
}

function ensureSheet_(ss, name, headers, moneyCols) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const first = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  if (first.join('') === '') {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  headers.forEach(function (h, i) {
    const col = sh.getRange(1, i + 1, sh.getMaxRows(), 1);
    if (TEXT_COLS.indexOf(h) >= 0) col.setNumberFormat('@');
    if (moneyCols.indexOf(h) >= 0) col.setNumberFormat('0.00');
  });
  sh.getRange(1, 1, 1, headers.length).setNumberFormat('@');
  return sh;
}

/* ------------------------------------------------------------------ http */

function doGet() {
  return ContentService.createTextOutput(APP_FULL + ' API is running.');
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const ss = sheet_();
    const user = auth_(ss, req.pin);
    const fn = ACTIONS[req.action];
    if (!fn) throw new Error('Unknown action: ' + req.action);
    const ctx = { ss: ss, user: user, stations: stations_(ss) };
    if (req.action !== 'changePin') ctx.station = pickStation_(ctx, req.station);
    out = { ok: true, data: fn(ctx, req) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

const ACTIONS = {
  ping: function () { return { app: APP_FULL, time: new Date().toISOString() }; },
  load: actLoad_,
  saveSlip: actSaveSlip_,
  saveIn: function (ctx, req) { return actSaveMove_(ctx, req, 'IN'); },
  saveOut: function (ctx, req) { return actSaveMove_(ctx, req, 'OUT'); },
  update: actUpdate_,
  bulkUpdate: actBulkUpdate_,
  voidEntry: actVoid_,
  setMonth: actSetMonth_,
  buildPdf: actBuildPdf_,
  changePin: actChangePin_
};

/* ------------------------------------------------------------------ people and stations */

function sheet_() {
  return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
}

function auth_(ss, pin) {
  const p = String(pin || '').trim();
  const u = p ? readTable_(ss.getSheetByName('Users')).filter(function (x) {
    return x.PIN === p && x.Active !== 'N' && x.Name;
  })[0] : null;
  if (!u) {
    Utilities.sleep(800); // slow down guessing
    throw new Error('Wrong PIN');
  }
  return { name: u.Name, station: u.Station || '', role: u.Role === 'ADMIN' ? 'ADMIN' : 'MANAGER' };
}

function stations_(ss) {
  return readTable_(ss.getSheetByName('Stations')).filter(function (s) {
    return s.Code && s.Active !== 'N';
  }).map(function (s) {
    return {
      Code: s.Code, Name: s.Name || s.Code,
      Categories: String(s.Categories || '').split(',').map(function (c) { return c.trim(); }).filter(String)
    };
  });
}

/** Which station this request works on. Managers are always held to their own station. */
function pickStation_(ctx, wanted) {
  const codes = ctx.stations.map(function (s) { return s.Code; });
  const u = ctx.user;
  let code;
  if (u.station !== 'ALL') {
    if (wanted && wanted !== u.station) throw new Error('You can only open ' + u.station);
    code = u.station;
  } else {
    code = wanted || codes[0];
  }
  const st = ctx.stations.filter(function (s) { return s.Code === code; })[0];
  if (!st) throw new Error('Unknown station: ' + code);
  return st;
}

function checkAlloc_(ctx, alloc, required) {
  if (!alloc) { if (required) throw new Error('Pick a category'); return; }
  if (ctx.station.Categories.indexOf(alloc) < 0) {
    throw new Error('"' + alloc + '" is not a category for ' + ctx.station.Name);
  }
}

/* ------------------------------------------------------------------ actions */

function actLoad_(ctx) {
  const code = ctx.station.Code;
  const visible = ctx.user.station === 'ALL' ? ctx.stations : ctx.stations.filter(function (s) { return s.Code === code; });
  return {
    user: { name: ctx.user.name, role: ctx.user.role, station: ctx.user.station },
    stations: visible,
    station: code,
    entries: entriesFor_(ctx.ss, code),
    months: monthsFor_(ctx.ss, code),
    nextVoucher: getConfig_(ctx.ss, 'NextVoucher_' + code) || '1'
  };
}

/**
 * req: { date, vendor, side, note, image (base64 jpeg, optional),
 *        lines: [{ alloc, amount, vat }] (1 or 2 lines) }
 */
function actSaveSlip_(ctx, req) {
  const ss = ctx.ss, code = ctx.station.Code;
  checkDate_(req.date);
  if (SLIP_SIDES.indexOf(req.side) < 0) throw new Error('Pick how it was paid');
  const lines = (req.lines || []).filter(function (l) { return Number(l.amount) > 0; });
  if (!lines.length || lines.length > 2) throw new Error('A slip needs 1 or 2 lines');
  lines.forEach(function (l) {
    checkAlloc_(ctx, l.alloc, true);
    if (Number(l.vat || 0) < 0 || Number(l.vat || 0) >= Number(l.amount)) throw new Error('VAT looks wrong');
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const voucher = nextVoucher_(ss, code);
    const total = lines.reduce(function (s, l) { return s + Number(l.amount); }, 0);
    const file = req.image ? saveImage_(ctx.station, req.image, req.date,
      (req.vendor || 'Slip') + ' ' + total.toFixed(2) + ' V' + voucher) : null;
    const created = new Date().toISOString();
    const rows = lines.map(function (l, i) {
      return rowFrom_({
        ID: Utilities.getUuid(), Created: created, Date: req.date, Kind: 'SLIP', Type: 'PURCHASE',
        Name: clean_(req.vendor), Voucher: String(voucher), Line: String(i + 1), Alloc: l.alloc,
        Amount: round2_(l.amount), VAT: round2_(l.vat || 0), Side: req.side, Direction: '',
        ReconMonth: req.date.slice(0, 7), OnStatement: '', LinkID: '',
        Status: req.side === 'OWN' ? 'OPEN' : 'OK', Note: clean_(req.note),
        FileId: file ? file.id : '', FileUrl: file ? file.url : '',
        Station: code, By: ctx.user.name
      });
    });
    appendRows_(ss.getSheetByName('Entries'), rows);
    return { voucher: voucher, entries: rows.map(objFromRow_) };
  } finally {
    lock.releaseLock();
  }
}

/**
 * req: { date, type, name, amount, side:'BANK'|'CASH', direction, alloc, note, linkId, image }
 */
function actSaveMove_(ctx, req, kind) {
  const ss = ctx.ss, code = ctx.station.Code;
  checkDate_(req.date);
  const types = kind === 'IN' ? IN_TYPES : OUT_TYPES;
  if (types.indexOf(req.type) < 0) throw new Error('Pick a type');
  if (!(Number(req.amount) > 0) && !(req.type === 'KITCHEN_SALES' && Number(req.amount) === 0 && /swipes/i.test(req.note || ''))) throw new Error('Amount must be more than 0');
  checkAlloc_(ctx, req.alloc, false);
  if (req.linkId && !entriesFor_(ss, code).some(function (x) { return x.ID === req.linkId || 'V' + x.Voucher === req.linkId; })) {
    throw new Error('Linked entry not found at ' + ctx.station.Name);
  }
  const side = req.side === 'BANK' ? 'BANK' : 'CASH';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const voucher = kind === 'OUT' ? String(nextVoucher_(ss, code)) : '';
    const file = req.image ? saveImage_(ctx.station, req.image, req.date,
      label_(req.type) + ' ' + (req.name || '') + ' ' + round2_(req.amount).toFixed(2) + (voucher ? ' V' + voucher : '')) : null;
    const row = rowFrom_({
      ID: Utilities.getUuid(), Created: new Date().toISOString(), Date: req.date, Kind: kind,
      Type: req.type, Name: clean_(req.name), Voucher: voucher, Line: '', Alloc: clean_(req.alloc || ''),
      Amount: round2_(req.amount), VAT: 0, Side: side,
      Direction: req.type === 'CASH_EXCHANGE' ? (req.direction === 'CASH_IN' ? 'CASH_IN' : 'EFT_IN') : '',
      ReconMonth: req.date.slice(0, 7), OnStatement: '', LinkID: clean_(req.linkId || ''),
      Status: req.type === 'STAFF_WITHDRAWAL' ? 'OPEN' : 'OK', Note: clean_(req.note),
      FileId: file ? file.id : '', FileUrl: file ? file.url : '',
      Station: code, By: ctx.user.name
    });
    appendRows_(ss.getSheetByName('Entries'), [row]);
    if (req.linkId) refreshOpenStatus_(ss, code);
    return { entries: [objFromRow_(row)] };
  } finally {
    lock.releaseLock();
  }
}

function actUpdate_(ctx, req) {
  return actBulkUpdate_(ctx, { updates: [{ id: req.id, fields: req.fields }] });
}

function actBulkUpdate_(ctx, req) {
  const ss = ctx.ss, code = ctx.station.Code;
  const sh = ss.getSheetByName('Entries');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ids = idColumn_(sh), stations = stationColumn_(sh);
    (req.updates || []).forEach(function (u) {
      const r = ids.indexOf(u.id);
      if (r < 0 || stations[r] !== code) throw new Error('Entry not found');
      Object.keys(u.fields || {}).forEach(function (k) {
        if (EDITABLE.indexOf(k) < 0) throw new Error('Cannot edit ' + k);
        let v = u.fields[k];
        if (k === 'Date') checkDate_(v);
        if (k === 'ReconMonth' && !/^\d{4}-\d{2}$/.test(v)) throw new Error('Bad month');
        if (k === 'Alloc') checkAlloc_(ctx, v, false);
        if (k === 'Amount' || k === 'VAT') v = round2_(v);
        else v = clean_(v);
        sh.getRange(r + 2, ENTRY_HEADERS.indexOf(k) + 1).setValue(v);
      });
    });
    refreshOpenStatus_(ss, code);
    return { entries: entriesFor_(ss, code) };
  } finally {
    lock.releaseLock();
  }
}

function actVoid_(ctx, req) {
  const ss = ctx.ss, code = ctx.station.Code;
  const sh = ss.getSheetByName('Entries');
  const ids = idColumn_(sh), stations = stationColumn_(sh);
  const r = ids.indexOf(req.id);
  if (r < 0 || stations[r] !== code) throw new Error('Entry not found');
  sh.getRange(r + 2, ENTRY_HEADERS.indexOf('Status') + 1).setValue('VOID');
  refreshOpenStatus_(ss, code);
  return { entries: entriesFor_(ss, code) };
}

function actSetMonth_(ctx, req) {
  const ss = ctx.ss, code = ctx.station.Code;
  if (!/^\d{4}-\d{2}$/.test(req.month || '')) throw new Error('Bad month');
  const sh = ss.getSheetByName('Months');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const r = monthRow_(sh, code, req.month);
    const existing = r > 0 ? sh.getRange(r, 1, 1, MONTH_HEADERS.length).getValues()[0]
      : [code, req.month, '', '', '', '', '', ''];
    const val = function (k, idx) {
      if (!(k in req)) return existing[idx];
      return req[k] === '' || req[k] === null ? '' : round2_(req[k]);
    };
    const row = [code, req.month, val('openBank', 2), val('openCash', 3), val('stmtClose', 4),
      val('cashCount', 5), existing[6] || '', new Date().toISOString()];
    const target = r > 0 ? sh.getRange(r, 1, 1, row.length) : sh.getRange(sh.getLastRow() + 1, 1, 1, row.length);
    target.setNumberFormats([['@', '@', '0.00', '0.00', '0.00', '0.00', '@', '@']]);
    target.setValues([row]);
    return { months: monthsFor_(ss, code) };
  } finally {
    lock.releaseLock();
  }
}

/** Combine every photo counted in a month (ReconMonth) for this station into one PDF in that month's folder. */
function actBuildPdf_(ctx, req) {
  const ss = ctx.ss, code = ctx.station.Code, month = req.month;
  if (!/^\d{4}-\d{2}$/.test(month || '')) throw new Error('Bad month');
  const entries = entriesFor_(ss, code).filter(function (e) {
    return e.ReconMonth === month && e.Status !== 'VOID' && e.FileId;
  });
  entries.sort(function (a, b) {
    return (a.Date + pad_(a.Voucher)).localeCompare(b.Date + pad_(b.Voucher));
  });
  const seen = {};
  const docs = [];
  entries.forEach(function (e) {
    if (seen[e.FileId]) { seen[e.FileId].allocs.push(e.Alloc); return; }
    seen[e.FileId] = { e: e, allocs: [e.Alloc] };
    docs.push(seen[e.FileId]);
  });
  if (!docs.length) throw new Error('No photos for ' + month + ' yet');

  const title = ctx.station.Name + ' ' + month + ' Petty Cash Slips';
  const doc = DocumentApp.create('tmp ' + title);
  const body = doc.getBody();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(36).setMarginRight(36);
  const maxW = body.getPageWidth() - 72;
  const maxH = body.getPageHeight() - 72 - 40;
  body.appendParagraph(APP_NAME + ' – ' + ctx.station.Name + ' – slips and proofs for ' + month).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(docs.length + ' documents. Built ' + new Date().toLocaleString() + ' by ' + ctx.user.name);
  docs.forEach(function (d, i) {
    const e = d.e;
    if (i > 0) body.appendPageBreak();
    const total = entries.filter(function (x) { return x.FileId === e.FileId; })
      .reduce(function (s, x) { return s + Number(x.Amount); }, 0);
    const head = [e.Voucher ? 'V' + e.Voucher : label_(e.Type), e.Date, e.Name || '',
      'R' + total.toFixed(2), d.allocs.filter(String).join('/'), sideLabel_(e)].filter(String).join('  ·  ');
    body.appendParagraph(head).editAsText().setBold(true);
    const blob = DriveApp.getFileById(e.FileId).getBlob();
    const img = body.appendImage(blob);
    const w = img.getWidth(), h = img.getHeight();
    const scale = Math.min(maxW / w, maxH / h, 1);
    img.setWidth(Math.round(w * scale)).setHeight(Math.round(h * scale));
  });
  doc.saveAndClose();
  const pdfBlob = DriveApp.getFileById(doc.getId()).getAs('application/pdf').setName(title + '.pdf');
  const folder = monthFolder_(ctx.station, month);
  const old = folder.getFilesByName(title + '.pdf');
  while (old.hasNext()) old.next().setTrashed(true);
  const pdf = folder.createFile(pdfBlob);
  DriveApp.getFileById(doc.getId()).setTrashed(true);

  // remember the link on the month row
  const sh = ss.getSheetByName('Months');
  const r = monthRow_(sh, code, month);
  if (r < 0) {
    const t = sh.getRange(sh.getLastRow() + 1, 1, 1, MONTH_HEADERS.length);
    t.setNumberFormats([['@', '@', '0.00', '0.00', '0.00', '0.00', '@', '@']]);
    t.setValues([[code, month, '', '', '', '', pdf.getUrl(), new Date().toISOString()]]);
  } else {
    sh.getRange(r, MONTH_HEADERS.indexOf('PdfUrl') + 1).setValue(pdf.getUrl());
  }
  return { url: pdf.getUrl(), count: docs.length, months: monthsFor_(ss, code) };
}

function actChangePin_(ctx, req) {
  const p = String(req.newPin || '');
  if (!/^\d{4,8}$/.test(p)) throw new Error('PIN must be 4 to 8 digits');
  const sh = ctx.ss.getSheetByName('Users');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const users = readTable_(sh);
    const me = users.findIndex(function (u) { return u.PIN === String(req.pin || '').trim() && u.Active !== 'N'; });
    if (me < 0) throw new Error('Wrong PIN');
    if (users.some(function (u, i) { return i !== me && u.PIN === p; })) throw new Error('That PIN is taken – choose another');
    sh.getRange(me + 2, USER_HEADERS.indexOf('PIN') + 1).setNumberFormat('@').setValue(p);
    return { changed: true };
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------------ helpers */

function entriesFor_(ss, code) {
  return readTable_(ss.getSheetByName('Entries')).filter(function (e) { return e.Station === code; });
}

function monthsFor_(ss, code) {
  return readTable_(ss.getSheetByName('Months')).filter(function (m) { return m.Station === code; });
}

function monthRow_(sh, code, month) {
  const data = sh.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === code && cellText_(data[i][1], 'Month', tz) === month) return i + 1;
  }
  return -1;
}

/** Staff withdrawals stay OPEN until repayments linked to them cover the amount. Own-card slips likewise. */
function refreshOpenStatus_(ss, code) {
  const sh = ss.getSheetByName('Entries');
  const all = readTable_(sh);
  const statusCol = ENTRY_HEADERS.indexOf('Status') + 1;
  const rows = all.filter(function (x) { return x.Station === code; });
  all.forEach(function (e, i) {
    if (e.Station !== code || e.Status === 'VOID') return;
    let want = null;
    if (e.Kind === 'OUT' && e.Type === 'STAFF_WITHDRAWAL') {
      const paid = rows.filter(function (x) {
        return x.Status !== 'VOID' && x.Kind === 'IN' && x.Type === 'REPAYMENT' && x.LinkID === e.ID;
      }).reduce(function (s, x) { return s + Number(x.Amount); }, 0);
      want = paid + 0.005 >= Number(e.Amount) ? 'OK' : 'OPEN';
    }
    if (e.Kind === 'SLIP' && e.Side === 'OWN') {
      const slipTotal = rows.filter(function (x) {
        return x.Status !== 'VOID' && x.Kind === 'SLIP' && x.Voucher === e.Voucher;
      }).reduce(function (s, x) { return s + Number(x.Amount); }, 0);
      const paid = rows.filter(function (x) {
        return x.Status !== 'VOID' && x.Kind === 'OUT' && x.Type === 'REIMBURSE' && x.LinkID === 'V' + e.Voucher;
      }).reduce(function (s, x) { return s + Number(x.Amount); }, 0);
      want = paid + 0.005 >= slipTotal ? 'OK' : 'OPEN';
    }
    if (e.Kind === 'SLIP' && e.Side !== 'OWN' && e.Status === 'OPEN') want = 'OK';
    if (e.Kind === 'OUT' && e.Type !== 'STAFF_WITHDRAWAL' && e.Status === 'OPEN') want = 'OK';
    if (want && want !== e.Status) sh.getRange(i + 2, statusCol).setValue(want);
  });
}

function saveImage_(station, b64, date, name) {
  const clean = String(b64).replace(/^data:image\/\w+;base64,/, '');
  const bytes = Utilities.base64Decode(clean);
  const fileName = (date + ' ' + station.Code + ' ' + name).replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim() + '.jpg';
  const blob = Utilities.newBlob(bytes, 'image/jpeg', fileName);
  const file = monthFolder_(station, date.slice(0, 7)).createFile(blob);
  return { id: file.getId(), url: file.getUrl() };
}

function monthFolder_(station, month) {
  const root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID'));
  const s = root.getFoldersByName(station.Name);
  const sf = s.hasNext() ? s.next() : root.createFolder(station.Name);
  const it = sf.getFoldersByName(month);
  return it.hasNext() ? it.next() : sf.createFolder(month);
}

function nextVoucher_(ss, code) {
  const key = 'NextVoucher_' + code;
  const n = parseInt(getConfig_(ss, key) || '1', 10);
  setConfig_(ss, key, String(n + 1));
  return n;
}

function getConfig_(ss, key) {
  const data = ss.getSheetByName('Config').getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) if (data[i][0] === key) return data[i][1];
  return '';
}

function setConfig_(ss, key, value) {
  const sh = ss.getSheetByName('Config');
  const data = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) { sh.getRange(i + 1, 2).setValue(value); return; }
  }
  sh.appendRow([key, value]);
}

/** A random PIN no-one else has. */
function newPin_(taken) {
  let p;
  do { p = String(100000 + Math.floor(Math.random() * 900000)); } while (taken.indexOf(p) >= 0);
  return p;
}

const MONEY_FIELDS = ['Amount', 'VAT', 'OpenBank', 'OpenCash', 'StmtClose', 'CashCount'];

/**
 * Reads raw values (not display text) so the sheet's locale can never turn 1332.45 into "1332,45".
 * Money -> "1332.45", dates (if Sheets ever converted one) -> "yyyy-MM-dd", everything else -> trimmed text.
 */
function readTable_(sh) {
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const head = data[0].map(String);
  const tz = Session.getScriptTimeZone();
  return data.slice(1).filter(function (r) { return r.join('') !== ''; }).map(function (r) {
    const o = {};
    head.forEach(function (h, i) { o[h] = cellText_(r[i], h, tz); });
    return o;
  });
}

function cellText_(v, h, tz) {
  if (v === '' || v === null || v === undefined) return '';
  if (MONEY_FIELDS.indexOf(h) >= 0) return (Math.round(Number(v) * 100) / 100).toFixed(2);
  if (v instanceof Date) return Utilities.formatDate(v, tz, h === 'ReconMonth' || h === 'Month' ? 'yyyy-MM' : 'yyyy-MM-dd');
  return String(v).trim();
}

function idColumn_(sh) { return column_(sh, 'ID'); }
function stationColumn_(sh) { return column_(sh, 'Station'); }
function column_(sh, name) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, ENTRY_HEADERS.indexOf(name) + 1, n, 1).getDisplayValues().map(function (r) { return r[0]; });
}

function appendRows_(sh, rows) {
  const range = sh.getRange(sh.getLastRow() + 1, 1, rows.length, ENTRY_HEADERS.length);
  const fmt = rows.map(function () {
    return ENTRY_HEADERS.map(function (h) { return (h === 'Amount' || h === 'VAT') ? '0.00' : '@'; });
  });
  range.setNumberFormats(fmt); // text columns stay text (dates/months never turn into Sheets dates)
  range.setValues(rows);
}

function rowFrom_(o) { return ENTRY_HEADERS.map(function (h) { return o[h] === undefined ? '' : o[h]; }); }

function objFromRow_(r) {
  const o = {};
  ENTRY_HEADERS.forEach(function (h, i) {
    o[h] = (h === 'Amount' || h === 'VAT') ? Number(r[i]).toFixed(2) : String(r[i]);
  });
  return o;
}

function checkDate_(d) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d || '')) throw new Error('Bad date');
}

function round2_(v) { return Math.round(Number(v || 0) * 100) / 100; }
// Text columns are formatted as plain text, so nothing typed here is ever run as a formula.
function clean_(v) { return String(v === undefined || v === null ? '' : v).slice(0, 300); }
function pad_(v) { return ('000000' + (v || '')).slice(-6); }

function label_(type) {
  return ({
    PURCHASE: 'Slip', KITCHEN_SALES: 'Kitchen sales', STORE_SALES: 'Store cash sales',
    CASH_EXCHANGE: 'Cash exchange', REPAYMENT: 'Repayment', TRANSFER: 'Transfer from main',
    OTHER_IN: 'Other in', STAFF_WITHDRAWAL: 'Staff withdrawal', CASH_DRAW: 'Cash draw',
    REIMBURSE: 'Own card reimbursed', BANK_FEE: 'Bank fee', OTHER_OUT: 'Other out'
  })[type] || type || '';
}

function sideLabel_(e) {
  if (e.Kind === 'SLIP') return ({ CARD: 'Card', CASH: 'Cash', EFT: 'EFT main acc', OWN: 'Own card' })[e.Side] || '';
  return e.Side === 'BANK' ? 'Bank' : 'Cash box';
}


/** Run once from the editor to grant every permission the web app needs (Sheets, Drive, Docs). Changes nothing. */
function authorizeAll() {
  SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')).getName();
  DriveApp.getRootFolder().getName();
  DocumentApp.getActiveDocument;
  Logger.log('All permissions granted.');
}
