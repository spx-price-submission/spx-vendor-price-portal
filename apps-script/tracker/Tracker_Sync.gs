/**
 * Manpower Bidding Price Tracker – Google Apps Script
 * Paste into Extensions > Apps Script of the TRACKER Google Sheet. Save, reload the sheet,
 * then run menu  Bidding Tools > Set up daily sync  once.
 *
 *  - pullSubmissions(): copies new vendor submissions into the Database tab
 *      (daily trigger 18:00 Bangkok + "Retrieve now" checkbox on the Sync tab + menu)
 *  - syncLists():       pushes vendor & facility names from Lists to the vendor form file
 *  - saveFormEntry() / clearForm(): internal Form tab (replaces the Excel macro)
 */

const CFG = {
  TZ: 'Asia/Bangkok',
  DAILY_HOUR: 18,
  DB_SHEET: 'Database', DB_FIRST: 6, DB_LAST: 1005,   // Database rows with formulas
  SYNC_SHEET: 'Sync', SRC_CELL: 'C5', CHECK_CELL: 'C7', LAST_CELL: 'C9', RESULT_CELL: 'C10', LOG_ROW: 15,
  LISTS_SHEET: 'Lists', VENDOR_RANGE: 'B6:B25', FACILITY_RANGE: 'D6:D35', MONTH_RANGE: 'V6:V29',
  TREND_SHEET: 'Trend', TREND_PAIR_ROWS: 120,          // rows per vendor×facility table on Trend
  SRC_SUBMISSIONS: 'Submissions', SRC_FORM: 'Form Responses', SRC_VENDORS: 'Vendors', SRC_FACILITIES: 'Facilities'
};

// ---------------------------------------------------------------- menu & setup
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Bidding Tools')
    .addItem('Retrieve vendor submissions now', 'pullSubmissionsManual')
    .addItem('Push vendor & facility lists to vendor form', 'syncLists')
    .addSeparator()
    .addItem('Save Form tab entry', 'saveFormEntry')
    .addItem('Clear Form tab', 'clearForm')
    .addSeparator()
    .addItem('Set up daily sync (run once)', 'setupTriggers')
    .addToUi();
}

function setupTriggers() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['pullSubmissionsDaily', 'onEditInstalled'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('pullSubmissionsDaily').timeBased().everyDays(1).atHour(CFG.DAILY_HOUR).inTimezone(CFG.TZ).create();
  ScriptApp.newTrigger('onEditInstalled').forSpreadsheet(ss).onEdit().create();

  const sync = ss.getSheetByName(CFG.SYNC_SHEET);
  const cb = sync.getRange(CFG.CHECK_CELL);
  cb.clearContent().setFontStyle('normal').insertCheckboxes().setValue(false);
  sync.getRange('D7').setValue('← tick to retrieve now').setFontColor('#9C5700');
  SpreadsheetApp.getUi().alert('Done. New vendor submissions will be pulled every day around 18:00 (Bangkok).\n' +
    'Tick the checkbox on the Sync tab (C7) or use Bidding Tools > Retrieve to pull at any time.');
}

/** Installable onEdit: the Sync tab checkbox works as a "Retrieve now" button. */
function onEditInstalled(e) {
  const r = e && e.range;
  if (!r || r.getSheet().getName() !== CFG.SYNC_SHEET || r.getA1Notation() !== CFG.CHECK_CELL) return;
  if (r.getValue() !== true) return;
  try { pullSubmissions_('Checkbox'); } finally { r.setValue(false); }
}

function pullSubmissionsDaily() { pullSubmissions_('Daily 18:00'); }
function pullSubmissionsManual() {
  const res = pullSubmissions_('Menu');
  SpreadsheetApp.getUi().alert(res);
}

// ---------------------------------------------------------------- retrieve
/**
 * Reads new rows from the vendor submission file – from either or both sources:
 *   'Submissions'    – the Apps Script web app (Option: script web app)
 *   'Form Responses' – the hosted web page that saves through a Google Form (Option A)
 * Each vendor line becomes TWO Database rows: 'Without shift' and 'With shift'.
 */
function pullSubmissions_(trigger) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return log_(trigger, 0, 'Skipped – another retrieval is running.');
  try {
    const ss = SpreadsheetApp.getActive();
    const src = openSource_();
    const lines = collectWebAppLines_(src).concat(collectFormLines_(src));
    const rejected = lines.filter(l => l.reject);
    rejected.forEach(l => l.stamp('REJECTED – ' + l.reject));
    const good = lines.filter(l => !l.reject && l.rows.length);
    if (!good.length) {
      return log_(trigger, 0, rejected.length ? 'No valid new lines. Rejected ' + rejected.length + ' (see Imported At column in the vendor file).'
                                             : 'Up to date – no new submissions.');
    }
    const rows = [].concat.apply([], good.map(l => l.rows));
    const added = ensureInLists_(ss, rows);          // new vendor / facility names → Lists, so Tracker & Trend pick them up
    const written = appendToDatabase_(ss, rows);
    SpreadsheetApp.flush();                           // recalculate Tracker / Trend before we report

    const stamp = new Date();
    let left = written, done = 0;
    good.forEach(l => {
      if (left >= l.rows.length) { l.stamp(stamp); left -= l.rows.length; done++; } else { left = 0; }
    });

    let msg = written < rows.length
      ? 'Imported ' + written + ' of ' + rows.length + ' rows – Database is full, extend its formula rows.'
      : 'Imported ' + done + ' line(s) as ' + written + ' price row(s).';
    if (rejected.length) msg += ' Rejected ' + rejected.length + ' (bad name/code or data – see vendor file).';
    if (added.length) msg += ' Added to Lists: ' + added.join(', ') + '.';
    const warn = healthCheck_(ss, rows);
    if (warn.length) msg += ' ⚠ ' + warn.join(' ');
    return log_(trigger, written, msg);
  } catch (err) {
    return log_(trigger, 0, 'ERROR: ' + err.message);
  } finally {
    lock.releaseLock();
  }
}

/** Builds the two Database rows (without / with shift). Database C:M =
 *  Date, Vendor, Facility, Nationality, Scheme, Shift, Price Type, Price, BAU HC, Campaign HC, Remark */
function toRows_(date, vendor, facility, nat, scheme, type, priceNo, priceWith, bauHc, campHc, remark) {
  const fac = String(facility).replace(/\s*\(BM\)\s*$/i, '').trim();
  const rows = [];
  [['Without shift', priceNo], ['With shift', priceWith]].forEach(([shift, price]) => {
    if (price === '' || price === null || isNaN(Number(price)) || Number(price) <= 0) return;
    rows.push([date, vendor, fac, nat, scheme, shift, type, Number(price), bauHc, campHc, remark]);
  });
  return rows;
}

/** Source 1 – 'Submissions' tab written by the Apps Script web app (already validated there). */
function collectWebAppLines_(src) {
  const sub = src.getSheetByName(CFG.SRC_SUBMISSIONS);
  if (!sub || sub.getLastRow() < 2) return [];
  const data = sub.getRange(2, 1, sub.getLastRow() - 1, 13).getValues();   // A:M
  const out = [];
  data.forEach((v, i) => {
    if (v[0] === '' || v[9] !== '') return;                            // J = Imported At
    const row = i + 2;
    // A ID, B Timestamp, C Vendor, D Facility, E Type, F Price without shift, G Max BAU HC, H Max Camp HC,
    // I Remark, J Imported At, K Nationality, L Working Scheme, M Price with shift
    const remark = [v[8], 'Vendor form ' + v[0]].filter(String).join(' · ');
    out.push({
      rows: toRows_(v[1], v[2], v[3], String(v[10] || 'TH').toUpperCase(), String(v[11] || '8hr'), v[4], v[5], v[12], v[6], v[7], remark),
      stamp: x => sub.getRange(row, 10).setValue(x)
    });
  });
  return out;
}

/** Source 2 – 'Form Responses' tab filled by the hosted web page through a Google Form.
 *  Vendor name + access code are checked here against the Vendors tab (the page itself is public). */
function collectFormLines_(src) {
  const sh = src.getSheetByName(CFG.SRC_FORM);
  if (!sh || sh.getLastRow() < 2) return [];
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const c = name => head.indexOf(name);
  let impCol = c('Imported At');
  if (impCol < 0) { impCol = head.length; sh.getRange(1, impCol + 1).setValue('Imported At'); }
  const need = ['Timestamp', 'Vendor Name', 'Access Code', 'Facility', 'Nationality', 'Working Scheme', 'Price Type',
                'Price without shift', 'Price with shift', 'Max BAU HC', 'Max Campaign HC', 'Remark', 'Submission Ref'];
  const miss = need.filter(n => c(n) < 0);
  if (miss.length) throw new Error('Form Responses tab is missing column(s): ' + miss.join(', '));

  const vendors = vendorIndex_(src);
  const facs = (src.getSheetByName(CFG.SRC_FACILITIES) ? src.getSheetByName(CFG.SRC_FACILITIES).getDataRange().getValues().slice(1) : [])
    .map(r => String(r[0]).replace(/\s*\(BM\)\s*$/i, '').trim().toLowerCase()).filter(String);
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(head.length, impCol + 1)).getValues();
  const out = [];
  data.forEach((v, i) => {
    if (v[c('Timestamp')] === '' || String(v[impCol]) !== '') return;
    const row = i + 2;
    const stamp = x => sh.getRange(row, impCol + 1).setValue(x);
    const g = n => v[c(n)];
    const official = vendors[normName_(g('Vendor Name')) + '|' + String(g('Access Code')).trim().toLowerCase()];
    const nat = String(g('Nationality')).trim().toUpperCase(), scheme = String(g('Working Scheme')).trim();
    const type = String(g('Price Type')).trim(), fac = String(g('Facility')).trim();
    let reject = '';
    if (!official) reject = 'company name / access code';
    else if (facs.length && facs.indexOf(fac.replace(/\s*\(BM\)\s*$/i, '').toLowerCase()) < 0) reject = 'unknown facility ' + fac;
    else if (['TH', 'BM'].indexOf(nat) < 0 || ['8hr', '10hr'].indexOf(scheme) < 0 || ['BAU', 'Campaign'].indexOf(type) < 0) reject = 'invalid nationality / scheme / price type';
    else if (!(Number(g('Price without shift')) > 0) || !(Number(g('Price with shift')) > 0)) reject = 'invalid price';
    const remark = [String(g('Remark')).replace(/^[=+\-@]/, "'$&"), 'Web page ' + g('Submission Ref')].filter(String).join(' · ');
    out.push({
      reject: reject, stamp: stamp,
      rows: reject ? [] : toRows_(g('Timestamp'), official, fac, nat, scheme, type, g('Price without shift'), g('Price with shift'),
                                  g('Max BAU HC'), g('Max Campaign HC'), remark)
    });
  });
  return out;
}

/** "name|code" → official vendor name, from the Vendors tab of the vendor file. */
function vendorIndex_(src) {
  const vs = src.getSheetByName(CFG.SRC_VENDORS);
  const idx = {};
  if (!vs || vs.getLastRow() < 2) return idx;
  vs.getRange(2, 1, vs.getLastRow() - 1, 2).getValues().forEach(r => {
    const name = String(r[0]).trim(), code = String(r[1]).trim().toLowerCase();
    if (name && code) idx[normName_(name) + '|' + code] = name;
  });
  return idx;
}

/** Same rule as the vendor form: ignore case, spaces and punctuation; keep Thai letters. */
function normName_(v) {
  return String(v || '').toLowerCase().replace(/[^a-z0-9\u0E00-\u0E7F]/g, '');
}

/**
 * Tracker and Trend only show vendors / facilities that exist in the Lists tab.
 * Adds any new names from incoming rows into the first empty Lists cells. Returns the names added.
 */
function ensureInLists_(ss, rows) {
  const lists = ss.getSheetByName(CFG.LISTS_SHEET);
  const added = [];
  [[CFG.VENDOR_RANGE, 1], [CFG.FACILITY_RANGE, 2]].forEach(([a1, col]) => {
    const rng = lists.getRange(a1);
    const vals = rng.getValues().map(r => String(r[0]).trim());
    rows.forEach(r => {
      const name = String(r[col]).trim();
      if (!name || vals.some(v => v.toLowerCase() === name.toLowerCase())) return;
      const free = vals.indexOf('');
      if (free < 0) return;                              // list full – reported by healthCheck_
      vals[free] = name; added.push(name);
    });
    rng.setValues(vals.map(v => [v]));
  });
  return added;
}

/** Warnings that would stop new data showing in Tracker / Trend. */
function healthCheck_(ss, rows) {
  const warn = [];
  const lists = ss.getSheetByName(CFG.LISTS_SHEET);
  const vendors = lists.getRange(CFG.VENDOR_RANGE).getValues().map(r => String(r[0]).trim().toLowerCase());
  const facs = lists.getRange(CFG.FACILITY_RANGE).getValues().map(r => String(r[0]).trim().toLowerCase());
  const missing = rows.filter(r => vendors.indexOf(String(r[1]).trim().toLowerCase()) < 0 ||
                                   facs.indexOf(String(r[2]).trim().toLowerCase()) < 0);
  if (missing.length) warn.push(missing.length + ' row(s) use a vendor/facility that could not be added to Lists (list full).');

  const months = lists.getRange(CFG.MONTH_RANGE).getValues().map(r => r[0]).filter(d => d instanceof Date);
  const lastMonth = months.length ? months[months.length - 1] : null;
  if (lastMonth && rows.some(r => r[0] instanceof Date &&
      (r[0].getFullYear() * 12 + r[0].getMonth()) > (lastMonth.getFullYear() * 12 + lastMonth.getMonth()))) {
    warn.push('Some quotes are after the last Trend month – move Lists!X6 forward.');
  }

  const db = ss.getSheetByName(CFG.DB_SHEET);
  const used = db.getRange(CFG.DB_FIRST, 3, CFG.DB_LAST - CFG.DB_FIRST + 1, 1).getValues().filter(r => r[0] !== '').length;
  const cap = CFG.DB_LAST - CFG.DB_FIRST + 1;
  if (used > cap * 0.9) warn.push('Database ' + used + '/' + cap + ' rows used – archive old rows soon.');

  const pairs = db.getRange(CFG.DB_FIRST, 21, cap, 1).getValues().filter(r => r[0] !== '').length;   // column U = Combo Key
  if (pairs > CFG.TREND_PAIR_ROWS) warn.push(pairs + ' vendor×facility pairs – Trend tables show the first ' + CFG.TREND_PAIR_ROWS + '.');
  return warn;
}

/** Writes rows into Database C:L at the next empty row. Returns how many were written. */
function appendToDatabase_(ss, rows) {
  const db = ss.getSheetByName(CFG.DB_SHEET);
  const dates = db.getRange(CFG.DB_FIRST, 3, CFG.DB_LAST - CFG.DB_FIRST + 1, 1).getValues();
  let lastUsed = -1;
  dates.forEach((r, i) => { if (r[0] !== '') lastUsed = i; });
  const start = CFG.DB_FIRST + lastUsed + 1;
  const room = CFG.DB_LAST - start + 1;
  const n = Math.min(room, rows.length);
  if (n > 0) db.getRange(start, 3, n, rows[0].length).setValues(rows.slice(0, n));
  return Math.max(n, 0);
}

function openSource_() {
  const raw = String(SpreadsheetApp.getActive().getSheetByName(CFG.SYNC_SHEET).getRange(CFG.SRC_CELL).getValue()).trim();
  const m = raw.match(/\/d\/([a-zA-Z0-9_-]{20,})/) || raw.match(/^([a-zA-Z0-9_-]{20,})$/);
  if (!m) throw new Error('Paste the vendor submission file link in Sync!' + CFG.SRC_CELL + '.');
  return SpreadsheetApp.openById(m[1]);
}

function log_(trigger, n, msg) {
  const sync = SpreadsheetApp.getActive().getSheetByName(CFG.SYNC_SHEET);
  const now = new Date();
  sync.getRange(CFG.LAST_CELL).setValue(now).setNumberFormat('dd-mmm-yyyy hh:mm');
  sync.getRange(CFG.RESULT_CELL).setValue(msg).setFontColor(msg.indexOf('ERROR') === 0 ? '#C00000' : '#00873C');
  sync.insertRowBefore(CFG.LOG_ROW);
  sync.getRange(CFG.LOG_ROW, 2).setValue(now).setNumberFormat('dd-mmm-yyyy hh:mm');
  sync.getRange(CFG.LOG_ROW, 3).setValue(trigger);
  sync.getRange(CFG.LOG_ROW, 5).setValue(n);
  sync.getRange(CFG.LOG_ROW, 7).setValue(msg);
  const lastRow = sync.getLastRow();
  if (lastRow > CFG.LOG_ROW + 199) sync.deleteRows(CFG.LOG_ROW + 200, lastRow - CFG.LOG_ROW - 199);  // keep 200 entries
  return msg;
}

// ---------------------------------------------------------------- list sync
function syncLists() {
  const ss = SpreadsheetApp.getActive();
  const lists = ss.getSheetByName(CFG.LISTS_SHEET);
  const vendors = lists.getRange(CFG.VENDOR_RANGE).getValues().map(r => String(r[0]).trim()).filter(String);
  const facilities = lists.getRange(CFG.FACILITY_RANGE).getValues().map(r => String(r[0]).trim()).filter(String);
  const src = openSource_();

  // vendors: keep each vendor's access code + note
  const vs = src.getSheetByName(CFG.SRC_VENDORS);
  const old = vs.getLastRow() > 1 ? vs.getRange(2, 1, vs.getLastRow() - 1, 3).getValues() : [];
  const keep = {};
  old.forEach(r => { keep[String(r[0]).trim()] = [r[1], r[2]]; });
  if (vs.getLastRow() > 1) vs.getRange(2, 1, vs.getLastRow() - 1, 3).clearContent();
  if (vendors.length) vs.getRange(2, 1, vendors.length, 3).setValues(vendors.map(v => [v].concat(keep[v] || ['', ''])));

  const fs = src.getSheetByName(CFG.SRC_FACILITIES);
  if (fs.getLastRow() > 1) fs.getRange(2, 1, fs.getLastRow() - 1, 1).clearContent();
  if (facilities.length) fs.getRange(2, 1, facilities.length, 1).setValues(facilities.map(f => [f]));

  const dropped = old.map(r => String(r[0]).trim()).filter(v => v && vendors.indexOf(v) < 0);
  SpreadsheetApp.getUi().alert('Vendor form updated: ' + vendors.length + ' vendors, ' + facilities.length + ' facilities.' +
    (dropped.length ? '\nRemoved from the form (not in Lists): ' + dropped.join(', ') : ''));
}

// ---------------------------------------------------------------- internal Form tab
function saveFormEntry() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const get = n => ss.getRangeByName(n);
  SpreadsheetApp.flush();
  const status = String(get('frmStatus').getValue());
  if (status.charAt(0) !== '✓') {
    if (status.indexOf('Duplicate') >= 0) {
      if (ui.alert('Duplicate quote', status.replace('⚠', '').trim() + '\n\nSave anyway?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    } else {
      ui.alert('Cannot save yet', status.replace('⚠', '').trim(), ui.ButtonSet.OK);
      return;
    }
  }
  const row = [get('frmDate').getValue(), get('frmVendor').getValue(), get('frmFacility').getValue(),
               get('frmNat').getValue(), get('frmScheme').getValue(), get('frmShift').getValue(), get('frmType').getValue(),
               get('frmPrice').getValue(), get('frmBauHC').getValue(), get('frmCampHC').getValue(), get('frmRemark').getValue()];
  if (appendToDatabase_(ss, [row]) !== 1) { ui.alert('Database is full – extend its formula rows.'); return; }
  get('frmLastSaved').setValue('✓  Saved: ' + row[1] + '  ·  ' + row[2] + ' ' + row[3] + ' ' + row[4] + ' ' + row[5] + '  ·  ' + row[6] + ' ' +
    Utilities.formatString('%.2f', row[7]) + '   (' + Utilities.formatDate(new Date(), CFG.TZ, 'HH:mm') + ')');
  ['frmFacility', 'frmPrice', 'frmBauHC', 'frmCampHC', 'frmRemark'].forEach(n => get(n).clearContent());
  get('frmFacility').activate();
}

function clearForm() {
  const ss = SpreadsheetApp.getActive();
  ['frmDate', 'frmVendor', 'frmFacility', 'frmPrice', 'frmBauHC', 'frmCampHC', 'frmRemark', 'frmLastSaved']
    .forEach(n => ss.getRangeByName(n).clearContent());
  ss.getRangeByName('frmType').setValue('BAU');
  ss.getRangeByName('frmNat').setValue('TH');
  ss.getRangeByName('frmScheme').setValue('8hr');
  ss.getRangeByName('frmShift').setValue('With shift');
  ss.getRangeByName('frmDate').activate();
}
