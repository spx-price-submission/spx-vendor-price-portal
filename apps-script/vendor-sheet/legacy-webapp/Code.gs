/**
 * SPX Manpower Bidding – Vendor Price Submission (web app)
 * Paste into Extensions > Apps Script of the "Vendor Price Submission" Google Sheet (as Code.gs).
 * Add an HTML file named "Form" with the contents of Vendor_Form.html.
 * Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 *
 * Submissions tab columns:
 *  A Submission ID | B Timestamp | C Vendor Name | D Facility Name | E Price Type
 *  F Price without shift (THB) | G Max BAU HC | H Max Campaign HC | I Remark | J Imported At
 *  K Nationality | L Working Scheme | M Price with shift (THB)
 */

const SHEET_SUBMISSIONS = 'Submissions';
const SHEET_VENDORS = 'Vendors';
const SHEET_FACILITIES = 'Facilities';
const PRICE_TYPES = ['BAU', 'Campaign'];
const NATIONALITIES = ['TH', 'BM'];
const SCHEMES = ['8hr', '10hr'];
const MAX_LINES = 40;
const MAX_FAILED_LOGINS = 5;          // per company name, then locked for LOCK_MINUTES
const LOCK_MINUTES = 15;

function doGet() {
  return HtmlService.createTemplateFromFile('Form')
    .evaluate()
    .setTitle('SPX Manpower Bidding – Price Submission')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Lists for the form dropdowns. Vendor names and access codes are never sent to the browser. */
function getOptions() {
  const ss = SpreadsheetApp.getActive();
  return {
    facilities: baseFacilities_(ss),
    priceTypes: PRICE_TYPES,
    nationalities: NATIONALITIES,
    schemes: SCHEMES,
    month: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM yyyy')
  };
}

/**
 * payload = {vendor, code, lines:[{facility, nationality, scheme, type, price, priceShift, bauHc, campHc, remark}]}
 *   price = price without shift, priceShift = price with shift. Both prices and both Max HC are required.
 * Returns a receipt or throws an Error whose message is shown to the vendor.
 */
function submitQuotes(payload) {
  const ss = SpreadsheetApp.getActive();
  const vendor = String(payload && payload.vendor || '').trim();
  const code = String(payload && payload.code || '').trim();
  const lines = (payload && payload.lines) || [];

  // --- vendor + access code (checked again on every submission)
  const vendorName = checkVendor_(vendor, code);

  // --- lines
  if (!lines.length) throw new Error('Add at least one facility price. / กรุณาใส่ราคาอย่างน้อย 1 ไซต์');
  if (lines.length > MAX_LINES) throw new Error('Maximum ' + MAX_LINES + ' lines per submission.');
  const facilities = baseFacilities_(ss);
  const seen = {};
  const clean = lines.map((l, i) => {
    const n = 'Line ' + (i + 1) + ': ';
    const facility = String(l.facility || '').trim();
    const nationality = String(l.nationality || '').trim();
    const scheme = String(l.scheme || '').trim();
    const type = String(l.type || '').trim();
    const price = Number(l.price);
    const priceShift = Number(l.priceShift);
    if (facilities.indexOf(facility) < 0) throw new Error(n + 'choose a facility. / กรุณาเลือกไซต์');
    if (NATIONALITIES.indexOf(nationality) < 0) throw new Error(n + 'choose nationality TH or BM. / กรุณาเลือกสัญชาติ');
    if (SCHEMES.indexOf(scheme) < 0) throw new Error(n + 'choose working scheme 8hr or 10hr. / กรุณาเลือกรูปแบบการทำงาน');
    if (PRICE_TYPES.indexOf(type) < 0) throw new Error(n + 'choose BAU or Campaign. / กรุณาเลือกประเภทราคา');
    if (!(price > 0) || price > 100000) throw new Error(n + 'enter a valid price without shift. / กรุณาใส่ราคาไม่รวมค่ากะให้ถูกต้อง');
    if (!(priceShift > 0) || priceShift > 100000) throw new Error(n + 'enter a valid price with shift. / กรุณาใส่ราคารวมค่ากะให้ถูกต้อง');
    const key = [facility, nationality, scheme, type].join('|');
    if (seen[key]) throw new Error(n + facility + ' · ' + nationality + ' · ' + scheme + ' · ' + type +
      ' is entered twice. / รายการซ้ำกัน');
    seen[key] = true;
    return {
      key: key, facility: facility, nationality: nationality, scheme: scheme, type: type,
      price: Math.round(price * 100) / 100, priceShift: Math.round(priceShift * 100) / 100,
      bauHc: hc_(l.bauHc, n + 'Max BAU HC', true), campHc: hc_(l.campHc, n + 'Max Campaign HC', true),
      remark: safeText_(l.remark).slice(0, 300)
    };
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = ss.getSheetByName(SHEET_SUBMISSIONS);
    const now = new Date();
    const tz = Session.getScriptTimeZone();
    const monthKey = Utilities.formatDate(now, tz, 'yyyy-MM');

    // earlier submissions this month by the same vendor (to flag revisions)
    const last = sheet.getLastRow();
    const prev = last > 1 ? sheet.getRange(2, 1, last - 1, 13).getValues() : [];
    const earlier = {};
    prev.forEach(r => {
      if (r[1] instanceof Date && Utilities.formatDate(r[1], tz, 'yyyy-MM') === monthKey && r[2] === vendorName) {
        earlier[[r[3], r[10] || 'TH', r[11], r[4]].join('|')] = true;
      }
    });

    const ref = 'S' + Utilities.formatDate(now, tz, 'yyMMddHHmmss') + '-' +
      Math.random().toString(36).slice(2, 6).toUpperCase();
    const rows = clean.map((l, i) => {
      l.revision = !!earlier[l.key];
      return [ref + '-' + (i + 1), now, vendorName, l.facility, l.type, l.price,
              l.bauHc, l.campHc, (l.revision ? '[Revision] ' : '') + l.remark, '',
              l.nationality, l.scheme, l.priceShift];
    });
    sheet.getRange(last + 1, 1, rows.length, rows[0].length).setValues(rows);

    return {
      ref: ref,
      when: Utilities.formatDate(now, tz, 'd MMM yyyy, HH:mm'),
      vendor: vendorName,
      lines: clean.map(l => ({
        facility: l.facility, nationality: l.nationality, scheme: l.scheme, type: l.type,
        price: l.price, priceShift: l.priceShift, bauHc: l.bauHc, campHc: l.campHc, revision: l.revision
      }))
    };
  } finally {
    lock.releaseLock();
  }
}

/** Step 1 of the form: company name (free text) + access code. Returns the official vendor name. */
function verifyVendor(name, code) {
  return { vendor: checkVendor_(name, code) };
}

/**
 * Matches the typed company name to the Vendors tab (ignoring case, spaces and punctuation) and checks
 * the access code. Every vendor must have a code. Same error for a wrong name or a wrong code, so the
 * form never reveals which vendor names exist. Repeated failures lock that name for a while.
 */
function checkVendor_(name, code) {
  const typed = norm_(name);
  const pass = String(code || '').trim();
  if (!typed || !pass) throw new Error('Please enter your company name and access code. / กรุณาใส่ชื่อบริษัทและรหัสเข้าใช้');
  const cache = CacheService.getScriptCache();
  const key = 'fail_' + typed.slice(0, 200);
  const fails = Number(cache.get(key) || 0);
  if (fails >= MAX_FAILED_LOGINS) {
    throw new Error('Too many attempts. Please try again in ' + LOCK_MINUTES + ' minutes or contact SPX. / ' +
      'ลองผิดหลายครั้งเกินไป กรุณาลองใหม่ใน ' + LOCK_MINUTES + ' นาที หรือติดต่อ SPX');
  }
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_VENDORS);
  const rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues() : [];
  const hit = rows.find(r => norm_(r[0]) === typed && String(r[1] || '').trim() !== '' &&
                             String(r[1]).trim().toLowerCase() === pass.toLowerCase());
  if (!hit) {
    cache.put(key, String(fails + 1), LOCK_MINUTES * 60);
    throw new Error('Company name or access code is not correct. / ชื่อบริษัทหรือรหัสเข้าใช้ไม่ถูกต้อง');
  }
  cache.remove(key);
  return String(hit[0]).trim();
}

/** "W.A.S  Add" → "wasadd"; keeps Thai letters. */
function norm_(v) {
  return String(v || '').toLowerCase().replace(/[^a-z0-9\u0E00-\u0E7F]/g, '');
}

// ---------- helpers ----------
function readColumn_(sheet, col) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues()
    .map(r => String(r[0]).trim()).filter(String);
}

/** Facility names without the " (BM)" suffix used in the tracker – nationality is chosen separately. */
function baseFacilities_(ss) {
  const out = [];
  readColumn_(ss.getSheetByName(SHEET_FACILITIES), 1).forEach(f => {
    const base = f.replace(/\s*\(BM\)\s*$/i, '').trim();
    if (base && out.indexOf(base) < 0) out.push(base);
  });
  return out;
}

function hc_(v, label, required) {
  if (v === '' || v === null || v === undefined) {
    if (required) throw new Error(label + ' is required (enter 0 if none). / กรุณาใส่จำนวนคน (ใส่ 0 ถ้าไม่มี)');
    return '';
  }
  const n = Number(v);
  if (!(n >= 0) || Math.floor(n) !== n || n > 100000) {
    throw new Error(label + ' must be a whole number. / จำนวนคนต้องเป็นจำนวนเต็ม');
  }
  return n;
}

/** Stops text being read as a formula when it lands in the sheet. */
function safeText_(v) {
  const s = String(v || '').trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
