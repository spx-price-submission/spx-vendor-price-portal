/**
 * Option A – hosted web page that saves through a hidden Google Form.
 * Add this file (Backend_Setup.gs) and an HTML file named exactly "WebPage" (WebPage.html)
 * to the Apps Script project of the VENDOR PRICE SUBMISSION Google Sheet. Save, reload the sheet.
 *
 *   Menu  Web Form > 1) Create form backend (run once)
 *         Web Form > 2) Build web page (index.html)   ← run again whenever codes or facilities change
 *
 * Upload the generated index.html to GitHub Pages / Netlify and send vendors that link.
 * Vendors never see this spreadsheet or the Google Form.
 */

const WEB = {
  FORM_TITLE: 'SPX Manpower Bidding – price submission backend (do not share)',
  RESP_SHEET: 'Form Responses',
  VENDORS: 'Vendors',
  FACILITIES: 'Facilities',
  ITERATIONS: 2000,                 // hash rounds – must match the page (it reads it from CONFIG)
  PROP_FORM: 'WEB_FORM_ID',
  PROP_SALT: 'WEB_SALT'
};

// key used by the page  →  question title in the Google Form (titles are what the tracker sync reads)
const WEB_FIELDS = [
  ['vendor', 'Vendor Name'], ['code', 'Access Code'], ['facility', 'Facility'], ['nationality', 'Nationality'],
  ['scheme', 'Working Scheme'], ['type', 'Price Type'], ['price', 'Price without shift'],
  ['priceShift', 'Price with shift'], ['bauHc', 'Max BAU HC'], ['campHc', 'Max Campaign HC'],
  ['remark', 'Remark'], ['ref', 'Submission Ref']
];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Web Form')
    .addItem('1) Create form backend (run once)', 'createFormBackend')
    .addItem('2) Build web page (index.html)', 'buildWebPage')
    .addToUi();
}

function createFormBackend() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty(WEB.PROP_FORM)) {
    ui.alert('The form backend already exists. Use "2) Build web page" instead.');
    return;
  }
  const form = FormApp.create(WEB.FORM_TITLE);
  form.setDescription('Backend for the SPX vendor price page. Vendors use the web page, not this form.');
  WEB_FIELDS.forEach(([, title]) => form.addTextItem().setTitle(title));
  form.setCollectEmail(false)
      .setAllowResponseEdits(false)
      .setLimitOneResponsePerUser(false)
      .setShowLinkToRespondAgain(false)
      .setConfirmationMessage('Received.');
  let loginNote = '';
  try { form.setRequireLogin(false); }            // lets people outside the company respond
  catch (e) { loginNote = '\n\n⚠ Could not allow outside responses automatically: ' + e.message +
              '\nOpen the form → Settings → Responses → turn OFF "Restrict to users in your organisation".'; }
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  props.setProperty(WEB.PROP_FORM, form.getId());
  props.setProperty(WEB.PROP_SALT, Utilities.getUuid());

  // find the new responses tab, give it a fixed name and add the "Imported At" column for the tracker
  SpreadsheetApp.flush();
  Utilities.sleep(3000);
  const sheet = ss.getSheets().find(s => {
    const url = s.getFormUrl();
    if (!url) return false;
    try { return FormApp.openByUrl(url).getId() === form.getId(); } catch (e) { return false; }
  });
  if (sheet) {
    if (!ss.getSheetByName(WEB.RESP_SHEET)) sheet.setName(WEB.RESP_SHEET);
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Imported At');
  } else {
    loginNote += '\n\n⚠ Rename the new "Form Responses 1" tab to exactly: ' + WEB.RESP_SHEET;
  }
  ui.alert('Form backend created.\n\nNext: Web Form > 2) Build web page.' + loginNote);
}

function buildWebPage() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getDocumentProperties();
  const formId = props.getProperty(WEB.PROP_FORM);
  if (!formId) { ui.alert('Run "1) Create form backend" first.'); return; }
  const form = FormApp.openById(formId);

  // entry IDs: build a pre-filled link and read the entry.NNN parameters in question order
  const items = form.getItems();
  const resp = form.createResponse();
  items.forEach(it => resp.withItemResponse(it.asTextItem().createResponse('x')));
  const ids = (resp.toPrefilledUrl().match(/entry\.\d+/g) || []);
  if (ids.length !== items.length) throw new Error('Could not read the form question IDs.');
  const entry = {};
  WEB_FIELDS.forEach(([key, title]) => {
    const k = items.findIndex(it => it.getTitle() === title);
    if (k < 0) throw new Error('Form question missing: ' + title + ' (do not rename the form questions).');
    entry[key] = ids[k];
  });

  // vendors → hashed name+code (names and codes are never written into the page)
  const salt = props.getProperty(WEB.PROP_SALT);
  const vs = ss.getSheetByName(WEB.VENDORS);
  const vRows = vs && vs.getLastRow() > 1 ? vs.getRange(2, 1, vs.getLastRow() - 1, 2).getValues() : [];
  const noCode = [];
  const hashes = [];
  vRows.forEach(r => {
    const name = String(r[0]).trim(), code = String(r[1]).trim();
    if (!name) return;
    if (!code) { noCode.push(name); return; }
    hashes.push(webHash_(salt, name, code));
  });
  hashes.sort();                                  // no link between list order and vendor order

  const fs = ss.getSheetByName(WEB.FACILITIES);
  const facilities = [];
  (fs && fs.getLastRow() > 1 ? fs.getRange(2, 1, fs.getLastRow() - 1, 1).getValues() : []).forEach(r => {
    const f = String(r[0]).replace(/\s*\(BM\)\s*$/i, '').trim();
    if (f && facilities.indexOf(f) < 0) facilities.push(f);
  });

  const config = {
    formAction: form.getPublishedUrl().replace(/\/viewform.*$/, '/formResponse'),
    entry: entry,
    facilities: facilities,
    salt: salt,
    iterations: WEB.ITERATIONS,
    hashes: hashes,
    builtAt: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm')
  };
  const configJs = '/* Generated by Web Form > Build web page – ' + config.builtAt + '. Do not edit by hand. */\n' +
    'var CONFIG = ' + JSON.stringify(config, null, 2) + ';\n';
  const html = HtmlService.createHtmlOutputFromFile('WebPage').getContent()
    .replace('/*__CONFIG__*/', 'var CONFIG = ' + JSON.stringify(config) + ';');
  const file = DriveApp.createFile('index.html', html, MimeType.HTML);          // all-in-one page (Netlify etc.)
  const cfgFile = DriveApp.createFile('config.js', configJs, 'application/javascript'); // GitHub repo: docs/config.js

  const msg = 'index.html created in your Google Drive (' + hashes.length + ' vendors, ' + facilities.length + ' facilities).' +
    (noCode.length ? '<br><br>⚠ No access code – cannot sign in: ' + noCode.join(', ') : '') +
    '<br><br>GitHub: <a href="' + cfgFile.getUrl() + '" target="_blank">download config.js</a> → replace docs/config.js in the repo.' +
    '<br>Other hosts: <a href="' + file.getUrl() + '" target="_blank">download index.html</a> (all-in-one).';
  ui.showModalDialog(HtmlService.createHtmlOutput('<div style="font-family:Arial;font-size:14px;line-height:1.5">' + msg + '</div>')
    .setWidth(460).setHeight(220), 'Web page ready');
}

/** salt|name|code → SHA-256, repeated ITERATIONS times. Same steps as the page (CONFIG.iterations). */
function webHash_(salt, name, code) {
  const text = salt + '|' + webNorm_(name) + '|' + String(code).trim().toLowerCase();
  let h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  for (let i = 1; i < WEB.ITERATIONS; i++) h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h);
  return h.map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}

/** Ignore case, spaces and punctuation; keep Thai letters (same rule as the page and the tracker). */
function webNorm_(v) {
  return String(v || '').toLowerCase().replace(/[^a-z0-9\u0E00-\u0E7F]/g, '');
}
