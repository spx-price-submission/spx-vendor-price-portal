// The access-code hash is computed in two places and MUST match exactly:
//   docs/index.html                      vendorHash()   – browser, Web Crypto
//   apps-script/vendor-sheet/Backend_Setup.gs  webHash_() – Apps Script, Utilities.computeDigest (signed bytes)
// and the name rule must match Tracker_Sync.gs normName_() (server-side re-check).
import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash, webcrypto } from 'node:crypto';
import { read, extractFunction } from './helpers.mjs';

globalThis.crypto ??= webcrypto;
const html = read('docs/index.html');
const gs = read('apps-script/vendor-sheet/Backend_Setup.gs');
const tracker = read('apps-script/tracker/Tracker_Sync.gs');
const ITER = Number(gs.match(/ITERATIONS:\s*(\d+)/)[1]);

const browserHash = (salt) => new Function('CONFIG',
  extractFunction(html, 'normName') + '\n' + extractFunction(html, 'vendorHash') + '\nreturn vendorHash;')({ salt, iterations: ITER });

const webNorm = new Function(extractFunction(gs, 'webNorm_') + '\nreturn webNorm_;')();
const trackerNorm = new Function(extractFunction(tracker, 'normName_') + '\nreturn normName_;')();

// Emulates Utilities.computeDigest: returns signed bytes, accepts string (UTF-8) or signed byte array
const computeDigest = (val) => [...createHash('sha256')
  .update(typeof val === 'string' ? Buffer.from(val, 'utf8') : Buffer.from(val.map(b => (b + 256) % 256)))
  .digest()].map(b => (b > 127 ? b - 256 : b));
function appsScriptHash(salt, name, code) {
  let h = computeDigest(salt + '|' + webNorm(name) + '|' + String(code).trim().toLowerCase());
  for (let i = 1; i < ITER; i++) h = computeDigest(h);
  return h.map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}

const cases = [
  // [name in Vendors tab, code in Vendors tab, what the vendor types, code the vendor types]
  ['WAS Add', 'A3F09C11B2E4', 'was add', ' a3f09c11b2e4 '],
  ['APLUS', 'ZX81QQ70PL0A', 'A Plus', 'zx81qq70pl0a'],
  ['บริษัท เอบีซี', 'K2K2K2K2K2K2', 'บริษัทเอบีซี', 'k2k2k2k2k2k2'],
];

test('browser hash equals Apps Script hash (incl. typed variations)', async () => {
  const salt = 'test-salt-1234';
  const vendorHash = browserHash(salt);
  for (const [name, code, typedName, typedCode] of cases) {
    assert.equal(await vendorHash(typedName, typedCode), appsScriptHash(salt, name, code), `mismatch for ${name}`);
  }
});

test('wrong code does not match', async () => {
  const salt = 'test-salt-1234';
  assert.notEqual(await browserHash(salt)('APLUS', 'WRONGCODE000'), appsScriptHash(salt, 'APLUS', 'ZX81QQ70PL0A'));
});

test('name normalisation is identical in page, Backend_Setup and Tracker_Sync', () => {
  const pageNorm = new Function(extractFunction(html, 'normName') + '\nreturn normName;')();
  for (const s of ['W.A.S  Add', 'APlus', 'บริษัท ABC (1)', ' R89 ']) {
    assert.equal(pageNorm(s), webNorm(s));
    assert.equal(webNorm(s), trackerNorm(s));
  }
});
