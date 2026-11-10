// The Google Form question titles are the contract between three files:
//   Backend_Setup.gs  WEB_FIELDS   (creates the form questions, maps page keys → entry IDs)
//   docs/index.html   sendLines()  (posts one value per page key)
//   Tracker_Sync.gs   collectFormLines_() `need` list (reads the Form Responses columns by title)
import test from 'node:test';
import assert from 'node:assert/strict';
import { read, extractFunction } from './helpers.mjs';

const gs = read('apps-script/vendor-sheet/Backend_Setup.gs');
const WEB_FIELDS = new Function(gs.match(/const WEB_FIELDS = \[[\s\S]*?\];/)[0] + '\nreturn WEB_FIELDS;')();
const pageKeys = [...extractFunction(read('docs/index.html'), 'sendLines').match(/var vals = \{([\s\S]*?)\};/)[1]
  .matchAll(/(\w+):/g)].map(m => m[1]);
const trackerNeed = new Function('return ' + extractFunction(read('apps-script/tracker/Tracker_Sync.gs'), 'collectFormLines_')
  .match(/const need = (\[[\s\S]*?\]);/)[1])();

test('page sends exactly the keys Backend_Setup maps', () => {
  assert.deepEqual([...pageKeys].sort(), WEB_FIELDS.map(f => f[0]).sort());
});

test('tracker reads exactly the question titles Backend_Setup creates (+ Timestamp)', () => {
  assert.deepEqual([...trackerNeed].sort(), ['Timestamp', ...WEB_FIELDS.map(f => f[1])].sort());
});
