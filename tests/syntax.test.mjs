// Every script must at least parse: .gs files (Apps Script V8 = modern JS) and the page's inline script.
import test from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { writeFileSync, mkdtempSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { ROOT, read } from './helpers.mjs';

const tmp = mkdtempSync(join(tmpdir(), 'syntax-'));
const check = (name, code) => {
  const f = join(tmp, name.replace(/[\/\\]/g, '_') + '.js');
  writeFileSync(f, code);
  execFileSync(process.execPath, ['--check', f], { stdio: 'pipe' });
};
const gsFiles = ['apps-script/vendor-sheet', 'apps-script/vendor-sheet/legacy-webapp', 'apps-script/tracker']
  .flatMap(d => readdirSync(join(ROOT, d)).filter(f => f.endsWith('.gs')).map(f => `${d}/${f}`));

for (const f of gsFiles) test(`parses: ${f}`, () => assert.doesNotThrow(() => check(f, read(f))));
for (const f of ['docs/index.html', 'apps-script/vendor-sheet/WebPage.html', 'apps-script/vendor-sheet/legacy-webapp/Form.html']) {
  test(`parses inline script: ${f}`, () => {
    const scripts = [...read(f).matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m => m[1]);
    assert.ok(scripts.length > 0);
    scripts.forEach((s, i) => assert.doesNotThrow(() => check(`${f}#${i}`, s)));
  });
}
test('docs/index.html and WebPage.html stay identical apart from the config.js tag', () => {
  const web = read('apps-script/vendor-sheet/WebPage.html');
  assert.ok(!web.includes('config.js'), 'WebPage.html must not reference config.js');
  const strip = s => s.split('\n').filter(l => !l.includes('config.js')).join('\n');
  assert.equal(strip(read('docs/index.html')), strip(web));
});
