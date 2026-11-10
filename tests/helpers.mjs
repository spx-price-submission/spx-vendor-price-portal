import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

export const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
export const read = (p) => readFileSync(join(ROOT, p), 'utf8');

/** Pulls one top-level `function name(...) {...}` out of a source file (2-space or 0 indent). */
export function extractFunction(src, name) {
  const start = src.search(new RegExp(`function ${name}\\s*\\(`));
  if (start < 0) throw new Error(`function ${name} not found`);
  let i = src.indexOf('{', start), depth = 0;
  for (; i < src.length; i++) {
    if (src[i] === '{') depth++;
    else if (src[i] === '}' && --depth === 0) return src.slice(start, i + 1);
  }
  throw new Error(`function ${name} not closed`);
}

/** Loads docs/config.js and returns CONFIG. */
export function loadConfig() {
  return new Function(read('docs/config.js') + '\nreturn CONFIG;')();
}
