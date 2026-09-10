/**
 * Cap DOCUMENTATION_VERSIONS to latest + newest frozen trees so Pages builds fit.
 * Older frozen trees remain on disk for archaeology; they are just not published.
 * Required pins stay published for visual deep-links.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const docsRoot = path.join(root, 'docs');
/** Newest continuous freezes published in the version selector / Pages site. */
const MAX_PUBLISHED_FROZEN_VERSIONS = 20;
/** Frozen trees that visual contracts still deep-link. */
const REQUIRED_PUBLISHED_FROZEN_VERSIONS = ['0.38.0', '0.1.0'];

function compareVersion(a, b) {
  const pa = a.split('.').map(Number);
  const pb = b.split('.').map(Number);
  for (let i = 0; i < 3; i += 1) {
    const delta = (pa[i] || 0) - (pb[i] || 0);
    if (delta !== 0) return delta;
  }
  return 0;
}

const newest = fs
  .readdirSync(docsRoot, { withFileTypes: true })
  .filter((d) => d.isDirectory() && /^\d+\.\d+\.\d+$/.test(d.name))
  .map((d) => d.name)
  .sort((a, b) => compareVersion(b, a))
  .slice(0, MAX_PUBLISHED_FROZEN_VERSIONS);
const required = REQUIRED_PUBLISHED_FROZEN_VERSIONS.filter((version) =>
  fs.existsSync(path.join(docsRoot, version)),
);
const frozen = [...new Set([...newest, ...required])].sort((a, b) =>
  compareVersion(b, a),
);
const merged = ['latest', ...frozen];

const file = path.join(root, 'website', 'documentation-site.ts');
const text = fs.readFileSync(file, 'utf8');
const start = text.indexOf('export const DOCUMENTATION_VERSIONS = [');
const marker = '] as const;';
const end = text.indexOf(marker, start);
if (start < 0 || end < 0) {
  throw new Error('DOCUMENTATION_VERSIONS missing');
}
const body = merged.map((v) => `  '${v}',`).join('\n');
fs.writeFileSync(
  file,
  `${text.slice(0, start)}export const DOCUMENTATION_VERSIONS = [\n${body}\n] as const;${text.slice(end + marker.length)}`,
);

const testFile = path.join(root, 'tests', 'website-documentation-site.test.ts');
let test = fs.readFileSync(testFile, 'utf8');
const nl = test.includes('\r\n') ? '\r\n' : '\n';
const testBody = merged.map((v) => `    '${v}',`).join(nl);
const replaced = test.replace(
  /expect\(DOCUMENTATION_VERSIONS\)\.toEqual\(\[[\s\S]*?\]\);/,
  `expect(DOCUMENTATION_VERSIONS).toEqual([${nl}${testBody}${nl}  ]);`,
);
if (replaced === test) {
  throw new Error('test expectation rewrite failed');
}
fs.writeFileSync(testFile, replaced);

console.log(
  `Published ${merged.length} versions (latest + ${frozen.length} frozen: ${frozen[0]} … ${frozen.at(-1)}; required ${required.join(',') || 'none'})`,
);
