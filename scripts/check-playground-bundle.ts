import { existsSync, readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { gzipSync } from 'node:zlib';

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const playgroundDist = resolve(repositoryRoot, 'playground-dist');
const indexPath = resolve(playgroundDist, 'index.html');
const maximumInitialJavaScriptGzipBytes = 220 * 1024;

assert(
  existsSync(indexPath),
  'Playground output is missing. Run bun run playground:build first.',
);

const html = readFileSync(indexPath, 'utf8');
const initialScripts = Array.from(
  html.matchAll(/<script\b[^>]*\bsrc=["']([^"']+)["']/g),
  (match) => match[1],
);
assert(initialScripts.length > 0, 'No initial JavaScript assets were found.');

const assets = initialScripts.map((source) => {
  const pathname = new URL(source, 'https://a3s-office.invalid').pathname;
  const staticIndex = pathname.indexOf('/static/');
  assert(staticIndex >= 0, `Unexpected initial script URL: ${source}`);
  const path = resolve(playgroundDist, pathname.slice(staticIndex + 1));
  assert(existsSync(path), `Initial script is missing: ${path}`);
  const sourceBytes = readFileSync(path);
  return {
    source,
    gzipBytes: gzipSync(sourceBytes, { level: 9 }).byteLength,
  };
});

const totalGzipBytes = assets.reduce(
  (total, asset) => total + asset.gzipBytes,
  0,
);
assert(
  totalGzipBytes <= maximumInitialJavaScriptGzipBytes,
  [
    `Initial Playground JavaScript is ${formatKibibytes(totalGzipBytes)} gzip.`,
    `The budget is ${formatKibibytes(maximumInitialJavaScriptGzipBytes)}.`,
    ...assets.map(
      (asset) => `- ${asset.source}: ${formatKibibytes(asset.gzipBytes)} gzip`,
    ),
  ].join('\n'),
);

for (const editorChunk of [
  'document-editor',
  'markdown-editor',
  'spreadsheet-editor',
  'presentation-editor',
  'pdf-viewer',
]) {
  const files = Array.from(
    new Bun.Glob(`0~${editorChunk}.js`).scanSync({
      cwd: resolve(repositoryRoot, 'dist'),
      onlyFiles: true,
    }),
  );
  assert(
    files.length === 1,
    `Expected one independent ${editorChunk} library chunk, found ${files.length}.`,
  );
}

console.log(
  `Initial Playground JavaScript: ${formatKibibytes(totalGzipBytes)} gzip across ${assets.length} files.`,
);
console.log(
  `Performance budget: ${formatKibibytes(maximumInitialJavaScriptGzipBytes)} gzip.`,
);

function formatKibibytes(bytes: number): string {
  return `${(bytes / 1024).toFixed(1)} KiB`;
}

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) throw new Error(message);
}
