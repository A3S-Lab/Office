import { readFileSync, writeFileSync } from 'node:fs';

const files = [
  'src/internal/features/work/editors/document-font-dialog.tsx',
  'src/internal/features/work/editors/document-font-dialog-run-border-section.tsx',
  'src/internal/features/work/editors/document-font-dialog-run-shading-section.tsx',
];

const strings = new Set();
for (const file of files) {
  const src = readFileSync(file, 'utf8');
  for (const match of src.matchAll(/['"`]([^'"`]*[\u4e00-\u9fff][^'"`]*)['"`]/g)) {
    strings.add(match[1]);
  }
}

const sorted = [...strings].sort((a, b) => a.localeCompare(b, 'zh'));
writeFileSync('/tmp/font-dialog-cjk.json', JSON.stringify(sorted, null, 2));
console.log('count', sorted.length);
for (const s of sorted.slice(0, 80)) console.log(s);
