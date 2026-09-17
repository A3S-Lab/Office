/**
 * Freeze docs/latest into docs/<version> for the current tip packaging,
 * rewrite indexes, then cap DOCUMENTATION_VERSIONS.
 *
 * Usage: node scripts/freeze-tip-docs.mjs 0.256.0
 */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const docsRoot = path.join(root, 'docs');
const version = process.argv[2];
if (!/^\d+\.\d+\.\d+$/.test(version ?? '')) {
  throw new Error('Usage: node scripts/freeze-tip-docs.mjs <semver>');
}

function rewriteEnglishIndex(version, text) {
  const nl = text.includes('\r\n') ? '\r\n' : '\n';
  let next = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  next = next.replace(
    /^title:\s*.+$/m,
    `title: A3S Office ${version} documentation`,
  );
  next = next.replace(
    /^# A3S Office(?: \d+\.\d+\.\d+)? documentation\s*\n+[\s\S]*?(?=\n## )/m,
    `# A3S Office ${version} documentation\n\nThis is the frozen documentation index for A3S Office ${version}. Use it to find\nintegration guides, API references, and engineering notes.\n\n`,
  );
  if (!next.includes('frozen documentation index')) {
    throw new Error(`EN freeze rewrite failed for ${version}`);
  }
  return next.replace(/\n/g, nl);
}

function rewriteChineseIndex(version, text) {
  const nl = text.includes('\r\n') ? '\r\n' : '\n';
  let next = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  next = next.replace(/^title:\s*.+$/m, `title: A3S Office ${version} 文档`);
  next = next.replace(
    /^# A3S Office(?: \d+\.\d+\.\d+)? 文档\s*\n+[\s\S]*?(?=\n## )/m,
    `# A3S Office ${version} 文档\n\n这是 A3S Office ${version} 的冻结文档索引，提供本版本可用的 API、接入步骤和工程参考。\n\n`,
  );
  if (
    !next.includes('冻结文档索引') ||
    !next.includes(`# A3S Office ${version} 文档`)
  ) {
    throw new Error(`ZH freeze rewrite failed for ${version}`);
  }
  return next.replace(/\n/g, nl);
}

function stripPlaygroundRows(text) {
  const nl = text.includes('\r\n') ? '\r\n' : '\n';
  const next = text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .split('\n')
    .filter((line) => {
      if (!line.includes('|')) return true;
      return !(
        line.includes('<PlaygroundLink') ||
        line.includes('在线体验') ||
        line.includes('Open the Playground') ||
        line.includes('Product storytelling') ||
        line.includes('product-home demonstrations') ||
        line.includes('产品首页')
      );
    })
    .join('\n');
  return next.replace(/\n/g, nl);
}

const dest = path.join(docsRoot, version);
if (fs.existsSync(dest)) fs.rmSync(dest, { recursive: true, force: true });
fs.cpSync(path.join(docsRoot, 'latest'), dest, { recursive: true });

const en = path.join(dest, 'en', 'index.mdx');
const zh = path.join(dest, 'zh', 'index.mdx');
fs.writeFileSync(
  en,
  stripPlaygroundRows(
    rewriteEnglishIndex(version, fs.readFileSync(en, 'utf8')),
  ),
);
fs.writeFileSync(
  zh,
  stripPlaygroundRows(
    rewriteChineseIndex(version, fs.readFileSync(zh, 'utf8')),
  ),
);

const cap = spawnSync(
  process.execPath,
  ['scripts/cap-published-docs-versions.mjs'],
  {
    cwd: root,
    stdio: 'inherit',
  },
);
if (cap.status !== 0) {
  process.exit(cap.status ?? 1);
}

console.log(`Froze docs/${version} and capped DOCUMENTATION_VERSIONS`);
