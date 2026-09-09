/**
 * Freeze docs/latest from each Office release tag into docs/<version>,
 * rewrite indexes as frozen documentation, and register DOCUMENTATION_VERSIONS.
 *
 * Usage:
 *   node scripts/freeze-docs-versions.mjs
 *   node scripts/freeze-docs-versions.mjs --from 0.61.0 --to 0.94.0
 */
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const docsRoot = path.join(root, 'docs');

function parseArgs(argv) {
  const args = { from: '0.61.0', to: null };
  for (let i = 0; i < argv.length; i += 1) {
    if (argv[i] === '--from') args.from = argv[++i];
    if (argv[i] === '--to') args.to = argv[++i];
  }
  return args;
}

function compareVersion(a, b) {
  const pa = a.split('.').map(Number);
  const pb = b.split('.').map(Number);
  for (let i = 0; i < 3; i += 1) {
    const d = (pa[i] || 0) - (pb[i] || 0);
    if (d) return d;
  }
  return 0;
}

function listVersions(from, to) {
  const tags = execFileSync('git', ['tag', '--list', 'v0.*'], {
    cwd: root,
    encoding: 'utf8',
  })
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((tag) => tag.slice(1));
  return tags
    .filter((version) => {
      try {
        if (compareVersion(version, from) <= 0) return false;
        if (to && compareVersion(version, to) > 0) return false;
        return true;
      } catch {
        return false;
      }
    })
    .sort(compareVersion);
}

function extractLatestFromTag(version) {
  const tag = `v${version}`;
  const dest = path.join(docsRoot, version);
  if (fs.existsSync(dest)) fs.rmSync(dest, { recursive: true, force: true });
  fs.mkdirSync(dest, { recursive: true });
  execFileSync('git', ['cat-file', '-e', `${tag}:docs/latest`], {
    cwd: root,
    stdio: 'pipe',
  });
  const archive = execFileSync(
    'git',
    ['archive', '--format=tar', tag, 'docs/latest'],
    { cwd: root, maxBuffer: 64 * 1024 * 1024 },
  );
  const tmp = path.join(root, `.freeze-tmp-${version}.tar`);
  const extractDir = path.join(root, `.freeze-extract-${version}`);
  fs.writeFileSync(tmp, archive);
  try {
    fs.rmSync(extractDir, { recursive: true, force: true });
    fs.mkdirSync(extractDir, { recursive: true });
    execFileSync('tar', ['-xf', tmp, '-C', extractDir], { cwd: root });
    fs.cpSync(path.join(extractDir, 'docs', 'latest'), dest, {
      recursive: true,
    });
  } finally {
    fs.rmSync(tmp, { force: true });
    fs.rmSync(extractDir, { recursive: true, force: true });
  }
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

function findDonorVersion(version) {
  const older = listFrozenDirs()
    .filter((candidate) => compareVersion(candidate, version) < 0)
    .sort(compareVersion)
    .reverse();
  return older.find((candidate) =>
    fs.existsSync(path.join(docsRoot, candidate, 'en', 'changelog.mdx')),
  );
}

function ensureChangelogRoute(version) {
  for (const lang of ['en', 'zh']) {
    const destDir = path.join(docsRoot, version, lang);
    const changelogPath = path.join(destDir, 'changelog.mdx');
    if (fs.existsSync(changelogPath)) continue;
    const donor = findDonorVersion(version);
    if (!donor) {
      throw new Error(`No changelog donor for ${version}`);
    }
    fs.copyFileSync(
      path.join(docsRoot, donor, lang, 'changelog.mdx'),
      changelogPath,
    );
    const metaPath = path.join(destDir, '_meta.json');
    const meta = JSON.parse(fs.readFileSync(metaPath, 'utf8'));
    if (!meta.includes('changelog')) {
      const indexAt = meta.indexOf('index');
      if (indexAt >= 0) meta.splice(indexAt + 1, 0, 'changelog');
      else meta.unshift('changelog');
      fs.writeFileSync(metaPath, `${JSON.stringify(meta, null, 2)}\n`);
    }
    const navPath = path.join(destDir, '_nav.json');
    const nav = JSON.parse(fs.readFileSync(navPath, 'utf8'));
    if (!nav.some((item) => item.link === '/changelog.html')) {
      const donorNav = JSON.parse(
        fs.readFileSync(path.join(docsRoot, donor, lang, '_nav.json'), 'utf8'),
      );
      const donorItem = donorNav.find(
        (item) => item.link === '/changelog.html',
      );
      if (!donorItem) {
        throw new Error(`Donor ${donor}/${lang} missing changelog nav`);
      }
      nav.splice(1, 0, donorItem);
      fs.writeFileSync(navPath, `${JSON.stringify(nav, null, 2)}\n`);
    }
  }
}

function rewriteIndexes(version) {
  const en = path.join(docsRoot, version, 'en', 'index.mdx');
  const zh = path.join(docsRoot, version, 'zh', 'index.mdx');
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
  ensureChangelogRoute(version);
}

function listFrozenDirs() {
  return fs
    .readdirSync(docsRoot, { withFileTypes: true })
    .filter((d) => d.isDirectory() && /^\d+\.\d+\.\d+$/.test(d.name))
    .map((d) => d.name)
    .sort((a, b) => compareVersion(b, a));
}

/** Keep the Pages/Rspress multi-version build under GitHub-hosted RAM limits. */
const MAX_PUBLISHED_FROZEN_VERSIONS = 10;
/** Frozen trees that visual contracts still deep-link. */
const REQUIRED_PUBLISHED_FROZEN_VERSIONS = ['0.38.0', '0.1.0'];

function registerVersions() {
  const newest = listFrozenDirs().slice(0, MAX_PUBLISHED_FROZEN_VERSIONS);
  const required = REQUIRED_PUBLISHED_FROZEN_VERSIONS.filter((version) =>
    fs.existsSync(path.join(docsRoot, version)),
  );
  const frozen = [
    ...new Set([...newest, ...required]),
  ].sort((a, b) => compareVersion(b, a));
  const merged = ['latest', ...frozen];
  const file = path.join(root, 'website', 'documentation-site.ts');
  const text = fs.readFileSync(file, 'utf8');
  const start = text.indexOf('export const DOCUMENTATION_VERSIONS = [');
  const marker = '] as const;';
  const end = text.indexOf(marker, start);
  if (start < 0 || end < 0) throw new Error('DOCUMENTATION_VERSIONS missing');
  const body = merged.map((v) => `  '${v}',`).join('\n');
  const next =
    text.slice(0, start) +
    `export const DOCUMENTATION_VERSIONS = [\n${body}\n] as const;` +
    text.slice(end + marker.length);
  fs.writeFileSync(file, next);

  const testFile = path.join(
    root,
    'tests',
    'website-documentation-site.test.ts',
  );
  let test = fs.readFileSync(testFile, 'utf8');
  const nl = test.includes('\r\n') ? '\r\n' : '\n';
  const testBody = merged.map((v) => `    '${v}',`).join(nl);
  const replaced = test.replace(
    /expect\(DOCUMENTATION_VERSIONS\)\.toEqual\(\[[\s\S]*?\]\);/,
    `expect(DOCUMENTATION_VERSIONS).toEqual([${nl}${testBody}${nl}  ]);`,
  );
  if (replaced === test) throw new Error('test expectation rewrite failed');
  fs.writeFileSync(testFile, replaced);
  return merged;
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  const to =
    args.to ??
    execFileSync('git', ['describe', '--tags', '--abbrev=0'], {
      cwd: root,
      encoding: 'utf8',
    })
      .trim()
      .replace(/^v/, '');
  const versions = listVersions(args.from, to).filter(
    (version) => !fs.existsSync(path.join(docsRoot, version)),
  );
  console.log(
    `Freezing ${versions.length} missing versions from >${args.from} to ${to}...`,
  );
  for (const version of versions) {
    process.stdout.write(`  ${version}... `);
    extractLatestFromTag(version);
    rewriteIndexes(version);
    console.log('ok');
  }
  // Always re-normalize indexes for newly requested range that already exist.
  for (const version of listVersions(args.from, to)) {
    if (fs.existsSync(path.join(docsRoot, version, 'en', 'index.mdx'))) {
      rewriteIndexes(version);
    }
  }
  const merged = registerVersions();
  const archived = listFrozenDirs().length - (merged.length - 1);
  console.log(
    `Registered ${merged.length - 1} published frozen versions + latest in DOCUMENTATION_VERSIONS` +
      (archived > 0
        ? ` (${archived} older frozen trees kept on disk but not published)`
        : ''),
  );
}

main();
