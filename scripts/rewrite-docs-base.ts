import { readdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';

const TEXT_EXTENSIONS = new Set([
  '.css',
  '.html',
  '.js',
  '.json',
  '.map',
  '.md',
  '.mjs',
  '.txt',
  '.xml',
]);

const URL_PATH_SEGMENT = '[A-Za-z0-9._~-]+';

const repositoryRoot = path.resolve(import.meta.dirname, '..');
const docsOutput = path.join(repositoryRoot, 'playground-dist', 'docs');
const siteOrigin = process.env.DOCS_ORIGIN ?? 'https://a3s-lab.github.io';
const playgroundBase = normalizeBase(
  process.env.A3S_OFFICE_PLAYGROUND_BASE ?? '/',
);
const targetDocsBase = `${playgroundBase}docs/`;

if (import.meta.main) {
  const changedFiles = await rewriteDocsBaseReferences(
    docsOutput,
    targetDocsBase,
    siteOrigin,
  );
  console.log(
    changedFiles === 0
      ? `Documentation base already matches ${targetDocsBase}`
      : `Rewrote ${changedFiles} documentation file(s) for ${targetDocsBase}`,
  );
}

export async function rewriteDocsBaseReferences(
  outputDirectory: string,
  targetBase: string,
  origin: string,
): Promise<number> {
  const normalizedTargetBase = normalizeBase(targetBase);
  if (!normalizedTargetBase.endsWith('/docs/')) {
    throw new Error(`Documentation base must end with /docs/: ${targetBase}`);
  }

  const files = await collectTextFiles(outputDirectory);
  let changedFiles = 0;
  const normalizedOrigin = origin.replace(/\/$/, '');
  const staleReferences = new Map<string, string[]>();

  for (const file of files) {
    const source = await readFile(file, 'utf8');
    const rewritten = rewriteText(
      source,
      normalizedTargetBase,
      normalizedOrigin,
    );
    if (rewritten !== source) {
      await writeFile(file, rewritten);
      changedFiles += 1;
    }

    const staleBases = findDocumentationBases(
      rewritten,
      normalizedOrigin,
    ).filter((base) => base !== normalizedTargetBase);
    if (staleBases.length > 0) {
      staleReferences.set(path.relative(outputDirectory, file), staleBases);
    }
  }

  if (staleReferences.size > 0) {
    const details = [...staleReferences]
      .map(([file, bases]) => `${file}: ${bases.join(', ')}`)
      .join('; ');
    throw new Error(
      `Documentation output contains stale base reference(s): ${details}`,
    );
  }

  return changedFiles;
}

function rewriteText(
  source: string,
  targetBase: string,
  origin: string,
): string {
  const absoluteBasePattern = createAbsoluteBasePattern(origin);
  const withOrigin = source.replace(
    absoluteBasePattern,
    `${origin}${targetBase}`,
  );
  return withOrigin.replace(
    createLocalBasePattern(),
    (_match, boundary: string, _sourceBase: string) =>
      `${boundary}${targetBase}`,
  );
}

function findDocumentationBases(source: string, origin: string): string[] {
  const bases = new Set<string>();
  for (const match of source.matchAll(createAbsoluteBasePattern(origin))) {
    if (match[1]) bases.add(match[1]);
  }
  for (const match of source.matchAll(createLocalBasePattern())) {
    if (match[2]) bases.add(match[2]);
  }
  return [...bases];
}

function createAbsoluteBasePattern(origin: string): RegExp {
  return new RegExp(
    `${escapeRegExp(origin)}(\\/(?:${URL_PATH_SEGMENT}\\/)*docs\\/)`,
    'g',
  );
}

function createLocalBasePattern(): RegExp {
  return new RegExp(
    `(^|[^A-Za-z0-9._~:/-])(\\/(?:${URL_PATH_SEGMENT}\\/)*docs\\/)`,
    'gm',
  );
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

async function collectTextFiles(directory: string): Promise<string[]> {
  const entries = await readdir(directory, { withFileTypes: true });
  const files: string[] = [];
  for (const entry of entries) {
    const entryPath = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      files.push(...(await collectTextFiles(entryPath)));
    } else if (TEXT_EXTENSIONS.has(path.extname(entry.name))) {
      files.push(entryPath);
    }
  }
  return files;
}

function normalizeBase(value: string): string {
  const withLeadingSlash = value.startsWith('/') ? value : `/${value}`;
  return withLeadingSlash.endsWith('/')
    ? withLeadingSlash
    : `${withLeadingSlash}/`;
}
