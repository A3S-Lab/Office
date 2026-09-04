import { spawnSync } from 'node:child_process';
import { access, cp, mkdir, rm } from 'node:fs/promises';
import { resolve } from 'node:path';

const siteBase = normalizeBase(
  process.env.A3S_OFFICE_SITE_BASE ??
    process.env.A3S_OFFICE_PLAYGROUND_BASE ??
    '/',
);

const environment = { ...process.env };

// Rspress' persistent cache does not include `config.base` in its cache key.
// A cached root deployment bundle therefore contains `base:"/"` even when
// the generated HTML is being published under `/Office/`, which makes the
// client-side router turn the otherwise valid page into a 404 after hydration.
// Keep root builds fast, but make every sub-path build self-contained.
if (siteBase !== '/') {
  environment.RSPRESS_PERSISTENT_CACHE = 'false';
}

const repositoryRoot = resolve(import.meta.dirname, '..');
const siteOutput = resolve(repositoryRoot, 'playground-dist');
const docsOutput = resolve(repositoryRoot, '.docs-build');
const rspress = await resolveRspressBinary();

await rm(docsOutput, { force: true, recursive: true });
await rm(siteOutput, { force: true, recursive: true });

runRspress('website/rspress.config.ts');
runRspress('website/rspress.docs.config.ts');

const docsTarget = resolve(siteOutput, 'docs');
await mkdir(docsTarget, { recursive: true });
await cp(docsOutput, docsTarget, { recursive: true });
await rm(docsOutput, { force: true, recursive: true });

function runRspress(config: string) {
  const result = spawnSync(rspress, ['build', '-c', config], {
    cwd: repositoryRoot,
    env: environment,
    shell: process.platform === 'win32',
    stdio: 'inherit',
  });

  if (result.error) throw result.error;
  if ((result.status ?? 1) !== 0) {
    process.exit(result.status ?? 1);
  }
}

async function resolveRspressBinary(): Promise<string> {
  const binDirectory = resolve(import.meta.dirname, '../node_modules/.bin');
  const candidates =
    process.platform === 'win32'
      ? ['rspress.cmd', 'rspress.exe', 'rspress.bunx']
      : ['rspress'];
  for (const candidate of candidates) {
    const executable = resolve(binDirectory, candidate);
    try {
      await access(executable);
      return executable;
    } catch {
      // Try the next package-manager shim.
    }
  }
  return resolve(
    import.meta.dirname,
    '../node_modules/@rspress/core/bin/rspress.js',
  );
}

function normalizeBase(value: string): string {
  const withLeadingSlash = value.startsWith('/') ? value : `/${value}`;
  return withLeadingSlash.endsWith('/')
    ? withLeadingSlash
    : `${withLeadingSlash}/`;
}
