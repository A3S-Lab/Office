import { spawnSync } from 'node:child_process';
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

const rspress = resolve(
  import.meta.dirname,
  '../node_modules/.bin',
  process.platform === 'win32' ? 'rspress.cmd' : 'rspress',
);
const result = spawnSync(
  rspress,
  ['build', '-c', 'website/rspress.config.ts'],
  {
    env: environment,
    stdio: 'inherit',
  },
);

if (result.error) throw result.error;
process.exit(result.status ?? 1);

function normalizeBase(value: string): string {
  const withLeadingSlash = value.startsWith('/') ? value : `/${value}`;
  return withLeadingSlash.endsWith('/')
    ? withLeadingSlash
    : `${withLeadingSlash}/`;
}
