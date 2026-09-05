import { existsSync } from 'node:fs';
import path from 'node:path';

const repositoryRoot = path.resolve(import.meta.dirname, '..');
const playgroundEntry = path.join(
  repositoryRoot,
  'playground-dist',
  'playground',
  'index.html',
);

if (!existsSync(playgroundEntry)) {
  console.error(
    [
      'The static Playground bundle is missing.',
      `Expected: ${path.relative(repositoryRoot, playgroundEntry)}`,
      'Run `bun run playground:build:ui` (or `bun run playground:build`) before starting the preview.',
    ].join('\n'),
  );
  process.exit(1);
}

console.log(
  `Playground bundle ready: ${path.relative(repositoryRoot, playgroundEntry)}`,
);
