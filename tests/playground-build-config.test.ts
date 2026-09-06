import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { expect, test } from '@rstest/core';

test('copies shared shell assets into the production Playground bundle', async () => {
  const config = await readFile(
    path.resolve(import.meta.dirname, '../playground/rsbuild.config.ts'),
    'utf8',
  );

  expect(config).toContain("from: '../docs/public/a3s-logo.png'");
  expect(config).toContain("to: 'a3s-logo.png'");
  expect(config).toContain("from: '../docs/public/favicon.svg'");
  expect(config).toContain("to: 'favicon.svg'");

  const playground = await readFile(
    path.resolve(import.meta.dirname, '../playground/src/main.tsx'),
    'utf8',
  );
  expect(playground).toContain("navigationHref('/playground/a3s-logo.png')");
});
