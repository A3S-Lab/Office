import { readdir, readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { describe, expect, test } from '@rstest/core';

describe('GitHub Actions browser test policy', () => {
  test('keeps A3S Test local and Playwright in CI and Pages', async () => {
    const workflowDirectory = resolve(process.cwd(), '.github/workflows');
    const workflowNames = (await readdir(workflowDirectory)).filter(
      (name) => name.endsWith('.yml') || name.endsWith('.yaml'),
    );
    const workflows = await Promise.all(
      workflowNames.map(async (name) => ({
        name,
        source: await readFile(resolve(workflowDirectory, name), 'utf8'),
      })),
    );
    const actionSource = workflows
      .map(({ name, source }) => `# ${name}\n${source}`)
      .join('\n');

    expect(actionSource).not.toMatch(/a3s[-_ ]?test/i);
    expect(actionSource).not.toMatch(/test:e2e/i);
    for (const workflowName of ['ci.yml', 'pages.yml']) {
      const workflowSource = workflows.find(
        ({ name }) => name === workflowName,
      )?.source;
      expect(workflowSource).toContain('bun run playground:visual');
      expect(workflowSource).toContain('bun scripts/create-e2e-fixtures.ts');
    }
  });
});
