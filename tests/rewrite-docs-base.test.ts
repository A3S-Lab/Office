import { expect, test } from '@rstest/core';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { rewriteDocsBaseReferences } from '../scripts/rewrite-docs-base';

test('rewrites only root documentation references for a Pages subpath', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'a3s-office-docs-'));
  const htmlPath = path.join(directory, 'index.html');
  try {
    await writeFile(
      htmlPath,
      [
        '<a href="/docs/guide/index.html">Docs</a>',
        '<script src="/Office/docs/static/index.js"></script>',
        'https://a3s-lab.github.io/docs/llms.txt',
      ].join('\n'),
    );

    await expect(
      rewriteDocsBaseReferences(
        directory,
        '/Office/docs/',
        'https://a3s-lab.github.io',
      ),
    ).resolves.toBe(1);

    await expect(readFile(htmlPath, 'utf8')).resolves.toBe(
      [
        '<a href="/Office/docs/guide/index.html">Docs</a>',
        '<script src="/Office/docs/static/index.js"></script>',
        'https://a3s-lab.github.io/Office/docs/llms.txt',
      ].join('\n'),
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test('removes stale Pages prefixes when building root documentation', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'a3s-office-docs-'));
  const scriptPath = path.join(directory, 'runtime.js');
  try {
    await writeFile(
      scriptPath,
      [
        'const base = "/Office/docs/";',
        'const publicPath = "/Preview/docs/static/app.js";',
        'const canonical = "https://a3s-lab.github.io/Office/docs/guide/";',
        'const external = "https://example.com/Preview/docs/guide/";',
      ].join('\n'),
    );

    await expect(
      rewriteDocsBaseReferences(
        directory,
        '/docs/',
        'https://a3s-lab.github.io',
      ),
    ).resolves.toBe(1);

    await expect(readFile(scriptPath, 'utf8')).resolves.toBe(
      [
        'const base = "/docs/";',
        'const publicPath = "/docs/static/app.js";',
        'const canonical = "https://a3s-lab.github.io/docs/guide/";',
        'const external = "https://example.com/Preview/docs/guide/";',
      ].join('\n'),
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test('leaves already normalized documentation references unchanged', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'a3s-office-docs-'));
  const htmlPath = path.join(directory, 'index.html');
  try {
    const source = [
      '<a href="/Office/docs/guide/">Docs</a>',
      '<script src="/Office/docs/static/index.js"></script>',
      'https://a3s-lab.github.io/Office/docs/llms.txt',
    ].join('\n');
    await writeFile(htmlPath, source);

    await expect(
      rewriteDocsBaseReferences(
        directory,
        '/Office/docs/',
        'https://a3s-lab.github.io',
      ),
    ).resolves.toBe(0);

    await expect(readFile(htmlPath, 'utf8')).resolves.toBe(source);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
