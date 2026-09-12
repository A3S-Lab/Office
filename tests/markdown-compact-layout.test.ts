import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { expect, test } from '@rstest/core';

const stylesheet = readFileSync(
  join(process.cwd(), 'src/styles/work-markdown.css'),
  'utf8',
);

function queryBlock(source: string, atRule: string): string {
  const start = source.indexOf(atRule);
  expect(start, atRule).toBeGreaterThanOrEqual(0);
  const open = source.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < source.length; index += 1) {
    if (source[index] === '{') depth += 1;
    if (source[index] === '}') {
      depth -= 1;
      if (depth === 0) return source.slice(open, index + 1);
    }
  }
  throw new Error(`${atRule} is missing a closing brace.`);
}

test('markdown compact layout follows the editor container and fills preview', () => {
  expect(stylesheet).toContain('container-name: work-markdown;');
  expect(stylesheet).toContain('container-type: inline-size;');

  const viewport = queryBlock(stylesheet, '@media (max-width: 640px)');
  const container = queryBlock(
    stylesheet,
    '@container work-markdown (max-width: 640px)',
  );
  expect(container).toBe(viewport);

  for (const block of [viewport, container]) {
    expect(block).toContain('padding: 28px 20px 64px;');
    expect(block).toContain(
      '[data-compact-pane="preview"]\n    .work-markdown-canvas',
    );
    expect(block).toContain('min-height: 100%;');
    expect(block).toContain('margin: 0;');
    expect(block).toContain(
      'background: color-mix(in srgb, var(--a3s-bg) 82%, var(--a3s-panel));',
    );
  }
});
