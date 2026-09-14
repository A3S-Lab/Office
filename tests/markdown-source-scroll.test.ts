import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { expect, test } from '@rstest/core';

const stylesheet = readFileSync(
  join(process.cwd(), 'src/styles/work-markdown.css'),
  'utf8',
);

function ruleBlock(source: string, selector: string): string {
  const start = source.indexOf(selector);
  expect(start, selector).toBeGreaterThanOrEqual(0);
  const open = source.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < source.length; index += 1) {
    if (source[index] === '{') depth += 1;
    if (source[index] === '}') {
      depth -= 1;
      if (depth === 0) return source.slice(open, index + 1);
    }
  }
  throw new Error(`${selector} is missing a closing brace.`);
}

test('Markdown source pane scrolls inside the textarea, not the pane shell', () => {
  const pane = ruleBlock(stylesheet, '.work-markdown-pane.source {');
  expect(pane).toContain('overflow: hidden;');

  const textarea = ruleBlock(stylesheet, '.work-markdown-pane.source textarea {');
  expect(textarea).toContain('min-height: 0;');
  expect(textarea).toContain('overflow: auto;');
  expect(textarea).not.toContain('min-height: 100%;');

  const splitTextarea = ruleBlock(
    stylesheet,
    '.work-markdown-workspace.split .work-markdown-pane.source textarea {',
  );
  expect(splitTextarea).not.toContain('min-height: calc(100% - 32px)');
  expect(splitTextarea).not.toContain('min-height: 100%;');
});
