import { expect, test } from '@rstest/core';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const stylesRoot = join(
  dirname(fileURLToPath(import.meta.url)),
  '../src/styles',
);

test('Presentation ribbon does not crush the shared 74px toolbar height', () => {
  const css = readFileSync(join(stylesRoot, 'work-presentation.css'), 'utf8');
  expect(css).not.toMatch(
    /\.presentation-toolbar\s*\{[^}]*\bheight:\s*44px\b/s,
  );
  expect(css).not.toMatch(
    /\.presentation-toolbar\s*\{[^}]*\bflex-basis:\s*44px\b/s,
  );
  expect(css).toContain('.presentation-font-family-select');
  expect(css).toContain('.presentation-align-select');
  expect(css).toContain(
    '.work-presentation-design-controls .work-office-select',
  );
});

test('shared ribbon chrome sizes OfficeSelect triggers, not only native select', () => {
  const css = readFileSync(join(stylesRoot, 'work-office-chrome.css'), 'utf8');
  expect(css).toContain(
    '.work-office-ribbon .work-office-toolbar .work-office-select',
  );
  expect(css).toMatch(
    /\.work-office-ribbon\s+\.work-office-toolbar\s+\.work-office-select\s*>\s*button\[role="combobox"\]/,
  );
  expect(css).toContain('.work-office-ribbon .work-office-toolbar select');
  expect(css).toContain(
    '.work-office-ribbon .work-office-toolbar .work-color-tool',
  );
});

test('Spreadsheet ribbon gives font-size selects a dedicated width', () => {
  const css = readFileSync(join(stylesRoot, 'work-spreadsheet.css'), 'utf8');
  expect(css).toMatch(/\.work-office-select\.work-spreadsheet-font-size/);
});

test('Markdown ribbon gives the paragraph-style select a fixed ribbon width', () => {
  const css = readFileSync(join(stylesRoot, 'work-markdown.css'), 'utf8');
  expect(css).toContain('.markdown-paragraph-style-select');
  expect(css).toContain('.markdown-toolbar .work-office-select');
});

test('ported OfficeSelect menus keep floating top/left instead of trigger-relative offset', () => {
  const css = readFileSync(
    join(stylesRoot, 'work-office-controls.css'),
    'utf8',
  );
  expect(css).toMatch(
    /\.work-office-select-menu:not\(\[data-floating="true"\]\)/,
  );
  expect(css).toMatch(/\.work-office-select-menu\[data-floating="true"\]/);
});
