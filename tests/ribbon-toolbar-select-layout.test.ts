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
  expect(css).not.toMatch(/\.work-presentation-design-panel\s+select\b/);
  expect(css).not.toMatch(/\.work-presentation-chart-panel\s+select\b/);
});

test('shared ribbon chrome sizes OfficeSelect triggers without native select leftovers', () => {
  const css = readFileSync(join(stylesRoot, 'work-office-chrome.css'), 'utf8');
  expect(css).toContain(
    '.work-office-ribbon .work-office-toolbar .work-office-select',
  );
  expect(css).toMatch(
    /\.work-office-ribbon\s+\.work-office-toolbar\s+\.work-office-select\s*>\s*button\[role="combobox"\]/,
  );
  expect(css).not.toContain('.work-office-ribbon .work-office-toolbar select');
  expect(css).toContain(
    '.work-office-ribbon .work-office-toolbar .work-color-tool',
  );
});

test('Spreadsheet ribbon gives named OfficeSelects dedicated widths without a 62px crush default', () => {
  const css = readFileSync(join(stylesRoot, 'work-spreadsheet.css'), 'utf8');
  expect(css).toMatch(/\.work-office-select\.work-spreadsheet-font-size/);
  expect(css).toMatch(
    /\.work-spreadsheet-ribbon-toolbar\s+\.work-office-select\s*\{[^}]*width:\s*auto/s,
  );
  expect(css).not.toMatch(
    /\.work-spreadsheet-ribbon-toolbar\s+\.work-office-select\s*\{[^}]*width:\s*62px/s,
  );
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

test('Document page-setup ribbon choices are wide enough for Chinese labels', () => {
  const css = readFileSync(
    join(stylesRoot, 'work-document-controls.css'),
    'utf8',
  );
  expect(css).toContain('.work-document-page-setup-choice');
  expect(css).toMatch(
    /\.work-document-ribbon[\s\S]*\.work-document-page-setup-choice[\s\S]*width:\s*118px/m,
  );
  expect(css).toMatch(
    /\.work-document-page-setup-choice[\s\S]*button\[role="combobox"\][\s\S]*height:\s*24px/m,
  );
});

test('Spreadsheet border settings use OfficeSelect and OfficeColorPicker, not native select', () => {
  const css = readFileSync(
    join(stylesRoot, 'work-spreadsheet-chrome.css'),
    'utf8',
  );
  expect(css).toContain(
    '.work-spreadsheet-border-settings .work-office-select',
  );
  expect(css).toContain(
    '.work-spreadsheet-border-settings .work-office-color-picker',
  );
  expect(css).not.toMatch(/\.work-spreadsheet-border-settings\s+select\b/);
});

test('Document contextual ribbon selects keep fixed widths and 24px table comboboxes', () => {
  const css = readFileSync(
    join(stylesRoot, 'work-document-controls.css'),
    'utf8',
  );
  expect(css).toContain('.work-document-connector-kind-select');
  expect(css).toContain('.work-document-field-insert-select');
  expect(css).toContain('.work-document-picture-wrap-distance-select');
  expect(css).toMatch(
    /\.work-document-ribbon\s+\.work-document-connector-width-select\s*\{[^}]*width:\s*118px/s,
  );
  expect(css).toMatch(
    /\.work-document-ribbon\s+\.work-document-text-box-border-width-select\s*\{[^}]*width:\s*118px/s,
  );
  expect(css).toMatch(
    /\.work-document-ribbon\s+\.work-document-line-height-select\s*\{[^}]*width:\s*96px/s,
  );
  expect(css).toMatch(
    /\.work-document-table-border-target-select\s*\{[^}]*width:\s*110px/s,
  );
  expect(css).toMatch(
    /\.work-document-table-border-select\s*>\s*button\[role="combobox"\]\s*\{[^}]*height:\s*24px/s,
  );
  expect(css).toMatch(
    /\.work-document-table-layout-select\s*>\s*button\[role="combobox"\]\s*\{[^}]*height:\s*24px/s,
  );
});

test('shared editor toolbar no longer styles native select', () => {
  const css = readFileSync(join(stylesRoot, 'work-editor.css'), 'utf8');
  expect(css).not.toMatch(/\.work-office-toolbar\s+select\s*\{/);
});
