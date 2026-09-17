import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { expect, test } from '@rstest/core';

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
  expect(css).toContain('.presentation-align-menu');
  expect(css).toContain('.presentation-align-trigger');
  expect(css).not.toContain('.presentation-align-select');
  expect(css).toContain(
    '.work-presentation-design-controls .work-office-select',
  );
  expect(css).not.toMatch(/\.work-presentation-design-panel\s+select\b/);
  expect(css).not.toMatch(/\.work-presentation-chart-panel\s+select\b/);
});

test('Presentation align-to-slide stays executable for a single selection unit', () => {
  const source = readFileSync(
    join(
      dirname(fileURLToPath(import.meta.url)),
      '../src/internal/features/work/editors/presentation-editor.tsx',
    ),
    'utf8',
  );
  expect(source).toContain('canAlignElement: selectionUnits.length >= 1');
  expect(source).not.toContain('canAlignElement: selectionUnits.length >= 2');
});

test('Document column select closed value is the current count, not 更多分栏', () => {
  const source = readFileSync(
    join(
      dirname(fileURLToPath(import.meta.url)),
      '../src/internal/features/work/editors/document-page-layout-ribbon.tsx',
    ),
    'utf8',
  );
  expect(source).toContain('function documentColumnClosedValue');
  expect(source).toContain("{ value: '4', label: '四栏' }");
  expect(source).toContain(
    'const columnPreset = documentColumnClosedValue(layout.columns.count)',
  );
  expect(source).toContain("{ value: 'more', label: '更多分栏' }");
});

test('Document field insert is a command menu instead of a fake OfficeSelect', () => {
  const source = readFileSync(
    join(
      dirname(fileURLToPath(import.meta.url)),
      '../src/internal/features/work/editors/document-toolbar.tsx',
    ),
    'utf8',
  );
  expect(source).toContain('function DocumentFieldInsertMenu');
  expect(source).toContain('panelRole="menu"');
  expect(source).toContain('role="menuitem"');
  expect(source).toContain('className="work-document-field-insert-menu"');
  expect(source).not.toContain('className="work-document-field-insert-select"');
  expect(source).not.toContain("value: '', label: '插入域…'");
  expect(source).not.toContain('function DocumentFieldSelect');
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
  expect(css).toMatch(
    /\.work-office-select\.work-spreadsheet-number-format[\s\S]*?width:\s*112px/s,
  );
  expect(css).toMatch(
    /\.work-spreadsheet-date-time-trigger\.with-label\s*>\s*span\s*\{[^}]*max-width:\s*84px/s,
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

test('Spreadsheet row/column header overlays stay translucent so canvas labels remain visible', () => {
  const css = readFileSync(
    join(stylesRoot, 'work-spreadsheet-chrome.css'),
    'utf8',
  );
  expect(css).toMatch(
    /\.fortune-row-header-hover[\s\S]*?background:\s*color-mix\(\s*in srgb,\s*var\(--work-spreadsheet-accent\)\s+12%,\s*transparent/s,
  );
  expect(css).toMatch(
    /\.fortune-row-header-selected[\s\S]*?background:\s*color-mix\(\s*in srgb,\s*var\(--work-spreadsheet-accent\)\s+16%,\s*transparent/s,
  );
  expect(css).not.toMatch(
    /\.fortune-row-header-hover\s*\{[^}]*background:\s*var\(--work-spreadsheet-chrome-hover\)/s,
  );
  expect(css).not.toMatch(
    /\.fortune-row-header-selected\s*\{[^}]*background:\s*var\(--work-spreadsheet-chrome-active\)/s,
  );
});

test('Spreadsheet in-cell editor defaults to vertical middle alignment', () => {
  const css = readFileSync(
    join(stylesRoot, 'work-spreadsheet-chrome.css'),
    'utf8',
  );
  expect(css).toMatch(
    /\.luckysheet-input-box-inner\s*\{[^}]*display:\s*flex[^}]*align-items:\s*center/s,
  );
  expect(css).toContain(
    '.work-spreadsheet-editor[data-cell-vt="1"] .luckysheet-input-box-inner',
  );
  expect(css).toContain(
    '.work-spreadsheet-editor[data-cell-vt="2"] .luckysheet-input-box-inner',
  );
  expect(css).toMatch(
    /\.luckysheet-cell-input\s*\{[^}]*height:\s*auto\s*!important/s,
  );

  const alignmentRibbon = readFileSync(
    join(
      process.cwd(),
      'src/internal/features/work/editors/spreadsheet-alignment-ribbon.tsx',
    ),
    'utf8',
  );
  expect(alignmentRibbon).toContain('Number(toolbarCell?.vt ?? 0) === 0');
  expect(alignmentRibbon).not.toContain('Number(toolbarCell?.vt ?? 1) === 1');

  const editor = readFileSync(
    join(
      process.cwd(),
      'src/internal/features/work/editors/spreadsheet-editor.tsx',
    ),
    'utf8',
  );
  expect(editor).toContain('data-cell-vt={');
});

test('Document contextual ribbon selects keep fixed widths and 24px table comboboxes', () => {
  const css = readFileSync(
    join(stylesRoot, 'work-document-controls.css'),
    'utf8',
  );
  expect(css).toContain('.work-document-connector-kind-select');
  expect(css).toContain('.work-document-field-insert-menu');
  expect(css).toContain('.work-document-field-insert-trigger');
  expect(css).not.toContain('.work-document-field-insert-select');
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
    /\.work-document-table-border-target-select\s*\{[^}]*width:\s*128px/s,
  );
  expect(css).toMatch(
    /\.work-document-table-border-select\s*>\s*button\[role="combobox"\]\s*\{[^}]*height:\s*24px/s,
  );
  expect(css).toMatch(
    /\.work-document-table-layout-select\s*\{[^}]*width:\s*104px/s,
  );
  expect(css).toMatch(
    /\.work-document-table-layout-select\s*>\s*button\[role="combobox"\]\s*\{[^}]*height:\s*24px/s,
  );
});

test('Presentation animation trigger select fits long Chinese closed labels', () => {
  const css = readFileSync(
    join(stylesRoot, 'work-presentation-animations.css'),
    'utf8',
  );
  expect(css).toMatch(
    /\.work-presentation-animation-options\s+\.work-office-field\.trigger\s*\{[^}]*width:\s*168px/s,
  );
});

test('Spreadsheet sort levels no longer style orphaned native select', () => {
  const css = readFileSync(
    join(stylesRoot, 'work-spreadsheet-sort.css'),
    'utf8',
  );
  expect(css).not.toMatch(/\.work-spreadsheet-sort-level\s+select\b/);
  expect(css).toContain('.work-spreadsheet-sort-custom-list-manager select');
});

test('shared editor toolbar no longer styles native select', () => {
  const css = readFileSync(join(stylesRoot, 'work-editor.css'), 'utf8');
  expect(css).not.toMatch(/\.work-office-toolbar\s+select\s*\{/);
});

test('Dialog surface keeps overflow visible so portaled OfficeSelect menus are not clipped', () => {
  const css = readFileSync(join(stylesRoot, 'design-system-forms.css'), 'utf8');
  expect(css).toMatch(/\.ds-dialog\s*\{[^}]*overflow:\s*visible/s);
  expect(css).not.toMatch(/\.ds-dialog\s*\{[^}]*overflow:\s*hidden/s);
  expect(css).toMatch(/\.ds-dialog-body\s*\{[^}]*overflow:\s*auto/s);
  expect(css).toMatch(/\.ds-dialog-body\s*\{[^}]*flex:\s*1\s+1\s+auto/s);
});
