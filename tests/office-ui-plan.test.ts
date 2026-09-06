import path from 'node:path';
import { expect, test } from '@rstest/core';
import {
  createSurfacePlan,
  matrix,
  resolveSurface,
} from '../scripts/office-ui-ops-core';

test('derives a WPS-aware Writer workflow from the declarative matrix', () => {
  const plan = createSurfacePlan(resolveSurface('writer'));
  const commandIds = plan.commands.map((command) => command.id);

  expect(plan.surface).toMatchObject({
    id: 'writer',
    kind: 'document',
    wpsReference: true,
    acl: expect.arrayContaining([
      'tests/e2e/word-wps-shortcuts.acl',
      'tests/e2e/word-wps-layout-parity.acl',
      'tests/e2e/word-wps-font-grid-parity.acl',
    ]),
    fixtures: expect.arrayContaining([
      'word-wps-layout.docx',
      'word-wps-font-matrix.docx',
      'word-wps-cjk-font-matrix.docx',
      'word-wps-grid-matrix.docx',
      'word-wps-script-matrix.docx',
    ]),
  });
  expect(commandIds).toEqual([
    'fixtures',
    'check',
    'gate',
    ...writerSurfaceVisualCommandIds(plan.surface.visual.length),
    'agent',
    'wps-ui',
    'wps-fields',
  ]);
  expect(plan.commands.find((command) => command.id === 'gate')).toMatchObject({
    executable: 'bun',
    args: expect.arrayContaining([
      'gate',
      'writer',
      '--run',
      '--browser-driver',
      'standalone',
    ]),
  });
  expect(
    plan.commands.find((command) => command.id === 'fixtures'),
  ).toMatchObject({
    executable: 'bun',
    args: ['run', 'test:e2e:fixtures'],
  });
  expect(
    plan.commands.find((command) => command.id === 'visual-1'),
  ).toMatchObject({
    args: expect.arrayContaining(['--project', 'desktop-1280', 'compact-768']),
  });
});

function writerSurfaceVisualCommandIds(count: number): string[] {
  return Array.from({ length: count }, (_, index) => `visual-${index + 1}`);
}

test('keeps all five editor plans aligned with the shared evidence root', () => {
  const plans = matrix.surfaces.map(createSurfacePlan);
  expect(plans).toHaveLength(5);
  expect(plans.map((plan) => plan.surface.id)).toEqual([
    'writer',
    'spreadsheet',
    'presentation',
    'markdown',
    'pdf',
  ]);
  expect(
    plans.every((plan) =>
      path
        .resolve(plan.evidenceRoot)
        .startsWith(path.resolve(matrix.shared.evidenceRoot)),
    ),
  ).toBe(true);
  expect(
    plans
      .filter((plan) => plan.surface.id !== 'writer')
      .every(
        (plan) => !plan.commands.some((command) => command.id === 'wps-ui'),
      ),
  ).toBe(true);
});

test('keeps the focused WPS parity matrix broad across every editor surface', () => {
  const plans = matrix.surfaces.map(createSurfacePlan);
  expect(
    plans.find((plan) => plan.surface.id === 'writer')?.surface.visual,
  ).toEqual(
    expect.arrayContaining([
      'visual-tests/document-formatting.functional.spec.ts',
      'visual-tests/document-page-navigation.functional.spec.ts',
      'visual-tests/document-character-position.functional.spec.ts',
      'visual-tests/document-character-spacing.functional.spec.ts',
      'visual-tests/document-opentype.functional.spec.ts',
      'visual-tests/document-hidden-text.functional.spec.ts',
      'visual-tests/document-content-control.functional.spec.ts',
    ]),
  );
  expect(
    plans.find((plan) => plan.surface.id === 'spreadsheet')?.surface.acl,
  ).toEqual(
    expect.arrayContaining([
      'tests/e2e/spreadsheet-ribbon-alignment.acl',
      'tests/e2e/spreadsheet-font-dialog-shortcuts.acl',
      'tests/e2e/spreadsheet-auto-sum.acl',
      'tests/e2e/spreadsheet-paste-special.acl',
      'tests/e2e/spreadsheet-conditional-format.acl',
      'tests/e2e/spreadsheet-date-time.acl',
      'tests/e2e/spreadsheet-cell-style.acl',
      'tests/e2e/spreadsheet-rich-text.acl',
      'tests/e2e/spreadsheet-table-totals.acl',
      'tests/e2e/spreadsheet-ribbon-orientation-visibility.acl',
      'tests/e2e/spreadsheet-appearance-sort.acl',
      'tests/e2e/spreadsheet-custom-list-sort.acl',
      'tests/e2e/spreadsheet-owned-range-sort.acl',
      'tests/e2e/spreadsheet-row-sort.acl',
      'tests/e2e/spreadsheet-sort-range.acl',
      'tests/e2e/spreadsheet-text-sort.acl',
      'tests/e2e/spreadsheet-diagonal-borders.acl',
      'tests/e2e/spreadsheet-gradient-fill.acl',
      'tests/e2e/spreadsheet-pattern-fill.acl',
    ]),
  );
  expect(
    plans.find((plan) => plan.surface.id === 'writer')?.surface.acl,
  ).toEqual(
    expect.arrayContaining([
      'tests/e2e/word-character-position.acl',
      'tests/e2e/word-character-spacing.acl',
      'tests/e2e/word-emphasis.acl',
      'tests/e2e/word-hidden-text.acl',
      'tests/e2e/word-opentype-typography.acl',
      'tests/e2e/word-review-conflict-phone.acl',
      'tests/e2e/word-content-controls-phone.acl',
      'tests/e2e/word-move-revision.acl',
    ]),
  );
  expect(
    plans.find((plan) => plan.surface.id === 'spreadsheet')?.surface.visual,
  ).toEqual(
    expect.arrayContaining([
      'visual-tests/spreadsheet-auto-sum.functional.spec.ts',
      'visual-tests/spreadsheet-format-cells.functional.spec.ts',
      'visual-tests/spreadsheet-paste-special.functional.spec.ts',
      'visual-tests/spreadsheet-rich-text.functional.spec.ts',
      'visual-tests/spreadsheet-appearance-sort.functional.spec.ts',
      'visual-tests/spreadsheet-custom-list-sort.functional.spec.ts',
      'visual-tests/spreadsheet-owned-range-sort.functional.spec.ts',
      'visual-tests/spreadsheet-row-sort.functional.spec.ts',
      'visual-tests/spreadsheet-sort-range.functional.spec.ts',
      'visual-tests/spreadsheet-text-sort.functional.spec.ts',
      'visual-tests/spreadsheet-diagonal-borders.functional.spec.ts',
      'visual-tests/spreadsheet-gradient-fill.functional.spec.ts',
      'visual-tests/spreadsheet-pattern-fill.functional.spec.ts',
    ]),
  );
  expect(
    plans.find((plan) => plan.surface.id === 'presentation')?.surface.acl,
  ).toEqual(
    expect.arrayContaining([
      'tests/e2e/presentation-cut-paste-focus.acl',
      'tests/e2e/presentation-chinese-ime.acl',
      'tests/e2e/presentation-large-windowing.acl',
      'tests/e2e/presentation-phone-chart-pane.acl',
      'tests/e2e/presentation-phone-comments.acl',
    ]),
  );
  expect(plans.find((plan) => plan.surface.id === 'pdf')?.surface.acl).toEqual(
    expect.arrayContaining([
      'tests/e2e/pdf-thumbnail-keyboard.acl',
      'tests/e2e/pdf-page-drawer-phone.acl',
      'tests/e2e/pdf-large-windowing.acl',
    ]),
  );
  expect(
    plans.find((plan) => plan.surface.id === 'presentation')?.surface.visual,
  ).toEqual(
    expect.arrayContaining([
      'visual-tests/presentation-phone-chart-pane.functional.spec.ts',
      'visual-tests/presentation-phone-comments.functional.spec.ts',
    ]),
  );
});

test('keeps the standalone browser plan on the typed CDP adapter path', () => {
  const writerSurface = matrix.surfaces.find(
    (surface) => surface.id === 'writer',
  );
  if (!writerSurface) throw new Error('Writer surface is missing from matrix');
  const writerPlan = createSurfacePlan(writerSurface);
  const gate = writerPlan.commands.find((command) => command.id === 'gate');
  const agent = writerPlan.commands.find((command) => command.id === 'agent');
  expect(gate?.args).toEqual(
    expect.arrayContaining([
      '--browser-driver',
      'standalone',
      '--cdp-port',
      '9345',
    ]),
  );
  expect(agent?.args).toEqual(expect.arrayContaining(['--cdp-port', '9345']));
  expect(gate?.args).not.toContain('scripts/a3s-test-cdp-browser.cmd');
  expect(agent?.args).not.toContain('scripts/a3s-test-cdp-browser.cmd');
});
