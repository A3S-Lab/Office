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
    'visual-1',
    'visual-2',
    'visual-3',
    'visual-4',
    'visual-5',
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
    ]),
  );
  expect(
    plans.find((plan) => plan.surface.id === 'spreadsheet')?.surface.acl,
  ).toEqual(
    expect.arrayContaining([
      'tests/e2e/spreadsheet-ribbon-alignment.acl',
      'tests/e2e/spreadsheet-font-dialog-shortcuts.acl',
    ]),
  );
  expect(
    plans.find((plan) => plan.surface.id === 'presentation')?.surface.acl,
  ).toEqual(
    expect.arrayContaining(['tests/e2e/presentation-cut-paste-focus.acl']),
  );
  expect(plans.find((plan) => plan.surface.id === 'pdf')?.surface.acl).toEqual(
    expect.arrayContaining([
      'tests/e2e/pdf-thumbnail-keyboard.acl',
      'tests/e2e/pdf-page-drawer-phone.acl',
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
