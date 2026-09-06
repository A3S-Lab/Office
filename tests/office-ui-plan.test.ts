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
  });
  expect(commandIds).toEqual([
    'fixtures',
    'check',
    'gate',
    'visual-1',
    'visual-2',
    'visual-3',
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
