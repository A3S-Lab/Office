import { describe, expect, test } from '@rstest/core';
import {
  OFFICE_KERNEL_PROTOCOL_VERSION,
  type OfficeKernelSpreadsheetCalculatedCell,
  type OfficeKernelSpreadsheetCalculationRequest,
  type OfficeKernelSpreadsheetCoordinate,
  type OfficeKernelSpreadsheetInputSheet,
} from '../../../src/internal/kernel/office-kernel-protocol';
import { calculateSpreadsheetInJavaScript } from '../../../src/internal/kernel/office-kernel-spreadsheet-fallback';
import parityFixtures from '../../fixtures/spreadsheet-kernel-parity.json';
import dailyFixtures from './fixtures/r2-formula-daily.json';

/**
 * R2 formula-compatibility corpus gate.
 *
 * Pins the existing kernel scalar parity fixtures plus a bounded Traditional
 * Office / WPS daily-formula set. Unsupported functions must fail closed with
 * `office.kernel.spreadsheet.formula_unsupported` rather than invent values.
 */
describe('R2 formula-compatibility corpus', () => {
  const fixtures = [...parityFixtures, ...dailyFixtures];

  test('keeps the versioned corpus non-empty and uniquely named', () => {
    expect(fixtures.length).toBeGreaterThanOrEqual(18);
    const names = fixtures.map((fixture) => fixture.name);
    expect(new Set(names).size).toBe(names.length);
  });

  for (const fixture of fixtures) {
    test(`calculates ${fixture.name}`, async () => {
      const input: OfficeKernelSpreadsheetCalculationRequest = {
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: 'spreadsheetCalculation',
        requestId: 21,
        revision: 1,
        documentRevision: 1,
        sheets: fixture.sheets as OfficeKernelSpreadsheetInputSheet[],
        targets: fixture.targets as OfficeKernelSpreadsheetCoordinate[],
        ...('dateSystem' in fixture && fixture.dateSystem
          ? {
              dateSystem: fixture.dateSystem as '1900' | '1904',
            }
          : {}),
      };

      const result = await calculateSpreadsheetInJavaScript(input);

      expect(result.cells).toEqual(
        fixture.expectedCells as OfficeKernelSpreadsheetCalculatedCell[],
      );
      expect(result.calculationOrder).toEqual(
        fixture.expectedOrder as OfficeKernelSpreadsheetCoordinate[],
      );
      expect(result.issues.map(({ cell, code }) => ({ cell, code }))).toEqual(
        fixture.expectedIssues,
      );
    });
  }
});
