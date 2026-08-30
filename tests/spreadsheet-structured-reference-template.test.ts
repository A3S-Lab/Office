import { describe, expect, test } from '@rstest/core';
import { createSpreadsheetKernelWorkbook } from '../src/internal/features/work/editors/spreadsheet-calculation-projection';
import {
  createWorkArtifact,
  WORK_TEMPLATES,
} from '../src/internal/features/work/work-templates';
import { calculateSpreadsheetInJavaScript } from '../src/internal/kernel/office-kernel-spreadsheet-fallback';
import {
  OFFICE_KERNEL_PROTOCOL_VERSION,
  type OfficeKernelSpreadsheetCalculationRequest,
} from '../src/internal/kernel/office-kernel-protocol';

describe('structured-reference Playground template', () => {
  test('advertises calculated-column fill in the template metadata and sheet', () => {
    const template =
      // Keep this assertion tied to the public template registry rather than
      // duplicating the copy in the Playground UI.
      WORK_TEMPLATES.find(({ id }) => id === 'structured-references');
    expect(template?.description).toContain('插入行自动填充');

    const artifact = createWorkArtifact('structured-references');
    if (artifact.content.type !== 'spreadsheet') {
      throw new Error(
        'Expected the structured-reference template to be a workbook.',
      );
    }
    const salesSheet = artifact.content.sheets.find(
      ({ name }) => name === 'Sales',
    );
    expect(salesSheet?.data?.[1]?.[0]?.v).toContain('自动补齐 Revenue');
  });

  test('calculates table-local, selector, range, and qualified formulas', async () => {
    const artifact = createWorkArtifact('structured-references');
    if (artifact.content.type !== 'spreadsheet') {
      throw new Error(
        'Expected the structured-reference template to be a workbook.',
      );
    }
    const workbook = createSpreadsheetKernelWorkbook(artifact.content);
    if (!workbook) throw new Error('Expected a formula-bearing workbook.');

    const request: OfficeKernelSpreadsheetCalculationRequest = {
      protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
      kind: 'spreadsheetCalculation',
      requestId: 1,
      revision: 1,
      documentRevision: 1,
      sheets: workbook.sheets,
    };
    const result = await calculateSpreadsheetInJavaScript(request);
    const valueAt = (sheetId: string, row: number, column: number) =>
      result.cells.find(
        (cell) =>
          cell.sheetId === sheetId &&
          cell.row === row &&
          cell.column === column,
      )?.value;
    const salesSheet = workbook.sheets.find((sheet) => sheet.name === 'Sales');
    const summarySheet = workbook.sheets.find(
      (sheet) => sheet.name === 'Summary',
    );
    if (!salesSheet || !summarySheet)
      throw new Error('Template sheets missing.');

    expect(valueAt(salesSheet.id, 3, 3)).toEqual({
      kind: 'number',
      value: 576,
    });
    expect(valueAt(salesSheet.id, 7, 1)).toEqual({
      kind: 'number',
      value: 41,
    });
    expect(valueAt(salesSheet.id, 7, 3)).toEqual({
      kind: 'number',
      value: 2512,
    });
    expect(valueAt(salesSheet.id, 10, 1)).toEqual({
      kind: 'number',
      value: 4,
    });
    expect(valueAt(salesSheet.id, 12, 1)).toEqual({
      kind: 'number',
      value: 325,
    });
    expect(valueAt(summarySheet.id, 2, 1)).toEqual({
      kind: 'number',
      value: 2512,
    });
    expect(valueAt(summarySheet.id, 3, 1)).toEqual({
      kind: 'number',
      value: 4,
    });
    expect(result.issues).toEqual([]);
  });
});
