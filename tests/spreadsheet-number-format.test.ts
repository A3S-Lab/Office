import { describe, expect, test } from '@rstest/core';
import {
  adjustSpreadsheetNumberFormat,
  spreadsheetNumberFormatValue,
  spreadsheetNumberFormatCode,
  spreadsheetNumberFormatPreset,
} from '../src/internal/features/work/editors/spreadsheet-number-format';
import {
  createWorkArtifactBlob,
  importWorkFile,
} from '../src/internal/features/work/work-file-io';
import { createWorkArtifact } from '../src/internal/features/work/work-templates';

describe('spreadsheet number formatting', () => {
  test('maps common format codes to the primary ribbon presets', () => {
    expect(spreadsheetNumberFormatPreset()).toBe('general');
    expect(spreadsheetNumberFormatPreset('General')).toBe('general');
    expect(spreadsheetNumberFormatPreset('#,##0.00')).toBe('number');
    expect(spreadsheetNumberFormatPreset('¥ #,##0.00')).toBe('number');
    expect(spreadsheetNumberFormatPreset('0.0%')).toBe('percent');
    expect(spreadsheetNumberFormatPreset('yyyy-mm-dd')).toBe('custom');
    expect(spreadsheetNumberFormatPreset('m/d/yyyy 00:00')).toBe('custom');
    expect(spreadsheetNumberFormatPreset('@')).toBe('custom');
  });

  test('owns stable format codes for the exposed presets', () => {
    expect(spreadsheetNumberFormatCode('general')).toBe('General');
    expect(spreadsheetNumberFormatCode('number')).toBe('#,##0.00');
    expect(spreadsheetNumberFormatCode('percent')).toBe('0.00%');
  });

  test('builds the complete Fortune cell-format contract', () => {
    expect(spreadsheetNumberFormatValue('0.00%', { v: 0.5 })).toEqual({
      fa: '0.00%',
      t: 'n',
    });
    expect(spreadsheetNumberFormatValue('General', { v: 'Ready' })).toEqual({
      fa: 'General',
      t: 'g',
    });
  });

  test('adjusts decimals without dropping grouping, currency, or percent', () => {
    expect(adjustSpreadsheetNumberFormat('General', 1)).toBe('0.0');
    expect(adjustSpreadsheetNumberFormat('General', -1)).toBe('0');
    expect(adjustSpreadsheetNumberFormat('#,##0.00', -1)).toBe('#,##0.0');
    expect(adjustSpreadsheetNumberFormat('#,##0.0', -1)).toBe('#,##0');
    expect(adjustSpreadsheetNumberFormat('0%', 1)).toBe('0.0%');
    expect(adjustSpreadsheetNumberFormat('¥ #,##0.00', 1)).toBe('¥ #,##0.000');
    expect(adjustSpreadsheetNumberFormat('yyyy-mm-dd', 1)).toBe('yyyy-mm-dd');
  });

  test('ships the quarterly plan with readable progress values', () => {
    const content = createWorkArtifact('quarterly-plan').content;
    if (content.type !== 'spreadsheet')
      throw new Error('Expected the quarterly plan spreadsheet.');

    const firstProgressRow = content.sheets[0]?.data?.[3];
    expect(firstProgressRow?.[2]).toMatchObject({
      ct: { fa: '0%', t: 'n' },
      m: '100%',
      v: 1,
    });
    expect(firstProgressRow?.[5]).toMatchObject({
      ct: { fa: '0%', t: 'n' },
      f: '=SUM(C4:E4)/3',
    });
  });

  test('preserves the visible percentage format through XLSX export and import', async () => {
    const artifact = createWorkArtifact('quarterly-plan');
    const blob = await createWorkArtifactBlob(artifact);
    const imported = await importWorkFile(
      new File([blob], 'quarterly-plan.xlsx', { type: blob.type }),
    );
    if (imported.content.type !== 'spreadsheet')
      throw new Error('Expected an imported spreadsheet.');

    expect(imported.content.sheets[0]?.data?.[3]?.[2]).toMatchObject({
      ct: { fa: '0%' },
      v: 1,
    });
  });
});
