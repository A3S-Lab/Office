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
    expect(spreadsheetNumberFormatPreset('[$¥-804]#,##0.00')).toBe('currency');
    expect(
      spreadsheetNumberFormatPreset(
        '_([$¥-804]* #,##0.00_);_([$¥-804]* (#,##0.00)',
      ),
    ).toBe('accounting');
    expect(spreadsheetNumberFormatPreset('¥(0.00)')).toBe('accounting');
    expect(spreadsheetNumberFormatPreset('0.0%')).toBe('percent');
    expect(spreadsheetNumberFormatPreset('yyyy-mm-dd')).toBe('date');
    expect(spreadsheetNumberFormatPreset('h:mm AM/PM')).toBe('time');
    expect(spreadsheetNumberFormatPreset('0.00E+00')).toBe('scientific');
    expect(spreadsheetNumberFormatPreset('# ?/?')).toBe('fraction');
    expect(spreadsheetNumberFormatPreset('@')).toBe('text');
    expect(spreadsheetNumberFormatPreset('m/d/yyyy 00:00')).toBe('custom');
  });

  test('owns stable format codes for the exposed presets', () => {
    expect(spreadsheetNumberFormatCode('general')).toBe('General');
    expect(spreadsheetNumberFormatCode('number')).toBe('#,##0.00');
    expect(spreadsheetNumberFormatCode('currency')).toBe('[$¥-804]#,##0.00');
    expect(spreadsheetNumberFormatCode('accounting')).toBe(
      '_([$¥-804]* #,##0.00_);_([$¥-804]* (#,##0.00)',
    );
    expect(spreadsheetNumberFormatCode('percent')).toBe('0.00%');
    expect(spreadsheetNumberFormatCode('date')).toBe('yyyy-MM-dd');
    expect(spreadsheetNumberFormatCode('time')).toBe('hh:mm');
    expect(spreadsheetNumberFormatCode('scientific')).toBe('0.00E+00');
    expect(spreadsheetNumberFormatCode('fraction')).toBe('# ?/?');
    expect(spreadsheetNumberFormatCode('text')).toBe('@');
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
    expect(spreadsheetNumberFormatValue('yyyy-MM-dd', { v: 45_292 })).toEqual({
      fa: 'yyyy-MM-dd',
      t: 'd',
    });
    expect(spreadsheetNumberFormatValue('hh:mm', { v: 0.5 })).toEqual({
      fa: 'hh:mm',
      t: 'd',
    });
    expect(spreadsheetNumberFormatValue('@', { v: 123 })).toEqual({
      fa: '@',
      t: 's',
    });
    expect(spreadsheetNumberFormatValue('# ?/?', { v: 1.5 })).toEqual({
      fa: '# ?/?',
      t: 'n',
    });
  });

  test('adjusts decimals without dropping grouping, currency, or percent', () => {
    expect(adjustSpreadsheetNumberFormat('General', 1)).toBe('0.0');
    expect(adjustSpreadsheetNumberFormat('General', -1)).toBe('0');
    expect(adjustSpreadsheetNumberFormat('#,##0.00', -1)).toBe('#,##0.0');
    expect(adjustSpreadsheetNumberFormat('#,##0.0', -1)).toBe('#,##0');
    expect(adjustSpreadsheetNumberFormat('0%', 1)).toBe('0.0%');
    expect(adjustSpreadsheetNumberFormat('[$¥-804]#,##0.00', 1)).toBe(
      '[$¥-804]#,##0.000',
    );
    expect(adjustSpreadsheetNumberFormat('0.00E+00', -1)).toBe('0.0E+00');
    expect(adjustSpreadsheetNumberFormat('yyyy-mm-dd', 1)).toBe('yyyy-mm-dd');
    expect(adjustSpreadsheetNumberFormat('hh:mm', 1)).toBe('hh:mm');
    expect(adjustSpreadsheetNumberFormat('# ?/?', 1)).toBe('# ?/?');
    expect(adjustSpreadsheetNumberFormat('@', 1)).toBe('@');
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

  test('preserves the common WPS number-format matrix through XLSX', async () => {
    const artifact = createWorkArtifact('blank-spreadsheet');
    if (artifact.content.type !== 'spreadsheet')
      throw new Error('Expected a blank spreadsheet.');
    const cases = [
      { preset: 'number', value: 1234.5 },
      { preset: 'currency', value: 1234.5 },
      { preset: 'accounting', value: -1234.5 },
      { preset: 'percent', value: 0.125 },
      { preset: 'date', value: 45_292 },
      { preset: 'time', value: 0.5 },
      { preset: 'scientific', value: 1234.5 },
      { preset: 'fraction', value: 1.5 },
      { preset: 'text', value: '00123' },
    ] as const;
    const sheet = artifact.content.sheets[0];
    if (!sheet) throw new Error('Expected a worksheet.');
    sheet.row = 1;
    sheet.column = cases.length;
    sheet.data = [
      cases.map(({ preset, value }) => ({
        ct: spreadsheetNumberFormatValue(spreadsheetNumberFormatCode(preset), {
          v: value,
        }),
        v: value,
      })),
    ];

    const blob = await createWorkArtifactBlob(artifact);
    const imported = await importWorkFile(
      new File([blob], 'number-formats.xlsx', { type: blob.type }),
    );
    if (imported.content.type !== 'spreadsheet')
      throw new Error('Expected an imported spreadsheet.');

    expect(
      imported.content.sheets[0]?.data?.[0]
        ?.slice(0, cases.length)
        .map((cell) => cell?.ct?.fa),
    ).toEqual(cases.map(({ preset }) => spreadsheetNumberFormatCode(preset)));
  });
});
