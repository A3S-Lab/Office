import { describe, expect, test } from '@rstest/core';
import {
  applySpreadsheetCellStyle,
  canApplySpreadsheetCellStyle,
  MAX_SPREADSHEET_CELL_STYLE_CELLS,
  spreadsheetCellStyleDefinitions,
  spreadsheetCellStylePreset,
  spreadsheetCellStylePresetIds,
} from '../src/internal/features/work/editors/spreadsheet-cell-style';
import {
  createWorkArtifactBlob,
  importWorkFile,
} from '../src/internal/features/work/work-file-io';
import { createWorkArtifact } from '../src/internal/features/work/work-templates';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';
import { parseXml } from '../src/internal/features/work/work-ooxml-package';
import { readXlsxDirectCellStyles } from '../src/internal/features/work/work-xlsx-cell-styles';

describe('spreadsheet cell styles', () => {
  test('owns the common WPS built-in style catalog', () => {
    expect(spreadsheetCellStylePresetIds).toEqual([
      'normal',
      'good',
      'bad',
      'neutral',
      'calculation',
      'checkCell',
      'explanatoryText',
      'input',
      'linkedCell',
      'note',
      'output',
      'warningText',
      'heading1',
      'heading2',
      'heading3',
      'heading4',
      'total',
    ]);
    expect(
      spreadsheetCellStyleDefinitions.map(({ id, label, group }) => ({
        group,
        id,
        label,
      })),
    ).toEqual([
      { group: '常用', id: 'normal', label: '常规' },
      { group: '常用', id: 'good', label: '好' },
      { group: '常用', id: 'bad', label: '差' },
      { group: '常用', id: 'neutral', label: '适中' },
      { group: '数据和模型', id: 'calculation', label: '计算' },
      { group: '数据和模型', id: 'checkCell', label: '检查单元格' },
      { group: '数据和模型', id: 'explanatoryText', label: '解释性文本' },
      { group: '数据和模型', id: 'input', label: '输入' },
      { group: '数据和模型', id: 'linkedCell', label: '链接单元格' },
      { group: '数据和模型', id: 'note', label: '注释' },
      { group: '数据和模型', id: 'output', label: '输出' },
      { group: '数据和模型', id: 'warningText', label: '警告文本' },
      { group: '标题和汇总', id: 'heading1', label: '标题 1' },
      { group: '标题和汇总', id: 'heading2', label: '标题 2' },
      { group: '标题和汇总', id: 'heading3', label: '标题 3' },
      { group: '标题和汇总', id: 'heading4', label: '标题 4' },
      { group: '标题和汇总', id: 'total', label: '总计' },
    ]);
  });

  test('recognizes native cell appearances without a second style marker', () => {
    expect(spreadsheetCellStylePreset(null)).toBe('normal');
    expect(
      spreadsheetCellStylePreset({
        bg: '#C6EFCE',
        bl: 0,
        cl: 0,
        fc: '#006100',
        ff: 'Aptos',
        fs: 10,
        it: 0,
        un: 0,
      }),
    ).toBe('good');
    expect(
      spreadsheetCellStylePreset({
        bg: '#ffffff',
        bl: 0,
        cl: 0,
        fc: '#0563c1',
        ff: 'Aptos',
        fs: 10,
        it: 0,
        un: 1,
      }),
    ).toBe('linkedCell');
    expect(spreadsheetCellStylePreset({ bl: 1 })).toBe('custom');
    expect(
      spreadsheetCellStylePreset(
        {
          bg: '#ffffff',
          bl: 1,
          cl: 0,
          fc: '#172033',
          ff: 'Aptos',
          fs: 10,
          it: 0,
          un: 0,
        },
        { top: { color: '#172033', style: '8' } },
      ),
    ).toBe('total');
  });

  test('applies one immutable native style while preserving cell content', () => {
    const source = workbook();
    const sourceCell = source.sheets[0]?.data?.[0]?.[0];
    const next = applySpreadsheetCellStyle(
      source,
      'sheet-1',
      { row: [0, 0], column: [0, 1] },
      'good',
    );

    expect(next).not.toBeNull();
    expect(next).not.toBe(source);
    expect(next?.sheets[0]).not.toBe(source.sheets[0]);
    expect(next?.sheets[0]?.data).not.toBe(source.sheets[0]?.data);
    expect(source.sheets[0]?.data?.[0]?.[0]).toBe(sourceCell);
    expect(sourceCell).toMatchObject({ bg: '#112233', fc: '#ffffff' });
    expect(next?.sheets[0]?.data?.[0]?.[0]).toEqual({
      bg: '#c6efce',
      bl: 0,
      cl: 0,
      ct: { fa: '0.00', t: 'n' },
      f: '=5+5',
      fc: '#006100',
      ff: 'Aptos',
      fs: 10,
      hl: { c: 0, id: 'link-1', r: 0 },
      it: 0,
      m: '10.00',
      mc: { c: 0, cs: 1, r: 0, rs: 1 },
      ps: {
        height: null,
        isShow: false,
        left: null,
        top: null,
        value: 'Keep this note',
        width: null,
      },
      un: 0,
      v: 10,
    });
    expect(next?.sheets[0]?.data?.[0]?.[1]).toMatchObject({
      bg: '#c6efce',
      bl: 0,
      cl: 0,
      fc: '#006100',
      ff: 'Aptos',
      fs: 10,
      it: 0,
      un: 0,
    });
    expect(next?.sheets[0]?.config?.borderInfo).toEqual([
      { vendor: 'preserve' },
      {
        borderType: 'border-none',
        color: '#000000',
        range: [{ column: [0, 1], row: [0, 0] }],
        rangeType: 'range',
        style: '1',
      },
    ]);
  });

  test('adds WPS input borders after clearing the previous border state', () => {
    const next = applySpreadsheetCellStyle(
      workbook(),
      'sheet-1',
      { row: [0, 0], column: [0, 0] },
      'input',
    );

    expect(next?.sheets[0]?.config?.borderInfo).toEqual([
      { vendor: 'preserve' },
      {
        borderType: 'border-none',
        color: '#000000',
        range: [{ column: [0, 0], row: [0, 0] }],
        rangeType: 'range',
        style: '1',
      },
      {
        rangeType: 'cell',
        value: {
          b: { color: '#7f8fa6', style: '1' },
          col_index: 0,
          l: { color: '#7f8fa6', style: '1' },
          r: { color: '#7f8fa6', style: '1' },
          row_index: 0,
          t: { color: '#7f8fa6', style: '1' },
        },
      },
    ]);
  });

  test('applies heading borders to every selected cell like a WPS cell style', () => {
    const next = applySpreadsheetCellStyle(
      workbook(),
      'sheet-1',
      { row: [0, 1], column: [0, 1] },
      'heading1',
    );

    const cellBorders = (next?.sheets[0]?.config?.borderInfo ?? []).filter(
      (entry) => (entry as { rangeType?: unknown }).rangeType === 'cell',
    ) as Array<{
      value: {
        b?: { color: string; style: string };
        col_index: number;
        row_index: number;
      };
    }>;
    expect(cellBorders).toHaveLength(4);
    expect(
      cellBorders.map(({ value }) => ({
        border: value.b,
        column: value.col_index,
        row: value.row_index,
      })),
    ).toEqual([
      { border: { color: '#5b9bd5', style: '13' }, column: 0, row: 0 },
      { border: { color: '#5b9bd5', style: '13' }, column: 1, row: 0 },
      { border: { color: '#5b9bd5', style: '13' }, column: 0, row: 1 },
      { border: { color: '#5b9bd5', style: '13' }, column: 1, row: 1 },
    ]);
  });

  test('bounds materialized style ranges and rejects invalid targets', () => {
    const source = workbook();
    expect(
      canApplySpreadsheetCellStyle(
        source,
        'sheet-1',
        { row: [0, 0], column: [0, MAX_SPREADSHEET_CELL_STYLE_CELLS - 1] },
        'normal',
      ),
    ).toBe(true);
    expect(
      canApplySpreadsheetCellStyle(
        source,
        'sheet-1',
        { row: [0, 0], column: [0, MAX_SPREADSHEET_CELL_STYLE_CELLS] },
        'normal',
      ),
    ).toBe(false);
    expect(
      applySpreadsheetCellStyle(
        source,
        'missing',
        { row: [0, 0], column: [0, 0] },
        'normal',
      ),
    ).toBeNull();
    expect(
      applySpreadsheetCellStyle(
        source,
        'sheet-1',
        { row: [Number.NaN, 0], column: [0, 0] },
        'normal',
      ),
    ).toBeNull();
  });

  test('resolves WPS theme, indexed, tinted, and border colors on import', () => {
    const worksheet = parseXml(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"/></row></sheetData></worksheet>',
      'xl/worksheets/sheet1.xml',
    );
    const styles = parseXml(
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font/><font><color theme="4" tint="0.5"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor indexed="0"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top style="medium"><color theme="5"/></top><bottom/><diagonal/></border></borders><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/></cellXfs><colors><indexedColors><rgbColor rgb="FF102030"/></indexedColors></colors></styleSheet>',
      'xl/styles.xml',
    );
    const theme = parseXml(
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="WPS"><a:accent1><a:srgbClr val="336699"/></a:accent1><a:accent2><a:srgbClr val="CC3300"/></a:accent2></a:clrScheme></a:themeElements></a:theme>',
      'xl/theme/theme1.xml',
    );

    expect(readXlsxDirectCellStyles(worksheet, styles, theme)).toEqual([
      {
        border: {
          top: { color: '#cc3300', style: 'medium' },
        },
        column: 0,
        row: 0,
        style: {
          bg: '#102030',
          fc: '#8cb3d9',
        },
      },
    ]);
  });

  test('preserves every native XLSX underline variant for Fortune rendering', () => {
    const worksheet = parseXml(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"/><c r="B1" s="2"/><c r="C1" s="3"/><c r="D1" s="4"/></row></sheetData></worksheet>',
      'xl/worksheets/sheet1.xml',
    );
    const styles = parseXml(
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="5"><font/><font><u/></font><font><u val="double"/></font><font><u val="singleAccounting"/></font><font><u val="doubleAccounting"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="1"/><xf numFmtId="0" fontId="3" fillId="0" borderId="0" applyFont="1"/><xf numFmtId="0" fontId="4" fillId="0" borderId="0" applyFont="1"/></cellXfs></styleSheet>',
      'xl/styles.xml',
    );

    expect(readXlsxDirectCellStyles(worksheet, styles)).toEqual([
      { border: undefined, column: 0, row: 0, style: { un: 1 } },
      { border: undefined, column: 1, row: 0, style: { un: 2 } },
      { border: undefined, column: 2, row: 0, style: { un: 3 } },
      { border: undefined, column: 3, row: 0, style: { un: 4 } },
    ]);
  });

  test('round-trips the exact direct style through XLSX', async () => {
    const artifact = createWorkArtifact('blank-spreadsheet');
    if (artifact.content.type !== 'spreadsheet')
      throw new Error('Expected a blank spreadsheet.');
    const sheet = artifact.content.sheets[0];
    if (!sheet) throw new Error('Expected a worksheet.');
    sheet.data = [[{ ct: { fa: '0.00', t: 'n' }, v: 12.5 }]];
    const styled = applySpreadsheetCellStyle(
      artifact.content,
      sheet.id,
      { row: [0, 0], column: [0, 1] },
      'good',
    );
    if (!styled) throw new Error('Expected a styled workbook.');
    artifact.content = styled;

    const blob = await createWorkArtifactBlob(artifact);
    const imported = await importWorkFile(
      new File([blob], 'cell-styles.xlsx', { type: blob.type }),
    );
    if (imported.content.type !== 'spreadsheet')
      throw new Error('Expected an imported spreadsheet.');

    const importedCell = imported.content.sheets[0]?.data?.[0]?.[0];
    expect(importedCell).toMatchObject({
      bg: '#c6efce',
      ct: { fa: '0.00', t: 'n' },
      fc: '#006100',
      ff: 'Aptos',
      fs: 10,
    });
    expect(importedCell?.bl).toBeUndefined();
    expect(importedCell?.cl).toBeUndefined();
    expect(importedCell?.ht).toBeUndefined();
    expect(importedCell?.it).toBeUndefined();
    expect(importedCell?.tb).toBeUndefined();
    expect(importedCell?.un).toBeUndefined();
    expect(importedCell?.vt).toBeUndefined();
    expect(imported.content.sheets[0]?.data?.[0]?.[1]).toMatchObject({
      bg: '#c6efce',
      fc: '#006100',
      ff: 'Aptos',
      fs: 10,
    });
  });

  test('round-trips all advanced underline styles through native XLSX fonts', async () => {
    const artifact = createWorkArtifact('blank-spreadsheet');
    if (artifact.content.type !== 'spreadsheet')
      throw new Error('Expected a blank spreadsheet.');
    const sheet = artifact.content.sheets[0];
    if (!sheet) throw new Error('Expected a worksheet.');
    sheet.data = [
      [
        { un: 1, v: 'Single' },
        { un: 2, v: 'Double' },
        { un: 3, v: 'Single accounting' },
        { un: 4, v: 'Double accounting' },
      ],
    ];

    const blob = await createWorkArtifactBlob(artifact);
    const imported = await importWorkFile(
      new File([blob], 'underline-styles.xlsx', { type: blob.type }),
    );
    if (imported.content.type !== 'spreadsheet')
      throw new Error('Expected an imported spreadsheet.');

    expect(
      imported.content.sheets[0]?.data?.[0]
        ?.slice(0, 4)
        .map((cell) => cell?.un),
    ).toEqual([1, 2, 3, 4]);
  });

  test('round-trips built-in style borders through native XLSX records', async () => {
    const artifact = createWorkArtifact('blank-spreadsheet');
    if (artifact.content.type !== 'spreadsheet')
      throw new Error('Expected a blank spreadsheet.');
    const sheet = artifact.content.sheets[0];
    if (!sheet) throw new Error('Expected a worksheet.');
    sheet.data = [[{ v: 'Input' }, { v: 'Total' }]];
    const withInput = applySpreadsheetCellStyle(
      artifact.content,
      sheet.id,
      { row: [0, 0], column: [0, 0] },
      'input',
    );
    if (!withInput) throw new Error('Expected an input style.');
    const styled = applySpreadsheetCellStyle(
      withInput,
      sheet.id,
      { row: [0, 0], column: [1, 1] },
      'total',
    );
    if (!styled) throw new Error('Expected a total style.');
    artifact.content = styled;

    const blob = await createWorkArtifactBlob(artifact);
    const imported = await importWorkFile(
      new File([blob], 'cell-style-borders.xlsx', { type: blob.type }),
    );
    if (imported.content.type !== 'spreadsheet')
      throw new Error('Expected an imported spreadsheet.');

    expect(imported.content.sheets[0]?.config?.borderInfo).toEqual([
      {
        rangeType: 'cell',
        value: {
          b: { color: '#7f8fa6', style: '1' },
          col_index: 0,
          l: { color: '#7f8fa6', style: '1' },
          r: { color: '#7f8fa6', style: '1' },
          row_index: 0,
          t: { color: '#7f8fa6', style: '1' },
        },
      },
      {
        rangeType: 'cell',
        value: {
          col_index: 1,
          row_index: 0,
          t: { color: '#172033', style: '8' },
        },
      },
    ]);
  });
});

function workbook(): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        column: 2,
        config: { borderInfo: [{ vendor: 'preserve' }] },
        data: [
          [
            {
              bg: '#112233',
              bl: 1,
              cl: 1,
              ct: { fa: '0.00', t: 'n' },
              f: '=5+5',
              fc: '#ffffff',
              ff: 'Arial',
              fs: 12,
              hl: { c: 0, id: 'link-1', r: 0 },
              it: 1,
              m: '10.00',
              mc: { c: 0, cs: 1, r: 0, rs: 1 },
              ps: {
                height: null,
                isShow: false,
                left: null,
                top: null,
                value: 'Keep this note',
                width: null,
              },
              un: 1,
              v: 10,
            },
            null,
          ],
        ],
        id: 'sheet-1',
        name: 'Sheet 1',
        row: 1,
      },
    ],
  };
}
