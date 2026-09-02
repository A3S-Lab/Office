import type { Sheet } from '@fortune-sheet/core';
import { expect, test } from '@rstest/core';
import {
  buildConditionalRule,
  conditionalRuleDraftForRule,
  newConditionalRuleDraft,
} from '../src/internal/features/work/editors/spreadsheet-conditional-format-model';
import {
  evaluateSpreadsheetLocalFormula,
  MAX_SPREADSHEET_LOCAL_FORMULA_REFERENCED_CELLS,
} from '../src/internal/features/work/work-spreadsheet-local-formula';
import { spreadsheetConditionalFormatStyles } from '../src/internal/features/work/work-spreadsheet-conditional-format';
import {
  readXlsxConditionalFormats,
  readXlsxDifferentialFormats,
  writeXlsxConditionalFormats,
  XlsxDifferentialFormatWriter,
} from '../src/internal/features/work/work-xlsx-conditional-format';
import { parseXml } from '../src/internal/features/work/work-ooxml-package';
import {
  createWorkArtifact,
  WORK_TEMPLATES,
} from '../src/internal/features/work/work-templates';

test('evaluates formula rules with relative, absolute, range, and cross-sheet references', () => {
  const sheets = [
    {
      id: 'sheet-1',
      name: 'Orders',
      data: [
        [{ v: 'Ready' }, { v: 1 }],
        [{ v: 'Blocked' }, { v: 2 }],
      ],
    },
    { id: 'sheet-2', name: 'Limits', data: [[{ v: 1 }]] },
  ];

  expect(
    evaluateSpreadsheetLocalFormula('=AND($A1="Ready",B1>0)', {
      sheets,
      sheetId: 'sheet-1',
      row: 0,
      column: 0,
      anchorRow: 0,
      anchorColumn: 0,
    }),
  ).toEqual({ supported: true, value: true });
  expect(
    evaluateSpreadsheetLocalFormula('=AND($A1="Ready",B1>0)', {
      sheets,
      sheetId: 'sheet-1',
      row: 1,
      column: 0,
      anchorRow: 0,
      anchorColumn: 0,
    }),
  ).toEqual({ supported: true, value: false });
  expect(
    evaluateSpreadsheetLocalFormula('=SUM(B1:B2)>Limits!A1', {
      sheets,
      sheetId: 'sheet-1',
      row: 0,
      column: 0,
      anchorRow: 0,
      anchorColumn: 0,
    }),
  ).toEqual({ supported: true, value: true });
});

test('fails closed for unsupported references, uncached formulas, and oversized ranges', () => {
  const sheets = [
    {
      id: 'sheet-1',
      name: 'Orders',
      data: [[{ f: '=1' }]],
    },
  ];
  expect(
    evaluateSpreadsheetLocalFormula('=A:A>0', {
      sheets,
      sheetId: 'sheet-1',
      row: 0,
      column: 0,
    }),
  ).toMatchObject({ supported: false });
  expect(
    evaluateSpreadsheetLocalFormula('=A1>0', {
      sheets,
      sheetId: 'sheet-1',
      row: 0,
      column: 0,
    }),
  ).toMatchObject({ supported: false });
  expect(
    evaluateSpreadsheetLocalFormula('=SUM(A1:A1025)>0', {
      sheets,
      sheetId: 'sheet-1',
      row: 0,
      column: 0,
    }),
  ).toMatchObject({
    supported: false,
    message: `公式一次最多读取 ${MAX_SPREADSHEET_LOCAL_FORMULA_REFERENCED_CELLS.toLocaleString()} 个单元格。`,
  });
});

test('applies formula conditional formatting with range-relative references and stop-if-true', () => {
  const sheet = {
    id: 'sheet-1',
    name: 'Orders',
    data: [
      [{ v: 'Ready' }, { v: 1 }],
      [{ v: 'Blocked' }, { v: 2 }],
      [{ v: 'Ready' }, { v: 0 }],
    ],
    luckysheet_conditionformat_save: [
      {
        type: 'default',
        cellrange: [{ row: [0, 2], column: [0, 0] }],
        format: { textColor: '#006100', cellColor: '#c6efce' },
        conditionName: 'formula',
        conditionValue: ['=AND($A1="Ready",B1>0)'],
        stopIfTrue: true,
      },
      {
        type: 'default',
        cellrange: [{ row: [0, 2], column: [0, 0] }],
        format: { textColor: '#9c0006', cellColor: '#ffc7ce' },
        conditionName: 'formula',
        conditionValue: ['=A1="Blocked"'],
      },
    ],
  } as unknown as Sheet;

  const styles = spreadsheetConditionalFormatStyles(sheet);
  expect(styles.get('0_0')).toEqual({
    textColor: '#006100',
    cellColor: '#c6efce',
  });
  expect(styles.get('1_0')).toEqual({
    textColor: '#9c0006',
    cellColor: '#ffc7ce',
  });
  expect(styles.has('2_0')).toBe(false);
});

test('styles blank cells for bounded formula-rule ranges without materializing them', () => {
  const sheet = {
    id: 'sheet-1',
    name: 'Orders',
    data: [],
    luckysheet_conditionformat_save: [
      {
        type: 'default',
        cellrange: [{ row: [0, 1], column: [1, 1] }],
        format: { textColor: null, cellColor: '#fff2cc' },
        conditionName: 'formula',
        conditionValue: ['=ISBLANK(B1)'],
      },
    ],
  } as unknown as Sheet;

  expect([...spreadsheetConditionalFormatStyles(sheet).keys()]).toEqual([
    '0_1',
    '1_1',
  ]);
});

test('builds and reopens an editable formula conditional-format draft', () => {
  const draft = newConditionalRuleDraft('sheet-1');
  draft.type = 'formula';
  draft.reference = 'A2:A20';
  draft.formula = 'AND($B2="Ready",C2>0)';
  draft.formulaUseTextColor = false;
  draft.formulaUseCellColor = true;
  draft.formulaCellColor = '#d9eaff';
  draft.stopIfTrue = true;

  const built = buildConditionalRule(draft);
  if (!('rule' in built)) throw new Error(built.error);
  expect(built.rule).toMatchObject({
    type: 'default',
    conditionName: 'formula',
    conditionValue: ['=AND($B2="Ready",C2>0)'],
    stopIfTrue: true,
    format: { textColor: null, cellColor: '#d9eaff' },
  });
  expect(conditionalRuleDraftForRule('sheet-1', built.rule)).toMatchObject({
    type: 'formula',
    reference: 'A2:A20',
    formula: '=AND($B2="Ready",C2>0)',
    formulaUseTextColor: false,
    formulaUseCellColor: true,
  });
});

test('round-trips formula rules through native XLSX expression records', () => {
  const worksheet = parseXml(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
    'formula conditional-format worksheet',
  );
  const styles = parseXml(
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dxfs count="0"/></styleSheet>',
    'formula conditional-format styles',
  );
  const differentialFormats = new XlsxDifferentialFormatWriter(styles);
  writeXlsxConditionalFormats(
    worksheet,
    [
      {
        type: 'default',
        cellrange: [{ row: [1, 19], column: [0, 0] }],
        format: { textColor: '#006100', cellColor: '#c6efce' },
        conditionName: 'formula',
        conditionValue: ['=AND($B2="Ready",C2>0)'],
        stopIfTrue: true,
      },
    ],
    differentialFormats,
  );

  expect(worksheet.documentElement.outerHTML).toContain('type="expression"');
  expect(worksheet.documentElement.outerHTML).toContain(
    'AND($B2="Ready",C2&gt;0)',
  );
  const reopened = readXlsxConditionalFormats(
    worksheet,
    readXlsxDifferentialFormats(styles),
  );
  expect(reopened).toEqual([
    expect.objectContaining({
      type: 'default',
      conditionName: 'formula',
      conditionValue: ['AND($B2="Ready",C2>0)'],
      stopIfTrue: true,
    }),
  ]);
});

test('publishes a formula conditional-format template with local cross-sheet rules', () => {
  expect(WORK_TEMPLATES).toContainEqual(
    expect.objectContaining({
      id: 'conditional-format',
      kind: 'spreadsheet',
      name: '公式条件格式',
    }),
  );
  const artifact = createWorkArtifact('conditional-format');
  expect(artifact.title).toBe('公式条件格式示例');
  if (artifact.content.type !== 'spreadsheet') {
    throw new Error(
      'Expected the conditional-format template to be a workbook.',
    );
  }
  const status = artifact.content.sheets.find(({ name }) => name === 'Status');
  const limits = artifact.content.sheets.find(({ name }) => name === 'Limits');
  expect(status?.luckysheet_conditionformat_save).toEqual([
    expect.objectContaining({
      conditionName: 'formula',
      conditionValue: ['=$D4="阻塞"'],
      stopIfTrue: true,
    }),
    expect.objectContaining({
      conditionName: 'formula',
      conditionValue: ['=AND($C4>=Limits!$A$2,$D4<>"阻塞")'],
    }),
  ]);
  expect(limits?.data?.[1]?.[0]).toMatchObject({ v: 0.8 });
  expect(
    spreadsheetConditionalFormatStyles(
      status as Sheet,
      artifact.content.sheets,
    ).get('5_0'),
  ).toEqual({
    textColor: '#9c0006',
    cellColor: '#ffc7ce',
  });
  expect(
    spreadsheetConditionalFormatStyles(
      status as Sheet,
      artifact.content.sheets,
    ).get('3_0'),
  ).toEqual({
    textColor: '#006100',
    cellColor: '#c6efce',
  });
});
