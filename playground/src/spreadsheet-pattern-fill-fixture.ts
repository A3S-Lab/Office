import type { OfficeArtifact, SpreadsheetContent } from '@a3s-lab/office/core';
import {
  withXlsxPatternFill,
  xlsxPatternFillTypes,
} from '../../src/internal/features/work/work-xlsx-pattern-fill';

export const SPREADSHEET_PATTERN_FILL_FIXTURE = 'spreadsheet-pattern-fill';
export const SPREADSHEET_PATTERN_FILL_ARTIFACT_ID =
  'fixture-spreadsheet-pattern-fill';

const foregroundColors = ['#1d4ed8', '#b42318', '#067647', '#7a5af8'];
const backgroundColors = ['#eff6ff', '#fff4ed', '#ecfdf3', '#f4f3ff'];

export function createSpreadsheetPatternFillArtifact(): OfficeArtifact {
  const now = Date.now();
  const data: NonNullable<SpreadsheetContent['sheets'][number]['data']> = [];
  data[0] = [
    {
      v: 'Native XLSX pattern fills',
      m: 'Native XLSX pattern fills',
      bl: 1,
      fc: '#172033',
      fs: 15,
      bg: '#eaf0ff',
    },
  ];
  for (const [index, patternType] of xlsxPatternFillTypes.entries()) {
    const row = 2 + Math.floor(index / 4) * 2;
    const column = index % 4;
    const backgroundColor = backgroundColors[column] ?? '#ffffff';
    const foregroundColor = foregroundColors[column] ?? '#172033';
    data[row] ??= [];
    data[row][column] = withXlsxPatternFill(
      {
        v: patternType,
        m: patternType,
        bg: backgroundColor,
        fc: '#172033',
        fs: 10,
        ht: 0,
        vt: 0,
      },
      { backgroundColor, foregroundColor, patternType },
    );
  }

  return {
    id: SPREADSHEET_PATTERN_FILL_ARTIFACT_ID,
    kind: 'spreadsheet',
    title: 'Native pattern-fill gallery',
    favorite: false,
    createdAt: now,
    updatedAt: now,
    lastOpenedAt: now,
    revision: 1,
    content: {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'pattern-fill-sheet',
          name: 'Pattern Fills',
          status: 1,
          row: 14,
          column: 8,
          data,
          config: {
            columnlen: { 0: 150, 1: 150, 2: 150, 3: 150 },
            rowlen: { 0: 34, 2: 44, 4: 44, 6: 44, 8: 44, 10: 44 },
            merge: { '0_0': { r: 0, c: 0, rs: 1, cs: 4 } },
          },
          luckysheet_select_save: [
            {
              row: [2, 2],
              column: [0, 0],
              row_focus: 2,
              column_focus: 0,
            },
          ],
        },
      ],
    },
  };
}
