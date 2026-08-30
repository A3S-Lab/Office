import type { OfficeArtifact, SpreadsheetContent } from '@a3s-lab/office/core';

export const SPREADSHEET_TEXT_SORT_FIXTURE = 'spreadsheet-text-sort';
export const SPREADSHEET_TEXT_SORT_ARTIFACT_ID =
  'fixture-spreadsheet-text-sort';

export function createSpreadsheetTextSortArtifact(): OfficeArtifact {
  const now = Date.now();
  const data: NonNullable<SpreadsheetContent['sheets'][number]['data']> = [
    [
      { v: '姓名', m: '姓名', bl: 1, bg: '#d9ead3' },
      { v: '编号', m: '编号', bl: 1, bg: '#d9ead3' },
      { v: '说明', m: '说明', bl: 1, bg: '#d9ead3' },
    ],
    [
      { v: '赵', m: '赵' },
      { v: 'K2', m: 'K2' },
      { v: 'K2-Z', m: 'K2-Z', f: '=B2&"-Z"' },
    ],
    [
      { v: '阿', m: '阿' },
      { v: 'K11', m: 'K11' },
      { v: 'K11-A', m: 'K11-A', f: '=B3&"-A"' },
    ],
    [
      { v: '丁', m: '丁' },
      { v: 'K100', m: 'K100' },
      { v: 'K100-D', m: 'K100-D', f: '=B4&"-D"' },
    ],
    [
      { v: '安', m: '安' },
      { v: 'K1', m: 'K1' },
      { v: 'K1-N', m: 'K1-N', f: '=B5&"-N"' },
    ],
    [
      { v: '王', m: '王' },
      { v: 'A', m: 'A', fc: '#d84b4f' },
      { v: 'A-W', m: 'A-W', f: '=B6&"-W"' },
    ],
    [
      { v: '王', m: '王' },
      { v: 'a', m: 'a', bg: '#fff2cc' },
      { v: 'a-w', m: 'a-w', f: '=B7&"-w"' },
    ],
  ];

  return {
    id: SPREADSHEET_TEXT_SORT_ARTIFACT_ID,
    kind: 'spreadsheet',
    title: 'Chinese text sort workbook',
    favorite: false,
    createdAt: now,
    updatedAt: now,
    lastOpenedAt: now,
    revision: 1,
    content: {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'text-sort',
          name: 'Chinese names',
          status: 1,
          row: 20,
          column: 7,
          data,
          config: {
            columnlen: { 0: 110, 1: 110, 2: 150 },
          },
          luckysheet_select_save: [
            {
              row: [0, 6],
              column: [0, 2],
              row_focus: 1,
              column_focus: 0,
            },
          ],
        },
      ],
    },
  };
}
