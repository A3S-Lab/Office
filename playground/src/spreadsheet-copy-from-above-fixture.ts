import type { OfficeArtifact, SpreadsheetContent } from '@a3s-lab/office/core';

export const SPREADSHEET_COPY_FROM_ABOVE_FIXTURE =
  'spreadsheet-copy-from-above';
export const SPREADSHEET_COPY_FROM_ABOVE_ARTIFACT_ID =
  'fixture-spreadsheet-copy-from-above';

export function createSpreadsheetCopyFromAboveArtifact(): OfficeArtifact {
  const now = Date.now();
  const data: NonNullable<SpreadsheetContent['sheets'][number]['data']> = [];
  data[0] = [
    { v: 3, m: '3' },
    {
      f: '=$A$1+A1',
      v: 6,
      m: '6',
      bg: '#fff2cc',
      ct: { fa: 'General', t: 'n' },
    },
  ];
  data[1] = [
    { v: 'Target', m: 'Target', bl: 1 },
    {
      v: 'Replace me',
      m: 'Replace me',
      bl: 1,
      it: 1,
      bg: '#ddebf7',
      ct: { fa: '0.00', t: 'n' },
    },
  ];
  data[3] = [
    { v: "Ctrl+'", m: "Ctrl+'", bl: 1 },
    {
      v: 'Copy the exact formula from above',
      m: 'Copy the exact formula from above',
    },
  ];
  data[4] = [
    { v: "Ctrl+Shift+'", m: "Ctrl+Shift+'", bl: 1 },
    {
      v: 'Copy the calculated value from above',
      m: 'Copy the calculated value from above',
    },
  ];

  return {
    id: SPREADSHEET_COPY_FROM_ABOVE_ARTIFACT_ID,
    kind: 'spreadsheet',
    title: 'Copy from above workbook',
    favorite: false,
    createdAt: now,
    updatedAt: now,
    lastOpenedAt: now,
    revision: 1,
    content: {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'copy-from-above-sheet',
          name: 'Copy from Above',
          status: 1,
          row: 16,
          column: 8,
          data,
          config: {
            columnlen: { 0: 144, 1: 276 },
          },
          luckysheet_select_save: [
            {
              row: [1, 1],
              column: [1, 1],
              row_focus: 1,
              column_focus: 1,
            },
          ],
        },
      ],
    },
  };
}
