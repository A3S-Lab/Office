import type { OfficeArtifact, SpreadsheetContent } from '@a3s-lab/office/core';

export const SPREADSHEET_PASTE_SPECIAL_FIXTURE = 'spreadsheet-paste-special';
export const SPREADSHEET_PASTE_SPECIAL_ARTIFACT_ID =
  'fixture-spreadsheet-paste-special';

export function createSpreadsheetPasteSpecialArtifact(): OfficeArtifact {
  const now = Date.now();
  const data: NonNullable<SpreadsheetContent['sheets'][number]['data']> = [];
  data[0] = [
    {
      v: 10,
      m: '10',
      bg: '#fff2cc',
      bl: 1,
      ct: { fa: '0.00', t: 'n' },
      ps: {
        left: null,
        top: null,
        width: null,
        height: null,
        value: 'Source value',
        isShow: false,
      },
    },
    {
      v: 20,
      m: '20',
      f: '=$A1+A$1',
      bg: '#e2f0d9',
      ct: { fa: '0%', t: 'n' },
    },
  ];
  data[2] = [];
  data[2][2] = { v: 100, m: '100', bg: '#ddebf7', fs: 16 };
  data[2][3] = { v: 200, m: '200', bg: '#fce4d6', fs: 16 };

  return {
    id: SPREADSHEET_PASTE_SPECIAL_ARTIFACT_ID,
    kind: 'spreadsheet',
    title: 'Paste Special workbook',
    favorite: false,
    createdAt: now,
    updatedAt: now,
    lastOpenedAt: now,
    revision: 1,
    content: {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'paste-special-sheet',
          name: 'Paste Lab',
          status: 1,
          row: 20,
          column: 12,
          data,
          config: {
            columnlen: { 0: 108, 1: 108, 2: 108, 3: 108 },
          },
          luckysheet_select_save: [
            {
              row: [0, 0],
              column: [0, 0],
              row_focus: 0,
              column_focus: 0,
            },
          ],
        },
      ],
    },
  };
}
