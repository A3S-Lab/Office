import type { OfficeArtifact, SpreadsheetContent } from '@a3s-lab/office/core';

export const SPREADSHEET_RICH_TEXT_FIXTURE = 'spreadsheet-rich-text';
export const SPREADSHEET_RICH_TEXT_ARTIFACT_ID =
  'fixture-spreadsheet-rich-text';

export function createSpreadsheetRichTextArtifact(): OfficeArtifact {
  const now = Date.now();
  const data: NonNullable<SpreadsheetContent['sheets'][number]['data']> = [];
  data[0] = [
    {
      v: 'Native rich text',
      bl: 0,
      fc: '#172033',
      ff: 'Aptos',
      fs: 12,
      bg: '#f4f7ff',
      ct: {
        t: 'inlineStr',
        s: [
          {
            v: 'Native ',
            a3sXlsxColorOrigin: {
              baseColor: '#2f6fed',
              index: 4,
              kind: 'theme',
              renderedColor: '#2f6fed',
            },
            bl: 1,
            fc: '#2f6fed',
            ff: 'Aptos Display',
            fs: 14,
          },
          {
            v: 'rich ',
            a3sXlsxColorOrigin: {
              baseColor: '#159469',
              index: 10,
              kind: 'indexed',
              renderedColor: '#159469',
            },
            fc: '#159469',
            ff: 'Georgia',
            fs: 12,
            it: 1,
          },
          {
            v: 'text',
            a3sXlsxColorOrigin: {
              baseColor: '#000000',
              kind: 'automatic',
              renderedColor: '#000000',
            },
            cl: 1,
            fc: '#000000',
            un: 2,
          },
        ],
      },
    },
  ];
  data[2] = [
    { v: 'XLSX fidelity', bl: 1, bg: '#ddebf7' },
    {
      v: 'Shared and inline strings preserve native font runs',
      m: 'Shared and inline strings preserve native font runs',
    },
  ];
  data[3] = [
    { v: 'Format command', bl: 1, bg: '#e2f0d9' },
    {
      v: 'Font changes apply consistently to every visible run',
      m: 'Font changes apply consistently to every visible run',
    },
  ];

  return {
    id: SPREADSHEET_RICH_TEXT_ARTIFACT_ID,
    kind: 'spreadsheet',
    title: 'Native rich-text workbook',
    favorite: false,
    createdAt: now,
    updatedAt: now,
    lastOpenedAt: now,
    revision: 1,
    content: {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'rich-text-sheet',
          name: 'Rich Text',
          status: 1,
          row: 16,
          column: 8,
          data,
          config: {
            columnlen: { 0: 188, 1: 360 },
            rowlen: { 0: 36 },
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
