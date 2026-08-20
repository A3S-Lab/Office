import type { OfficeArtifact, SpreadsheetContent } from '@a3s-lab/office/core';

export const SPREADSHEET_HYPERLINK_FIXTURE = 'spreadsheet-hyperlink';
export const SPREADSHEET_HYPERLINK_ARTIFACT_ID =
  'fixture-spreadsheet-hyperlink';

export function createSpreadsheetHyperlinkArtifact(): OfficeArtifact {
  const now = Date.now();
  const inputs: NonNullable<SpreadsheetContent['sheets'][number]['data']> = [];
  inputs[0] = [
    {
      v: 'A3S Office',
      m: 'A3S Office',
      bg: '#fff2cc',
      fc: '#172033',
      bl: 1,
      un: 0,
      ps: {
        left: null,
        top: null,
        width: null,
        height: null,
        value: 'Formatting and comments must survive hyperlink edits.',
        isShow: false,
      },
    },
  ];
  inputs[1] = [
    null,
    {
      f: '="Formula link"',
      v: 'Formula link',
      m: 'Formula link',
      bg: '#e2f0d9',
    },
  ];

  return {
    id: SPREADSHEET_HYPERLINK_ARTIFACT_ID,
    kind: 'spreadsheet',
    title: 'Hyperlink workbook',
    favorite: false,
    createdAt: now,
    updatedAt: now,
    lastOpenedAt: now,
    revision: 1,
    content: {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'hyperlink-inputs',
          name: 'Inputs',
          status: 1,
          row: 40,
          column: 12,
          data: inputs,
          luckysheet_select_save: [
            {
              row: [0, 0],
              column: [0, 0],
              row_focus: 0,
              column_focus: 0,
            },
          ],
        },
        {
          id: 'hyperlink-archive',
          name: 'Archive 2025',
          status: 0,
          row: 100,
          column: 20,
          data: [[{ v: 'Archive', m: 'Archive' }]],
        },
        {
          id: 'hyperlink-hidden',
          name: 'Hidden Archive',
          status: 0,
          hide: 1,
          row: 20,
          column: 8,
          data: [[{ v: 'Hidden', m: 'Hidden' }]],
        },
      ],
    },
  };
}
