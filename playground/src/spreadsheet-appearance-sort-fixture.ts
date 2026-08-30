import type { OfficeArtifact, SpreadsheetContent } from '@a3s-lab/office/core';

export const SPREADSHEET_APPEARANCE_SORT_FIXTURE =
  'spreadsheet-appearance-sort';
export const SPREADSHEET_APPEARANCE_SORT_ARTIFACT_ID =
  'fixture-spreadsheet-appearance-sort';

export function createSpreadsheetAppearanceSortArtifact(): OfficeArtifact {
  const now = Date.now();
  const data: NonNullable<SpreadsheetContent['sheets'][number]['data']> = [
    [
      { v: 'Task', m: 'Task', bl: 1, bg: '#d9ead3' },
      { v: 'Status', m: 'Status', bl: 1, bg: '#d9ead3' },
      { v: 'Score', m: 'Score', bl: 1, bg: '#d9ead3' },
      { v: 'Calculated', m: 'Calculated', bl: 1, bg: '#d9ead3' },
    ],
    [
      { v: 'Alpha', m: 'Alpha' },
      { v: 'Ready', m: 'Ready', bg: '#d9ead3' },
      { v: 70, m: '70' },
      { v: 140, m: '140', f: '=C2*2' },
    ],
    [
      { v: 'Zulu', m: 'Zulu', fc: '#d84b4f' },
      { v: 'Blocked', m: 'Blocked', bg: '#fce8e6' },
      { v: 90, m: '90' },
      { v: 180, m: '180', f: '=C3+90' },
    ],
    [
      { v: 'Beta', m: 'Beta' },
      { v: 'Blocked', m: 'Blocked', bg: '#fce8e6' },
      { v: 60, m: '60' },
      { v: 120, m: '120', f: '=C4*2' },
    ],
    [
      { v: 'Omega', m: 'Omega' },
      { v: 'Ready', m: 'Ready', bg: '#d9ead3' },
      { v: 95, m: '95' },
      { v: 190, m: '190', f: '=SUM(C5,95)' },
    ],
    [
      { v: 'Gamma', m: 'Gamma' },
      { v: 'Review', m: 'Review', bg: '#fff2cc' },
      { v: 80, m: '80' },
      { v: 160, m: '160', f: '=C6*2' },
    ],
  ];

  return {
    id: SPREADSHEET_APPEARANCE_SORT_ARTIFACT_ID,
    kind: 'spreadsheet',
    title: 'Appearance sort workbook',
    favorite: false,
    createdAt: now,
    updatedAt: now,
    lastOpenedAt: now,
    revision: 1,
    content: {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'appearance-sort',
          name: 'Visual priority',
          status: 1,
          row: 24,
          column: 8,
          data,
          config: {
            columnlen: { 0: 150, 1: 110, 2: 90, 3: 120 },
          },
          luckysheet_select_save: [
            {
              row: [0, 5],
              column: [0, 3],
              row_focus: 1,
              column_focus: 1,
            },
          ],
          luckysheet_conditionformat_save: [
            {
              type: 'icons',
              cellrange: [{ row: [1, 5], column: [2, 2] }],
              format: {
                iconSet: '3TrafficLights1',
                showValue: true,
                reverse: false,
                percent: false,
                thresholds: [
                  { type: 'min', gte: true },
                  { type: 'num', value: 75, gte: true },
                  { type: 'num', value: 85, gte: true },
                ],
              },
            },
          ],
        },
      ],
    },
  };
}
