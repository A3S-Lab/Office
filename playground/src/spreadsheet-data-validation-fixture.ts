import type { OfficeArtifact, SpreadsheetContent } from '@a3s-lab/office/core';

export const SPREADSHEET_DATA_VALIDATION_FIXTURE =
  'spreadsheet-data-validation';
export const SPREADSHEET_DATA_VALIDATION_ARTIFACT_ID =
  'fixture-spreadsheet-data-validation';

export function createSpreadsheetDataValidationArtifact(): OfficeArtifact {
  const now = Date.now();
  const inputs: NonNullable<SpreadsheetContent['sheets'][number]['data']> = [];
  inputs[0] = [
    { v: 'Owner', m: 'Owner', bl: 1, bg: '#e2f0d9' },
    { v: 'State', m: 'State', bl: 1, bg: '#e2f0d9' },
    { v: 'Due date', m: 'Due date', bl: 1, bg: '#e2f0d9' },
    { v: 'Priority', m: 'Priority', bl: 1, bg: '#e2f0d9' },
    { v: 'Review', m: 'Review', bl: 1, bg: '#e2f0d9' },
  ];
  inputs[1] = [
    { v: 'Avery', m: 'Avery' },
    { v: 'Ready', m: 'Ready', bg: '#fff2cc' },
    { v: '2026-08-21', m: '2026-08-21' },
  ];
  inputs[2] = [
    { v: 'Morgan', m: 'Morgan' },
    { v: 'Blocked', m: 'Blocked', bg: '#fce8e6' },
  ];
  inputs[4] = [
    null,
    null,
    null,
    { v: 'Ready', m: 'Ready', bg: '#fff2cc' },
    { v: 'Blocked', m: 'Blocked', bg: '#fce8e6' },
  ];

  return {
    id: SPREADSHEET_DATA_VALIDATION_ARTIFACT_ID,
    kind: 'spreadsheet',
    title: 'Data validation workbook',
    favorite: false,
    createdAt: now,
    updatedAt: now,
    lastOpenedAt: now,
    revision: 1,
    content: {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'validation-inputs',
          name: 'Inputs',
          status: 1,
          row: 40,
          column: 12,
          data: inputs,
          luckysheet_select_save: [
            {
              row: [1, 2],
              column: [1, 1],
              row_focus: 1,
              column_focus: 1,
            },
            {
              row: [4, 4],
              column: [3, 4],
              row_focus: 4,
              column_focus: 4,
            },
          ],
          dataValidationRanges: [
            {
              ranges: [{ row: [1, 2], column: [2, 2] }],
              item: {
                type: 'date',
                type2: 'between',
                rangeTxt: 'C2:C3',
                value1: '2026-01-01',
                value2: '2026-12-31',
                validity: '',
                remote: false,
                allowBlank: true,
                showDropdownArrow: true,
                prohibitInput: true,
                errorStyle: 'information',
                errorTitle: 'Date outside 2026',
                errorMessage: 'Enter a date in calendar year 2026.',
                hintShow: true,
                hintTitle: 'Due date',
                hintValue: 'Use a date between 2026-01-01 and 2026-12-31.',
                checked: false,
              },
            },
          ],
        },
        {
          id: 'validation-lists',
          name: 'Lists',
          status: 0,
          row: 20,
          column: 4,
          data: [
            [{ v: 'Ready', m: 'Ready' }],
            [{ v: 'Blocked', m: 'Blocked' }],
            [{ v: 'In review', m: 'In review' }],
          ],
        },
      ],
    },
  };
}
