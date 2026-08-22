import type { OfficeArtifact, SpreadsheetContent } from '@a3s-lab/office/core';

export const SPREADSHEET_DATE_TIME_FIXTURE = 'spreadsheet-date-time';
export const SPREADSHEET_DATE_TIME_ARTIFACT_ID =
  'fixture-spreadsheet-date-time';

export function createSpreadsheetDateTimeArtifact(): OfficeArtifact {
  const now = Date.now();
  const data: NonNullable<SpreadsheetContent['sheets'][number]['data']> = [];
  data[0] = [
    {
      v: '传统 Office 日期与时间快捷输入',
      m: '传统 Office 日期与时间快捷输入',
      bl: 1,
      bg: '#e2f0d9',
    },
  ];
  data[1] = [
    { v: '当前日期（Ctrl+;）', m: '当前日期（Ctrl+;）', bl: 1 },
    { v: '等待输入', m: '等待输入' },
  ];
  data[2] = [
    { v: '当前时间（Ctrl+Shift+;）', m: '当前时间（Ctrl+Shift+;）', bl: 1 },
    { v: '等待输入', m: '等待输入' },
  ];
  data[4] = [
    { v: '语义', m: '语义', bl: 1, bg: '#ddebf7' },
    {
      v: '写入静态值；重新计算不会改变结果',
      m: '写入静态值；重新计算不会改变结果',
    },
  ];

  return {
    id: SPREADSHEET_DATE_TIME_ARTIFACT_ID,
    kind: 'spreadsheet',
    title: 'Date and time shortcut workbook',
    favorite: false,
    createdAt: now,
    updatedAt: now,
    lastOpenedAt: now,
    revision: 1,
    content: {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'date-time-sheet',
          name: 'Date and Time',
          status: 1,
          row: 20,
          column: 8,
          data,
          config: {
            columnlen: { 0: 188, 1: 248 },
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
