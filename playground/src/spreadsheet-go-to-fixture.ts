import type { OfficeArtifact, SpreadsheetContent } from '@a3s-lab/office/core';

export const SPREADSHEET_GO_TO_FIXTURE = 'spreadsheet-go-to';
export const SPREADSHEET_GO_TO_ARTIFACT_ID = 'fixture-spreadsheet-go-to';

export function createSpreadsheetGoToArtifact(): OfficeArtifact {
  const now = Date.now();
  return {
    id: SPREADSHEET_GO_TO_ARTIFACT_ID,
    kind: 'spreadsheet',
    title: 'Go To navigation workbook',
    favorite: false,
    createdAt: now,
    updatedAt: now,
    lastOpenedAt: now,
    revision: 1,
    content: {
      type: 'spreadsheet',
      sheets: [
        spreadsheetGoToSheet(
          'go-to-inputs',
          'Inputs',
          true,
          'Input',
          '新版发布',
        ),
        spreadsheetGoToSheet('go-to-archive', 'Archive 2025', false, 'Archive'),
      ],
      namedRanges: [
        {
          id: 'go-to-archive-block',
          name: 'ArchiveBlock',
          reference: "'Archive 2025'!$C$9:$E$12",
        },
      ],
    },
  };
}

function spreadsheetGoToSheet(
  id: string,
  name: string,
  active: boolean,
  label: string,
  findText?: string,
): SpreadsheetContent['sheets'][number] {
  const data: NonNullable<SpreadsheetContent['sheets'][number]['data']> = [];
  data[0] = [{ v: label, m: label }];
  if (findText) data[4] = [{ v: findText, m: findText }];
  return {
    id,
    name,
    status: active ? 1 : 0,
    row: 40,
    column: 12,
    data,
  };
}
