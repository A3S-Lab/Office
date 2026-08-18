import type { OfficeArtifact, SpreadsheetContent } from '@a3s-lab/office/core';
import {
  withCellProtection,
  withSheetProtection,
} from '../../src/internal/features/work/work-spreadsheet-protection';

export const MAXIMUM_SPARSE_SPREADSHEET_FIXTURE = 'spreadsheet-sparse-maximum';
export const MAXIMUM_SPARSE_SPREADSHEET_ARTIFACT_ID =
  'fixture-maximum-sparse-spreadsheet';

const MAXIMUM_ROW_COUNT = 1_048_576;
const MAXIMUM_COLUMN_COUNT = 16_384;
const maximumRange = {
  row: [0, MAXIMUM_ROW_COUNT - 1] as [number, number],
  column: [0, MAXIMUM_COLUMN_COUNT - 1] as [number, number],
};

export function createMaximumSparseSpreadsheetArtifact(): OfficeArtifact {
  const data: NonNullable<SpreadsheetContent['sheets'][number]['data']> = [];
  data.length = MAXIMUM_ROW_COUNT;
  data[0] = [];
  data[0].length = MAXIMUM_COLUMN_COUNT;
  data[0][0] = { v: 'Anchor', m: 'Anchor' };
  data[MAXIMUM_ROW_COUNT - 1] = [];
  data[MAXIMUM_ROW_COUNT - 1].length = MAXIMUM_COLUMN_COUNT;
  data[MAXIMUM_ROW_COUNT - 1][MAXIMUM_COLUMN_COUNT - 1] = {
    v: 20,
    m: '20',
  };

  const baseSheet: SpreadsheetContent['sheets'][number] = {
    id: 'maximum-sparse-sheet',
    name: 'Maximum sparse sheet',
    status: 1,
    order: 0,
    row: MAXIMUM_ROW_COUNT,
    column: MAXIMUM_COLUMN_COUNT,
    data,
    dataValidationRanges: [
      {
        ranges: [maximumRange],
        item: {
          type: 'dropdown',
          type2: '',
          rangeTxt: 'A1:XFD1048576',
          value1: 'Ready,Blocked',
          value2: '',
          validity: '',
          remote: false,
          prohibitInput: false,
          hintShow: true,
          hintValue: 'Choose a workflow state.',
        },
      },
    ],
    luckysheet_conditionformat_save: [
      {
        type: 'default',
        cellrange: [maximumRange],
        format: { textColor: '#ffffff', cellColor: '#b42318' },
        conditionName: 'greaterThan',
        conditionValue: [10],
      },
    ],
  };
  const sheet = withSheetProtection(
    withCellProtection(baseSheet, [maximumRange], false),
    true,
  );
  const now = Date.now();
  return {
    id: MAXIMUM_SPARSE_SPREADSHEET_ARTIFACT_ID,
    kind: 'spreadsheet',
    title: 'Maximum sparse workbook',
    favorite: false,
    createdAt: now,
    updatedAt: now,
    lastOpenedAt: now,
    revision: 1,
    content: { type: 'spreadsheet', sheets: [sheet] },
  };
}
