import type { SpreadsheetContent } from '../../src/core';

export function spreadsheetCollaborationFixture(): SpreadsheetContent {
  return {
    type: 'spreadsheet',
    calculation: {
      mode: 'automatic',
      fullCalculationOnLoad: false,
      forceFullCalculation: false,
      iterativeCalculation: true,
      maximumIterations: 100,
      maximumChange: 0.001,
      fullPrecision: true,
    },
    sheets: [
      {
        id: 'sheet-input',
        name: 'Inputs',
        row: 3,
        column: 3,
        status: 1,
        zoomRatio: 1.25,
        luckysheet_select_save: [{ row: [0, 0], column: [0, 0] }],
        config: {
          merge: { '0_0': { r: 0, c: 0, rs: 1, cs: 2 } },
          rowlen: { '1': 28 },
          columnlen: { '0': 120 },
          borderInfo: [{ rangeType: 'cell', value: { row_index: 0 } }],
        },
        data: [
          [
            {
              v: 'Revenue',
              m: 'Revenue',
              bl: 1,
              bg: '#E2E8F0',
              ps: {
                left: null,
                top: null,
                width: 160,
                height: 90,
                value: 'Imported OOXML note',
                isShow: false,
              },
            },
            null,
            null,
          ],
          [
            { v: 10, m: '10', ct: { fa: '0.00', t: 'n' } },
            { v: 20, m: '20', ct: { fa: '0.00', t: 'n' } },
            {
              v: 30,
              m: '30',
              f: '=SUM(A2:B2)',
              ct: { fa: '0.00', t: 'n' },
            },
          ],
          [],
        ],
        formulaMetadata: {
          ranges: [
            {
              type: 'dynamic-array',
              anchor: 'C2',
              reference: 'C2:C3',
              formula: '=SUM(A2:B2)',
            },
          ],
          sourceFormulas: { C2: '=SUM(A2:B2)' },
        },
        images: [
          {
            id: 'image-logo',
            name: 'Logo',
            altText: 'Company logo',
            contentType: 'image/png',
            src: 'data:image/png;base64,AA==',
            left: 10,
            top: 10,
            width: 100,
            height: 40,
          },
        ],
        charts: [
          {
            id: 'chart-revenue',
            name: 'Revenue chart',
            type: 'column',
            categories: ['A', 'B'],
            series: [{ name: 'Revenue', values: [10, 20] }],
            showLegend: true,
            left: 240,
            top: 20,
            width: 480,
            height: 288,
          },
        ],
      },
      {
        id: 'sheet-summary',
        name: 'Summary',
        row: 2,
        column: 2,
        status: 0,
        celldata: [
          { r: 0, c: 0, v: { v: 'Total', m: 'Total' } },
          { r: 0, c: 1, v: { v: 30, m: '30', f: '=Inputs!C2' } },
        ],
        pivotTables: [
          {
            id: 'pivot-summary',
            name: 'Summary pivot',
            sourceSheetId: 'sheet-input',
            sourceReference: 'Inputs!A1:C2',
            anchor: 'A1',
            rowFields: [0],
            columnFields: [],
            values: [{ fieldIndex: 2, aggregation: 'sum' }],
            rowGrandTotals: true,
            columnGrandTotals: true,
            styleName: 'PivotStyleMedium2',
            refreshOnLoad: true,
          },
        ],
      },
    ],
    namedRanges: [
      {
        id: 'range-revenue',
        name: 'RevenueData',
        reference: 'Inputs!$A$2:$C$2',
        scopeSheetId: 'sheet-input',
        comment: 'Imported defined name',
      },
    ],
    printAreas: [{ sheetId: 'sheet-input', reference: '$A$1:$C$3' }],
    printTitles: [{ sheetId: 'sheet-input', rows: '$1:$1' }],
    pageBreaks: [{ sheetId: 'sheet-input', rows: [2], columns: [1] }],
    pageSetups: [
      {
        sheetId: 'sheet-input',
        paperSize: 'a4',
        orientation: 'landscape',
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
        header: { center: '&BQuarterly Plan' },
        footer: { right: 'Page &P of &N' },
        margins: {
          top: 0.75,
          right: 0.7,
          bottom: 0.75,
          left: 0.7,
          header: 0.3,
          footer: 0.3,
        },
      },
    ],
  };
}
