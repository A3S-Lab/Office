import { describe, expect, test } from '@rstest/core';
import {
  applySpreadsheetHyperlink,
  createSpreadsheetHyperlinkDialogSource,
  removeSpreadsheetHyperlink,
  validateSpreadsheetHyperlinkRequest,
} from '../src/internal/features/work/editors/spreadsheet-hyperlink';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

describe('spreadsheet hyperlinks', () => {
  test('adds and edits a webpage hyperlink without replacing cell content or formatting', () => {
    const content = hyperlinkContent();

    const inserted = applySpreadsheetHyperlink(content, {
      sheetId: 'sheet-1',
      row: 1,
      column: 1,
      linkType: 'webpage',
      linkAddress: 'a3s.dev/office',
    });

    expect(inserted).not.toBeNull();
    expect(inserted?.sheets[0]?.hyperlink).toEqual({
      '1_1': {
        linkType: 'webpage',
        linkAddress: 'https://a3s.dev/office',
      },
      '4_4': { linkType: 'webpage', linkAddress: 'https://keep.example' },
    });
    expect(inserted?.sheets[0]?.data?.[1]?.[1]).toEqual({
      v: 'A3S Office',
      m: 'A3S Office',
      f: '="A3S Office"',
      bg: '#fff2cc',
      fc: '#d84b4f',
      un: 0,
      ps: expect.objectContaining({ value: 'Keep comment' }),
      hl: { r: 1, c: 1, id: 'sheet-1' },
    });
    expect(content.sheets[0]?.data?.[1]?.[1]).not.toHaveProperty('hl');
    expect(content.sheets[0]?.hyperlink).not.toHaveProperty('1_1');

    if (!inserted) {
      throw new Error('Expected the hyperlink insertion to succeed');
    }
    const edited = applySpreadsheetHyperlink(inserted, {
      sheetId: 'sheet-1',
      row: 1,
      column: 1,
      linkType: 'webpage',
      linkAddress: 'https://docs.a3s.dev/office',
    });

    expect(edited?.sheets[0]?.hyperlink?.['1_1']).toEqual({
      linkType: 'webpage',
      linkAddress: 'https://docs.a3s.dev/office',
    });
    expect(edited?.sheets[0]?.data?.[1]?.[1]).toMatchObject({
      v: 'A3S Office',
      m: 'A3S Office',
      f: '="A3S Office"',
      bg: '#fff2cc',
      fc: '#d84b4f',
      un: 0,
      ps: expect.objectContaining({ value: 'Keep comment' }),
    });
  });

  test('changes display text only when explicitly requested', () => {
    const content = hyperlinkContent();
    delete content.sheets[0]?.data?.[1]?.[1]?.f;

    const next = applySpreadsheetHyperlink(content, {
      sheetId: 'sheet-1',
      row: 1,
      column: 1,
      linkType: 'sheet',
      linkAddress: 'Archive 2025',
      displayText: 'Open archive',
    });

    expect(next?.sheets[0]?.data?.[1]?.[1]).toMatchObject({
      v: 'Open archive',
      m: 'Open archive',
      bg: '#fff2cc',
      fc: '#d84b4f',
      un: 0,
      hl: { r: 1, c: 1, id: 'sheet-1' },
    });
    expect(next?.sheets[0]?.hyperlink?.['1_1']).toEqual({
      linkType: 'sheet',
      linkAddress: 'Archive 2025',
    });

    expect(
      applySpreadsheetHyperlink(hyperlinkContent(), {
        sheetId: 'sheet-1',
        row: 1,
        column: 1,
        linkType: 'webpage',
        linkAddress: 'https://a3s.dev',
        displayText: 'Replace formula',
      }),
    ).toBeNull();
  });

  test('preserves celldata storage for cross-sheet ranges and maximum cells', () => {
    const content = sparseHyperlinkContent();

    const next = applySpreadsheetHyperlink(content, {
      sheetId: 'sparse',
      row: 1_048_575,
      column: 16_383,
      linkType: 'cellrange',
      linkAddress: "'Archive 2025'!$C$9:$E$12",
      displayText: 'Open archive range',
    });

    expect(next).not.toBeNull();
    expect(next?.sheets[0]).not.toHaveProperty('data');
    expect(next?.sheets[0]?.celldata).toEqual([
      { r: 0, c: 0, v: { v: 'Keep', bg: '#d9ead3' } },
      {
        r: 1_048_575,
        c: 16_383,
        v: {
          v: 'Open archive range',
          m: 'Open archive range',
          hl: { r: 1_048_575, c: 16_383, id: 'sparse' },
        },
      },
    ]);
    expect(next?.sheets[0]?.hyperlink?.['1048575_16383']).toEqual({
      linkType: 'cellrange',
      linkAddress: "'Archive 2025'!C9:E12",
    });
    expect(content.sheets[0]?.celldata).toHaveLength(1);
  });

  test('removes only hyperlink state and preserves malformed vendor records', () => {
    const content = hyperlinkContent();
    const sheet = requireFirstSheet(content);
    const cell = requireDenseCell(content, 1, 1);
    sheet.hyperlink = {
      '1_1': { linkType: 'webpage', linkAddress: 'https://a3s.dev' },
      vendor: { linkType: 'vendor', linkAddress: 'opaque' },
    };
    const row = sheet.data?.[1];
    if (!row) {
      throw new Error('Expected the hyperlink fixture to contain row 1');
    }
    row[1] = {
      ...cell,
      hl: { r: 1, c: 1, id: 'sheet-1' },
    };

    const next = removeSpreadsheetHyperlink(content, {
      sheetId: 'sheet-1',
      row: 1,
      column: 1,
    });

    expect(next?.sheets[0]?.hyperlink).toEqual({
      vendor: { linkType: 'vendor', linkAddress: 'opaque' },
    });
    expect(next?.sheets[0]?.data?.[1]?.[1]).toEqual({
      v: 'A3S Office',
      m: 'A3S Office',
      f: '="A3S Office"',
      bg: '#fff2cc',
      fc: '#d84b4f',
      un: 0,
      ps: expect.objectContaining({ value: 'Keep comment' }),
    });
    expect(content.sheets[0]?.data?.[1]?.[1]).toHaveProperty('hl');
  });

  test('rejects protected, pivot, missing, hidden, invalid, and out-of-bounds targets', () => {
    const protectedContent = hyperlinkContent();
    requireFirstSheet(protectedContent).config = {
      authority: { sheet: 1 },
    };
    expect(
      validateSpreadsheetHyperlinkRequest(protectedContent, {
        sheetId: 'sheet-1',
        row: 1,
        column: 1,
        linkType: 'webpage',
        linkAddress: 'https://a3s.dev',
      }),
    ).toMatchObject({ ok: false, code: 'protected-cell' });

    const pivotContent = hyperlinkContent();
    requireFirstSheet(pivotContent).pivotTables = [
      {
        id: 'pivot-1',
        name: 'PivotTable1',
        sourceSheetId: 'sheet-1',
        sourceReference: 'A1:B4',
        anchor: 'B2',
        rowFields: [0],
        columnFields: [],
        values: [{ fieldIndex: 1, aggregation: 'sum' }],
        rowGrandTotals: true,
        columnGrandTotals: true,
        styleName: 'PivotStyleMedium9',
        refreshOnLoad: false,
        outputReference: 'B2:C4',
      },
    ];
    expect(
      validateSpreadsheetHyperlinkRequest(pivotContent, {
        sheetId: 'sheet-1',
        row: 1,
        column: 1,
        linkType: 'webpage',
        linkAddress: 'https://a3s.dev',
      }),
    ).toMatchObject({ ok: false, code: 'pivot-cell' });

    const content = hyperlinkContent();
    const cases = [
      [
        {
          sheetId: 'sheet-1',
          row: 1,
          column: 1,
          linkType: 'webpage',
          linkAddress: 'javascript:alert(1)',
        },
        'invalid-web-address',
      ],
      [
        {
          sheetId: 'sheet-1',
          row: 1,
          column: 1,
          linkType: 'sheet',
          linkAddress: 'Missing',
        },
        'target-sheet-not-found',
      ],
      [
        {
          sheetId: 'sheet-1',
          row: 1,
          column: 1,
          linkType: 'sheet',
          linkAddress: 'Hidden',
        },
        'target-sheet-hidden',
      ],
      [
        {
          sheetId: 'sheet-1',
          row: 1,
          column: 1,
          linkType: 'cellrange',
          linkAddress: "'Archive 2025'!U1",
        },
        'target-out-of-bounds',
      ],
      [
        {
          sheetId: 'sheet-1',
          row: 40,
          column: 0,
          linkType: 'webpage',
          linkAddress: 'https://a3s.dev',
        },
        'source-out-of-bounds',
      ],
    ] as const;

    for (const [request, code] of cases) {
      expect(
        validateSpreadsheetHyperlinkRequest(content, request),
      ).toMatchObject({ ok: false, code });
    }
  });

  test('creates an edit source from the exact active cell', () => {
    const content = hyperlinkContent();
    const sheet = requireFirstSheet(content);
    const cell = requireDenseCell(content, 1, 1);
    sheet.hyperlink = {
      '1_1': { linkType: 'sheet', linkAddress: 'Archive 2025' },
    };
    const row = sheet.data?.[1];
    if (!row) {
      throw new Error('Expected the hyperlink fixture to contain row 1');
    }
    row[1] = {
      ...cell,
      hl: { r: 1, c: 1, id: 'sheet-1' },
    };

    expect(
      createSpreadsheetHyperlinkDialogSource(content, {
        sheetId: 'sheet-1',
        row: 1,
        column: 1,
      }),
    ).toEqual({
      sheetId: 'sheet-1',
      sheetName: 'Sheet 1',
      row: 1,
      column: 1,
      cellReference: 'B2',
      displayText: 'A3S Office',
      displayTextEditable: false,
      hasHyperlink: true,
      link: { linkType: 'sheet', linkAddress: 'Archive 2025' },
      sheetOptions: [
        { id: 'sheet-1', name: 'Sheet 1' },
        { id: 'archive', name: 'Archive 2025' },
      ],
    });
  });
});

function hyperlinkContent(): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet 1',
        status: 1,
        row: 40,
        column: 12,
        data: [
          [{ v: 'Heading', bl: 1 }],
          [
            null,
            {
              v: 'A3S Office',
              m: 'A3S Office',
              f: '="A3S Office"',
              bg: '#fff2cc',
              fc: '#d84b4f',
              un: 0,
              ps: {
                left: null,
                top: null,
                width: null,
                height: null,
                value: 'Keep comment',
                isShow: false,
              },
            },
          ],
        ],
        hyperlink: {
          '4_4': {
            linkType: 'webpage',
            linkAddress: 'https://keep.example',
          },
        },
      },
      {
        id: 'archive',
        name: 'Archive 2025',
        row: 100,
        column: 20,
      },
      {
        id: 'hidden',
        name: 'Hidden',
        hide: 1,
        row: 40,
        column: 12,
      },
    ],
  };
}

function sparseHyperlinkContent(): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sparse',
        name: 'Sparse',
        status: 1,
        row: 1_048_576,
        column: 16_384,
        celldata: [{ r: 0, c: 0, v: { v: 'Keep', bg: '#d9ead3' } }],
      },
      {
        id: 'archive',
        name: 'Archive 2025',
        row: 100,
        column: 20,
      },
    ],
  };
}

function requireFirstSheet(
  content: WorkSpreadsheetContent,
): WorkSpreadsheetContent['sheets'][number] {
  const sheet = content.sheets[0];
  if (!sheet) {
    throw new Error('Expected the hyperlink fixture to contain a sheet');
  }
  return sheet;
}

function requireDenseCell(
  content: WorkSpreadsheetContent,
  rowIndex: number,
  columnIndex: number,
): NonNullable<
  NonNullable<
    NonNullable<WorkSpreadsheetContent['sheets'][number]['data']>[number]
  >[number]
> {
  const cell = requireFirstSheet(content).data?.[rowIndex]?.[columnIndex];
  if (!cell) {
    throw new Error(
      `Expected the hyperlink fixture to contain cell ${rowIndex},${columnIndex}`,
    );
  }
  return cell;
}
