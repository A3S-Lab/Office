import { expect, test } from '@rstest/core';
import { readFileSync } from 'node:fs';
import * as Y from 'yjs';
import {
  createOfficeCollaborationSession,
  createOfficeSpreadsheetCollaborationBinding,
  initializeOfficeSpreadsheetCollaboration,
  readOfficeSpreadsheetCollaboration,
} from '../src/core';
import { spreadsheetCollaborationFixture as fixture } from './fixtures/spreadsheet-collaboration';
import {
  NATIVE_SPREADSHEET_CREATE_CELL_BASE64,
  NATIVE_SPREADSHEET_DELETE_CELL_BASE64,
  NATIVE_SPREADSHEET_SET_CELL_BASE64,
} from './fixtures/native-spreadsheet-cell-updates';

const BROWSER_SPREADSHEET_FIXTURE_BASE64 = readFileSync(
  'tests/fixtures/browser-spreadsheet-collaboration-update.base64',
  'utf8',
).trim();

test('initializes sparse typed Spreadsheet roots without a workbook blob', () => {
  const session = spreadsheetSession('spreadsheet-typed');
  const expected = fixtureWithoutTransientViewState();

  expect(initializeOfficeSpreadsheetCollaboration(session, fixture())).toEqual({
    initialized: true,
    content: expected,
  });
  expect(readOfficeSpreadsheetCollaboration(session)).toEqual(expected);
  expect(
    session.document.getMap(session.rootName('spreadsheet.sheets')).size,
  ).toBe(2);
  expect(
    session.document.share.has(session.rootName('spreadsheet.content')),
  ).toBe(false);
  const sheets = session.document.getMap(
    session.rootName('spreadsheet.sheets'),
  );
  const input = sheets.get('sheet-input') as Y.Map<unknown>;
  expect(input).toBeInstanceOf(Y.Map);
  expect(input.has('data')).toBe(false);
  expect(input.has('celldata')).toBe(false);
  expect((input.get('cellPresence') as Y.Map<unknown>).size).toBe(4);
});

test('rejects duplicate identities before bootstrap metadata is written', () => {
  const session = spreadsheetSession('spreadsheet-duplicate');
  const content = fixture();

  expect(() =>
    initializeOfficeSpreadsheetCollaboration(session, {
      ...content,
      sheets: [content.sheets[0], { ...content.sheets[1], id: 'sheet-input' }],
    }),
  ).toThrow(/unique sheet ID/);
  expect(session.document.getMap(session.rootName('metadata')).size).toBe(0);
});

test('bounds dense matrices while retaining sparse Excel coordinates', () => {
  const denseSession = spreadsheetSession('spreadsheet-dense-bound');
  expect(() =>
    initializeOfficeSpreadsheetCollaboration(denseSession, {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'sheet-large-dense',
          name: 'Large dense',
          data: Array.from({ length: 1_001 }, () => Array(1_000).fill(null)),
        },
      ],
    }),
  ).toThrow(/materialized dense cells/);
  expect(
    denseSession.document.getMap(denseSession.rootName('metadata')).size,
  ).toBe(0);

  const sparseSession = spreadsheetSession('spreadsheet-sparse-bound');
  initializeOfficeSpreadsheetCollaboration(sparseSession, {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-large-sparse',
        name: 'Large sparse',
        row: 1_048_576,
        column: 16_384,
        celldata: [{ r: 1_048_575, c: 16_383, v: { v: 'edge' } }],
      },
    ],
  });
  expect(
    readOfficeSpreadsheetCollaboration(sparseSession).sheets[0].celldata,
  ).toEqual([{ r: 1_048_575, c: 16_383, v: { v: 'edge' } }]);
});

test('rejects malformed Spreadsheet roots without writes during read', () => {
  const session = spreadsheetSession('spreadsheet-pure-read');
  initializeOfficeSpreadsheetCollaboration(session, fixture());
  const sheets = session.document.getMap(
    session.rootName('spreadsheet.sheets'),
  );
  (sheets.get('sheet-input') as Y.Map<unknown>).delete('cellPresence');
  let transactions = 0;
  const countTransaction = () => {
    transactions += 1;
  };
  session.document.on('afterTransaction', countTransaction);

  expect(() => readOfficeSpreadsheetCollaboration(session)).toThrow(
    /sheet cell presence is invalid/,
  );
  expect(transactions).toBe(0);
});

test('merges formula, style, note, config, and metadata edits by object field', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'spreadsheet-field-convergence',
  );
  const firstBinding = createOfficeSpreadsheetCollaborationBinding(first);
  const secondBinding = createOfficeSpreadsheetCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const secondBefore = secondBinding.content();

  firstBinding.replace(
    firstBefore,
    updateInputCell(firstBefore, 1, 2, (cell) => ({
      ...cell,
      f: '=SUM(A2:B2)*2',
      v: 60,
      m: '60',
    })),
  );
  secondBinding.replace(
    secondBefore,
    updateInputSheet(secondBefore, (sheet) => ({
      ...sheet,
      config: {
        ...sheet.config,
        rowlen: { ...sheet.config?.rowlen, '2': 36 },
      },
      formulaMetadata: {
        ...sheet.formulaMetadata,
        sourceFormulas: {
          ...sheet.formulaMetadata?.sourceFormulas,
          D2: '=C2*2',
        },
      },
      data: sheet.data?.map((row, rowIndex) =>
        row.map((cell, column) =>
          rowIndex === 1 && column === 2 && cell
            ? {
                ...cell,
                bg: '#FEF3C7',
                ps: {
                  left: null,
                  top: null,
                  width: 120,
                  height: 80,
                  value: 'Remote note',
                  isShow: false,
                },
              }
            : cell,
        ),
      ),
    })),
  );
  exchangeUpdates(firstDocument, secondDocument);

  const converged = firstBinding.content();
  expect(secondBinding.content()).toEqual(converged);
  expect(converged.sheets[0].data?.[1]?.[2]).toMatchObject({
    f: '=SUM(A2:B2)*2',
    v: 60,
    m: '60',
    bg: '#FEF3C7',
    ps: { value: 'Remote note' },
  });
  expect(converged.sheets[0].config?.rowlen).toMatchObject({
    '1': 28,
    '2': 36,
  });
  expect(converged.sheets[0].formulaMetadata?.sourceFormulas).toMatchObject({
    C2: '=SUM(A2:B2)',
    D2: '=C2*2',
  });
});

test('converges concurrent first writes to one blank cell by nested field', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'spreadsheet-blank-cell-convergence',
  );
  const firstBinding = createOfficeSpreadsheetCollaborationBinding(first);
  const secondBinding = createOfficeSpreadsheetCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const secondBefore = secondBinding.content();

  firstBinding.replace(
    firstBefore,
    updateInputCell(firstBefore, 2, 1, () => ({
      f: '=A2+B2',
      v: 30,
      m: '30',
    })),
  );
  secondBinding.replace(
    secondBefore,
    updateInputCell(secondBefore, 2, 1, () => ({
      bg: '#DBEAFE',
      ct: { fa: '$0.00', t: 'n' },
      ps: {
        left: null,
        top: null,
        width: 140,
        height: 80,
        value: 'Concurrent note',
        isShow: false,
      },
    })),
  );
  exchangeUpdates(firstDocument, secondDocument);

  const converged = firstBinding.content();
  expect(secondBinding.content()).toEqual(converged);
  expect(converged.sheets[0].data?.[2]?.[1]).toMatchObject({
    f: '=A2+B2',
    v: 30,
    m: '30',
    bg: '#DBEAFE',
    ct: { fa: '$0.00', t: 'n' },
    ps: { value: 'Concurrent note' },
  });
});

test('merges different fields in the same OOXML merge definition', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'spreadsheet-merge-config-convergence',
  );
  const firstBinding = createOfficeSpreadsheetCollaborationBinding(first);
  const secondBinding = createOfficeSpreadsheetCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const secondBefore = secondBinding.content();

  firstBinding.replace(
    firstBefore,
    updateInputSheet(firstBefore, (sheet) => ({
      ...sheet,
      config: {
        ...sheet.config,
        merge: {
          ...sheet.config?.merge,
          '0_0': { ...sheet.config?.merge?.['0_0'], rs: 2 } as {
            r: number;
            c: number;
            rs: number;
            cs: number;
          },
        },
      },
    })),
  );
  secondBinding.replace(
    secondBefore,
    updateInputSheet(secondBefore, (sheet) => ({
      ...sheet,
      config: {
        ...sheet.config,
        merge: {
          ...sheet.config?.merge,
          '0_0': { ...sheet.config?.merge?.['0_0'], cs: 3 } as {
            r: number;
            c: number;
            rs: number;
            cs: number;
          },
        },
      },
    })),
  );
  exchangeUpdates(firstDocument, secondDocument);

  expect(firstBinding.content().sheets[0].config?.merge?.['0_0']).toEqual({
    r: 0,
    c: 0,
    rs: 2,
    cs: 3,
  });
  expect(secondBinding.content()).toEqual(firstBinding.content());
});

test('preserves remotely added cells when applying an unrelated stale snapshot', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'spreadsheet-stale-cell',
  );
  const firstBinding = createOfficeSpreadsheetCollaborationBinding(first);
  const secondBinding = createOfficeSpreadsheetCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const stale = secondBinding.content();

  firstBinding.replace(
    firstBefore,
    updateInputCell(firstBefore, 2, 0, () => ({ v: 40, m: '40' })),
  );
  exchangeUpdates(firstDocument, secondDocument);
  secondBinding.replace(
    stale,
    updateInputCell(stale, 1, 0, (cell) => ({ ...cell, v: 11, m: '11' })),
  );
  exchangeUpdates(firstDocument, secondDocument);

  const converged = firstBinding.content();
  expect(secondBinding.content()).toEqual(converged);
  expect(converged.sheets[0].data?.[2]?.[0]).toMatchObject({ v: 40 });
  expect(converged.sheets[0].data?.[1]?.[0]).toMatchObject({ v: 11 });
});

test('rejects deleting a cell that was concurrently edited', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'spreadsheet-delete-edit',
  );
  const firstBinding = createOfficeSpreadsheetCollaborationBinding(first);
  const secondBinding = createOfficeSpreadsheetCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const stale = secondBinding.content();

  firstBinding.replace(
    firstBefore,
    updateInputCell(firstBefore, 1, 0, (cell) => ({ ...cell, v: 99, m: '99' })),
  );
  applyMissingUpdate(firstDocument, secondDocument);

  expect(() =>
    secondBinding.replace(
      stale,
      updateInputCell(stale, 1, 0, () => null),
    ),
  ).toThrow(/changed concurrently/);
});

test('rejects stale conflicts before writing any part of the transaction', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'spreadsheet-atomic-conflict',
  );
  const firstBinding = createOfficeSpreadsheetCollaborationBinding(first);
  const secondBinding = createOfficeSpreadsheetCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const stale = secondBinding.content();

  firstBinding.replace(
    firstBefore,
    updateInputCell(firstBefore, 1, 0, (cell) => ({ ...cell, v: 99, m: '99' })),
  );
  applyMissingUpdate(firstDocument, secondDocument);
  const updateBeforeFailure = Y.encodeStateVector(secondDocument);
  expect(() =>
    secondBinding.replace(
      stale,
      updateInputCell(
        updateInputCell(stale, 1, 0, () => null),
        1,
        1,
        (cell) => ({ ...cell, v: 25, m: '25' }),
      ),
    ),
  ).toThrow(/changed concurrently/);

  expect(Y.encodeStateVector(secondDocument)).toEqual(updateBeforeFailure);
  expect(secondBinding.content().sheets[0].data?.[1]?.[0]?.v).toBe(99);
  expect(secondBinding.content().sheets[0].data?.[1]?.[1]?.v).toBe(20);
});

test('rejects concurrent reuse of the same sheet ID for different sheets', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'spreadsheet-sheet-id-claim',
  );
  const firstBinding = createOfficeSpreadsheetCollaborationBinding(first);
  const secondBinding = createOfficeSpreadsheetCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const secondBefore = secondBinding.content();

  firstBinding.replace(firstBefore, {
    ...firstBefore,
    sheets: [...firstBefore.sheets, addedSheet('sheet-collision', 'Ada sheet')],
  });
  secondBinding.replace(secondBefore, {
    ...secondBefore,
    sheets: [
      ...secondBefore.sheets,
      addedSheet('sheet-collision', 'Grace sheet'),
    ],
  });
  exchangeUpdates(firstDocument, secondDocument);

  expect(() => firstBinding.content()).toThrow(/concurrently assigned/);
  expect(() => secondBinding.content()).toThrow(/concurrently assigned/);
});

test('detects concurrent independent Spreadsheet bootstrap', () => {
  const firstDocument = new Y.Doc();
  const secondDocument = new Y.Doc();
  const first = spreadsheetSession('spreadsheet-bootstrap-race', firstDocument);
  const second = spreadsheetSession(
    'spreadsheet-bootstrap-race',
    secondDocument,
  );
  initializeOfficeSpreadsheetCollaboration(first, fixture());
  const secondFixture = fixture();
  initializeOfficeSpreadsheetCollaboration(second, {
    ...secondFixture,
    calculation: {
      ...(secondFixture.calculation ?? {
        mode: 'automatic',
        fullCalculationOnLoad: false,
        forceFullCalculation: false,
        iterativeCalculation: false,
        maximumIterations: 100,
        maximumChange: 0.001,
        fullPrecision: true,
      }),
      maximumIterations: 200,
    },
  });

  exchangeUpdates(firstDocument, secondDocument);

  expect(() => first.metadata()).toThrow(/Multiple clients initialized/);
  expect(() => second.metadata()).toThrow(/Multiple clients initialized/);
});

test('keeps Spreadsheet undo local to one client', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'spreadsheet-local-undo',
  );
  const firstBinding = createOfficeSpreadsheetCollaborationBinding(first);
  const secondBinding = createOfficeSpreadsheetCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  firstBinding.replace(
    firstBefore,
    updateInputCell(firstBefore, 1, 0, (cell) => ({ ...cell, v: 12, m: '12' })),
  );
  exchangeUpdates(firstDocument, secondDocument);
  const secondBefore = secondBinding.content();
  secondBinding.stopCapturing();
  secondBinding.replace(
    secondBefore,
    updateInputCell(secondBefore, 1, 1, (cell) => ({
      ...cell,
      v: 25,
      m: '25',
    })),
  );
  exchangeUpdates(firstDocument, secondDocument);

  expect(firstBinding.undo()).toBe(true);
  exchangeUpdates(firstDocument, secondDocument);
  expect(firstBinding.content().sheets[0].data?.[1]?.[0]?.v).toBe(10);
  expect(firstBinding.content().sheets[0].data?.[1]?.[1]?.v).toBe(25);
  expect(secondBinding.content()).toEqual(firstBinding.content());
});

test('rejects Spreadsheet mutation outside edit mode', () => {
  const document = new Y.Doc();
  const writable = spreadsheetSession('spreadsheet-view', document);
  initializeOfficeSpreadsheetCollaboration(writable, fixture());
  const readOnly = createOfficeCollaborationSession({
    artifactId: 'spreadsheet-view',
    document,
    kind: 'spreadsheet',
    mode: 'view',
  });
  const binding = createOfficeSpreadsheetCollaborationBinding(readOnly);
  const before = binding.content();

  expect(() =>
    binding.replace(
      before,
      updateInputCell(before, 1, 0, () => ({ v: 1 })),
    ),
  ).toThrow(/cannot modify canonical content/);
  expect(() => binding.undo()).toThrow(/cannot modify canonical content/);
});

test('applies native Spreadsheet cell updates in Yjs across reordered delivery', () => {
  const orderedDocument = new Y.Doc();
  const reorderedDocument = new Y.Doc();
  for (const document of [orderedDocument, reorderedDocument]) {
    Y.applyUpdate(document, decodeBase64(BROWSER_SPREADSHEET_FIXTURE_BASE64));
    const binding = createOfficeSpreadsheetCollaborationBinding(
      spreadsheetSession('fixture-spreadsheet', document),
    );
    const before = binding.content();
    binding.replace(before, {
      ...before,
      sheets: before.sheets.map((sheet) => {
        if (sheet.id !== 'sheet-data') return sheet;
        const data = (sheet.data ?? []).map((row) => [...row]);
        const cell = data[1]?.[0];
        if (!cell) throw new Error('Expected the browser fixture data cell.');
        data[1][0] = {
          ...cell,
          bg: '#DBEAFE',
          ps: {
            left: null,
            top: null,
            width: 140,
            height: 80,
            value: 'Browser note',
            isShow: false,
          },
        };
        return { ...sheet, data };
      }),
    });
  }

  for (const encoded of [
    NATIVE_SPREADSHEET_SET_CELL_BASE64,
    NATIVE_SPREADSHEET_CREATE_CELL_BASE64,
    NATIVE_SPREADSHEET_DELETE_CELL_BASE64,
  ]) {
    Y.applyUpdate(orderedDocument, decodeBase64(encoded));
  }
  for (const encoded of [
    NATIVE_SPREADSHEET_DELETE_CELL_BASE64,
    NATIVE_SPREADSHEET_CREATE_CELL_BASE64,
    NATIVE_SPREADSHEET_SET_CELL_BASE64,
    NATIVE_SPREADSHEET_SET_CELL_BASE64,
  ]) {
    Y.applyUpdate(reorderedDocument, decodeBase64(encoded));
  }

  const contents = [orderedDocument, reorderedDocument].map((document) =>
    readOfficeSpreadsheetCollaboration(
      spreadsheetSession('fixture-spreadsheet', document),
    ),
  );
  expect(contents[0]).toEqual(contents[1]);
  const dataSheet = contents[0]?.sheets.find(({ id }) => id === 'sheet-data');
  expect(dataSheet?.data?.[1]?.[0]).toMatchObject({
    v: 12,
    m: '12',
    f: '=6*2',
    bg: '#DBEAFE',
    ct: { fa: '0.00', t: 'n' },
    ps: { value: 'Browser note' },
  });
  const emptySheet = contents[0]?.sheets.find(({ id }) => id === 'sheet-empty');
  expect(emptySheet?.data).toBeUndefined();
  expect(emptySheet?.celldata).toEqual([
    {
      r: 100,
      c: 5,
      v: {
        v: 'sparse native',
        m: 'sparse native',
        ps: { value: 'Agent note', isShow: false },
      },
    },
  ]);
  const sparseSheet = contents[0]?.sheets.find(
    ({ id }) => id === 'sheet-sparse',
  );
  expect(sparseSheet?.celldata).toEqual([]);
});

function fixtureWithoutTransientViewState() {
  const content = fixture();
  return {
    ...content,
    sheets: content.sheets.map(
      ({
        status: _status,
        zoomRatio: _zoomRatio,
        luckysheet_select_save: _selection,
        ...sheet
      }) => sheet,
    ),
  };
}

function spreadsheetSession(artifactId: string, document = new Y.Doc()) {
  return createOfficeCollaborationSession({
    artifactId,
    document,
    kind: 'spreadsheet',
  });
}

function connectedPair(artifactId: string) {
  const firstDocument = new Y.Doc();
  const first = spreadsheetSession(artifactId, firstDocument);
  initializeOfficeSpreadsheetCollaboration(first, fixture());
  const secondDocument = new Y.Doc();
  Y.applyUpdate(secondDocument, Y.encodeStateAsUpdate(firstDocument));
  const second = spreadsheetSession(artifactId, secondDocument);
  return { first, firstDocument, second, secondDocument };
}

function updateInputCell(
  content: ReturnType<typeof fixtureWithoutTransientViewState>,
  row: number,
  column: number,
  update: (
    cell: NonNullable<
      NonNullable<(typeof content.sheets)[number]['data']>[number][number]
    >,
  ) => NonNullable<(typeof content.sheets)[number]['data']>[number][number],
) {
  return updateInputSheet(content, (sheet) => {
    const data = (sheet.data ?? []).map((values) => [...values]);
    while (data.length <= row) data.push([]);
    while (data[row].length <= column) data[row].push(null);
    data[row][column] = update(data[row][column] ?? {});
    return { ...sheet, data };
  });
}

function updateInputSheet(
  content: ReturnType<typeof fixtureWithoutTransientViewState>,
  update: (
    sheet: (typeof content.sheets)[number],
  ) => (typeof content.sheets)[number],
) {
  return {
    ...content,
    sheets: content.sheets.map((sheet) =>
      sheet.id === 'sheet-input' ? update(sheet) : sheet,
    ),
  };
}

function addedSheet(id: string, name: string) {
  return {
    id,
    name,
    row: 1,
    column: 1,
    data: [[{ v: name, m: name }]],
  };
}

function exchangeUpdates(first: Y.Doc, second: Y.Doc): void {
  const firstUpdate = Y.encodeStateAsUpdate(first, Y.encodeStateVector(second));
  const secondUpdate = Y.encodeStateAsUpdate(
    second,
    Y.encodeStateVector(first),
  );
  Y.applyUpdate(first, secondUpdate, 'test-network');
  Y.applyUpdate(second, firstUpdate, 'test-network');
}

function applyMissingUpdate(source: Y.Doc, target: Y.Doc): void {
  Y.applyUpdate(
    target,
    Y.encodeStateAsUpdate(source, Y.encodeStateVector(target)),
    'test-network',
  );
}

function decodeBase64(value: string): Uint8Array {
  return Uint8Array.from(Buffer.from(value, 'base64'));
}
