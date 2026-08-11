import { Editor } from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { describe, expect, test } from '@rstest/core';
import { waitFor } from '@testing-library/react';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  DocumentPagination,
  measureDocumentLayoutBlocks,
} from '../src/internal/features/work/work-document-pagination';
import {
  canSetDocumentTableRowRepeatHeader,
  documentTableRowOptions,
} from '../src/internal/features/work/work-document-table-row';

function createTableEditor(): Editor {
  return new Editor({
    extensions: createWorkDocumentExtensions(),
    content: [
      '<section data-document-section="true">',
      '<table><tbody>',
      '<tr data-office-repeat-header="true" data-office-cant-split="true">',
      '<th><p>标题</p></th><th><p>负责人</p></th>',
      '</tr>',
      '<tr><td><p>方案</p></td><td><p>A3S</p></td></tr>',
      '</tbody></table>',
      '</section>',
    ].join(''),
  });
}

describe('document tables', () => {
  test('inserts the chosen table without replacing selected text', async () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true">',
        '<p>Keep this selected text.</p>',
        '<p>Continue here.</p>',
        '</section>',
      ].join(''),
    });
    const editorElement = editor.view.dom;
    document.body.append(editorElement);
    const selection = documentTextRange(editor, 'this selected');
    editor.commands.setTextSelection(selection);
    editor.commands.blur();
    let updateCount = 0;
    const countUpdate = () => {
      updateCount += 1;
    };
    editor.on('update', countUpdate);

    expect(
      editor.commands.insertDocumentTable({
        rows: 2,
        columns: 4,
      }),
    ).toBe(true);

    expect(editor.getText()).toContain('Keep this selected text.');
    expect(documentTableShape(editor)).toEqual([
      ['tableHeader', 'tableHeader', 'tableHeader', 'tableHeader'],
      ['tableCell', 'tableCell', 'tableCell', 'tableCell'],
    ]);
    expect(updateCount).toBe(1);
    await waitFor(() => expect(editor.isFocused).toBe(true));

    expect(editor.commands.undo()).toBe(true);
    expect(documentTableShape(editor)).toEqual([]);
    expect(editor.getText()).toContain('Keep this selected text.');

    editor.off('update', countUpdate);
    editor.destroy();
    editorElement.remove();
  });

  test('inserts and edits a nested table inside the active cell', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true">',
        '<table><tbody><tr><td><p>Outer cell</p></td></tr></tbody></table>',
        '</section>',
      ].join(''),
    });
    editor.commands.setTextSelection(
      documentTextRange(editor, 'Outer cell').to,
    );

    expect(
      editor.commands.insertDocumentTable(
        { rows: 2, columns: 2 },
        { headerRow: false, restoreFocus: false },
      ),
    ).toBe(true);
    const tables: ProseMirrorNode[] = [];
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'table') tables.push(node);
      return true;
    });
    expect(tables).toHaveLength(2);
    expect(tables[0]?.childCount).toBe(1);
    expect(tables[1]?.childCount).toBe(2);
    expect(editor.commands.setDocumentTableColumnWidth(90)).toBe(true);
    expect(tables[0]?.firstChild?.firstChild?.attrs.colwidth).toBeNull();
    const nestedWidths: unknown[] = [];
    editor.state.doc.descendants((node) => {
      if (node.type.name !== 'table') return true;
      if (!node.textContent.includes('Outer cell')) {
        node.forEach((row) => {
          nestedWidths.push(row.firstChild?.attrs.colwidth);
        });
      }
      return true;
    });
    expect(nestedWidths).toEqual([[90], [90]]);

    editor.destroy();
  });

  test('keeps row pagination properties in the TipTap document model', () => {
    const editor = createTableEditor();
    editor.commands.setTextSelection(4);

    expect(documentTableRowOptions(editor)).toEqual({
      cantSplit: true,
      repeatHeader: true,
    });
    expect(editor.getHTML()).toContain('data-office-cant-split="true"');
    expect(editor.getHTML()).toContain('data-office-repeat-header="true"');

    editor.destroy();
  });

  test('updates row pagination properties without rebuilding the table', () => {
    const editor = createTableEditor();
    editor.commands.setTextSelection(4);

    expect(
      editor.commands.setDocumentTableRowOptions(
        { cantSplit: false, repeatHeader: false },
        { restoreFocus: false },
      ),
    ).toBe(true);
    expect(documentTableRowOptions(editor)).toEqual({
      cantSplit: false,
      repeatHeader: false,
    });
    expect(editor.getHTML()).toContain('data-office-cant-split="false"');
    expect(editor.getHTML()).toContain('data-office-repeat-header="false"');

    editor.destroy();
  });

  test('offers repeated headers only for a contiguous leading row group', () => {
    const editor = createTableEditor();
    editor.commands.setTextSelection(4);
    editor.chain().addRowAfter().run();

    editor.commands.setTextSelection(tableRowTextPosition(editor, 1));
    expect(canSetDocumentTableRowRepeatHeader(editor)).toBe(true);

    editor.commands.setTextSelection(tableRowTextPosition(editor, 2));
    expect(canSetDocumentTableRowRepeatHeader(editor)).toBe(false);

    editor.destroy();
  });

  test('measures table rows as a paginated flow with a repeated header', () => {
    const editor = createTableEditor();
    const wrapper = editor.view.dom.querySelector<HTMLElement>('.tableWrapper');
    const rows = Array.from(
      editor.view.dom.querySelectorAll<HTMLElement>(':scope tr'),
    );
    if (!wrapper || rows.length !== 2)
      throw new Error('Expected the mounted TipTap table.');
    Object.defineProperty(wrapper, 'offsetHeight', {
      configurable: true,
      value: 52,
    });
    Object.defineProperty(rows[0], 'offsetHeight', {
      configurable: true,
      value: 20,
    });
    Object.defineProperty(rows[1], 'offsetHeight', {
      configurable: true,
      value: 30,
    });

    const snapshot = measureDocumentLayoutBlocks(editor);

    expect(snapshot.blocks).toHaveLength(2);
    expect(snapshot.blocks.map(({ block }) => block)).toMatchObject([
      {
        id: expect.stringContaining('-row-0'),
        height: 20,
        flowIndex: 0,
        flowCount: 2,
        repeatHeaderCount: 1,
        repeatHeaderHeight: 20,
      },
      {
        id: expect.stringContaining('-row-1'),
        height: 32,
        flowIndex: 1,
        flowCount: 2,
        repeatHeaderCount: 1,
        repeatHeaderHeight: 20,
      },
    ]);
    expect(snapshot.blocks[1]?.tableBreak?.repeatedHeaderRowsHtml[0]).toContain(
      '标题',
    );

    editor.destroy();
  });

  test('splits an allowed row at synchronized cell block boundaries', () => {
    const editor = createSplitRowEditor(false);
    mockSplitRowGeometry(editor);

    const snapshot = measureDocumentLayoutBlocks(
      editor,
      null,
      0,
      new Map(),
      100,
    );

    expect(snapshot.blocks).toHaveLength(3);
    expect(snapshot.blocks.map(({ block }) => block.height)).toEqual([
      20, 40, 40,
    ]);
    expect(snapshot.blocks[1]?.block).toMatchObject({
      flowIndex: 1,
      flowCount: 3,
      repeatHeaderCount: 1,
      repeatHeaderHeight: 20,
    });
    expect(snapshot.blocks[2]?.tableBreak?.cellBreaks).toHaveLength(2);
    expect(snapshot.blocks[2]?.selectionRanges).toHaveLength(2);

    editor.destroy();
  });

  test('keeps a cant-split row atomic during measurement', () => {
    const editor = createSplitRowEditor(true);
    mockSplitRowGeometry(editor);

    const snapshot = measureDocumentLayoutBlocks(
      editor,
      null,
      0,
      new Map(),
      100,
    );

    expect(snapshot.blocks).toHaveLength(2);
    expect(snapshot.blocks.map(({ block }) => block.height)).toEqual([20, 80]);
    expect(snapshot.blocks[1]?.tableBreak?.cellBreaks).toBeUndefined();

    editor.destroy();
  });

  test('splits a nested table at its own row boundaries', () => {
    const editor = createNestedTablePaginationEditor();
    mockNestedTableGeometry(editor);

    const snapshot = measureDocumentLayoutBlocks(
      editor,
      null,
      0,
      new Map(),
      50,
    );

    expect(snapshot.blocks).toHaveLength(3);
    expect(snapshot.blocks.map(({ block }) => block.height)).toEqual([
      40, 40, 40,
    ]);
    expect(snapshot.blocks.map(({ block }) => block.flowIndex)).toEqual([
      0, 1, 2,
    ]);
    expect(snapshot.blocks[1]?.selectionRanges?.[0]).toMatchObject({
      from: expect.any(Number),
      to: expect.any(Number),
    });

    editor.destroy();
  });

  test('continues a row-spanning cell through each physical row', () => {
    const editor = createRowSpanningTableEditor();
    mockRowSpanningTableGeometry(editor);

    const snapshot = measureDocumentLayoutBlocks(
      editor,
      null,
      0,
      new Map(),
      50,
    );

    expect(snapshot.blocks).toHaveLength(3);
    expect(snapshot.blocks.map(({ block }) => block.height)).toEqual([
      40, 40, 40,
    ]);
    expect(
      snapshot.blocks.map((block) => block.selectionRanges?.length),
    ).toEqual([2, 2, 2]);
    expect(snapshot.blocks[1]?.tableBreak?.cellBreaks).toMatchObject([
      { cellIndex: 0, alignmentOffset: 0 },
    ]);
    expect(snapshot.blocks[2]?.tableBreak?.cellBreaks).toMatchObject([
      { cellIndex: 0, alignmentOffset: 0 },
    ]);
    const spanningRanges = snapshot.blocks.map(
      (block) => block.selectionRanges?.[0],
    );
    expect(spanningRanges[0]?.to).toBe(spanningRanges[1]?.from);
    expect(spanningRanges[1]?.to).toBe(spanningRanges[2]?.from);
    const continuation = snapshot.blocks[1];
    if (!continuation?.tableBreak?.cellBreaks) {
      throw new Error('Expected a row-span continuation break.');
    }
    editor.commands.applyDocumentPagination(1, [
      {
        beforeBlockId: continuation.block.id,
        pageIndex: 1,
        spacerHeight: 120,
        remainingBodyHeight: 20,
        page: {
          width: 300,
          height: 200,
          marginTop: 20,
          marginRight: 20,
          marginBottom: 20,
          marginLeft: 20,
          headerHeight: 10,
          footerHeight: 10,
          pageGap: 30,
        },
        position: continuation.from,
        inlineOffsetLeft: continuation.inlineOffsetLeft,
        inlineOffsetRight: continuation.inlineOffsetRight,
        tableBreak: continuation.tableBreak,
      },
    ]);
    const pageBreak = editor.view.dom.querySelector<HTMLElement>(
      '.work-document-table-cell-page-break[data-cell-index="0"]',
    );
    expect(pageBreak?.classList.contains('is-leading')).toBe(true);
    expect(pageBreak?.closest('td')?.getAttribute('rowspan')).toBe('3');

    editor.destroy();
  });

  test('renders an internal row break inside every cell boundary', () => {
    const editor = createSplitRowEditor(false);
    mockSplitRowGeometry(editor);
    const snapshot = measureDocumentLayoutBlocks(
      editor,
      null,
      0,
      new Map(),
      100,
    );
    const fragment = snapshot.blocks[2];
    if (!fragment?.tableBreak?.cellBreaks) {
      throw new Error('Expected a fragmented table row.');
    }
    editor.commands.applyDocumentPagination(1, [
      {
        beforeBlockId: fragment.block.id,
        pageIndex: 1,
        spacerHeight: 120,
        remainingBodyHeight: 30,
        page: {
          width: 300,
          height: 200,
          marginTop: 20,
          marginRight: 20,
          marginBottom: 20,
          marginLeft: 20,
          headerHeight: 10,
          footerHeight: 10,
          pageGap: 30,
        },
        position: fragment.from,
        inlineOffsetLeft: fragment.inlineOffsetLeft,
        inlineOffsetRight: fragment.inlineOffsetRight,
        tableBreak: fragment.tableBreak,
      },
    ]);

    const cellBreaks = Array.from(
      editor.view.dom.querySelectorAll<HTMLElement>(
        '.work-document-table-cell-page-break',
      ),
    );
    expect(cellBreaks).toHaveLength(2);
    expect(cellBreaks[0]?.classList.contains('is-leading')).toBe(true);
    expect(cellBreaks[0]?.dataset.cellIndex).toBe('0');
    expect(cellBreaks[1]?.dataset.cellIndex).toBe('1');
    expect(
      cellBreaks[0]?.querySelector('.work-document-table-cell-page-spacer'),
    ).not.toBeNull();
    expect(
      cellBreaks[0]?.querySelector('.work-document-table-cell-repeated-header')
        ?.textContent,
    ).toContain('标题');
    expect(
      cellBreaks.every(
        (cellBreak) =>
          cellBreak.querySelector('.work-document-table-cell-repeated-header')
            ?.tagName === 'DIV',
      ),
    ).toBe(true);
    expect(
      editor.view.dom.querySelectorAll(
        '.work-document-table-repeated-header-cell',
      ),
    ).toHaveLength(4);
    expect(
      cellBreaks[1]?.querySelector('.work-document-table-cell-page-spacer'),
    ).toBeNull();
    expect(
      editor.view.dom.querySelector('.work-document-table-page-break'),
    ).toBeNull();

    editor.destroy();
  });
});

function documentTextRange(
  editor: Editor,
  text: string,
): { from: number; to: number } {
  let range: { from: number; to: number } | null = null;
  editor.state.doc.descendants((node, position) => {
    if (range || !node.isText || !node.text) return;
    const offset = node.text.indexOf(text);
    if (offset < 0) return;
    range = {
      from: position + offset,
      to: position + offset + text.length,
    };
  });
  if (!range) throw new Error(`Text "${text}" was not found.`);
  return range;
}

function documentTableShape(editor: Editor): string[][] {
  let shape: string[][] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name !== 'table') return true;
    shape = Array.from({ length: node.childCount }, (_, rowIndex) => {
      const row = node.child(rowIndex);
      return Array.from(
        { length: row.childCount },
        (_, cellIndex) => row.child(cellIndex).type.name,
      );
    });
    return false;
  });
  return shape;
}

function tableRowTextPosition(editor: Editor, rowIndex: number): number {
  let current = 0;
  let position = 0;
  editor.state.doc.descendants((node, nodePosition) => {
    if (node.type.name !== 'tableRow') return true;
    if (current === rowIndex) position = nodePosition + 3;
    current += 1;
    return false;
  });
  if (!position) throw new Error(`Table row ${rowIndex} was not found.`);
  return position;
}

function createSplitRowEditor(cantSplit: boolean): Editor {
  return new Editor({
    extensions: [...createWorkDocumentExtensions(), DocumentPagination],
    content: [
      '<section data-document-section="true">',
      '<table><tbody>',
      '<tr data-office-repeat-header="true">',
      '<th><p>标题</p></th><th><p>负责人</p></th>',
      '</tr>',
      `<tr data-office-cant-split="${cantSplit}">`,
      '<td><p>第一段</p><p>第二段</p></td>',
      '<td><p>第一项</p><p>第二项</p></td>',
      '</tr>',
      '</tbody></table>',
      '</section>',
      '<p></p>',
    ].join(''),
  });
}

function createNestedTablePaginationEditor(): Editor {
  return new Editor({
    extensions: [...createWorkDocumentExtensions(), DocumentPagination],
    content: [
      '<section data-document-section="true">',
      '<table><tbody><tr><td>',
      '<table><tbody>',
      '<tr><td><p>Nested one</p></td></tr>',
      '<tr><td><p>Nested two</p></td></tr>',
      '<tr><td><p>Nested three</p></td></tr>',
      '</tbody></table>',
      '</td></tr></tbody></table>',
      '</section>',
    ].join(''),
  });
}

function createRowSpanningTableEditor(): Editor {
  return new Editor({
    extensions: [...createWorkDocumentExtensions(), DocumentPagination],
    content: [
      '<section data-document-section="true">',
      '<table><tbody>',
      '<tr><td rowspan="3"><p>Span one</p><p>Span two</p><p>Span three</p></td><td><p>Side one</p></td></tr>',
      '<tr><td><p>Side two</p></td></tr>',
      '<tr><td><p>Side three</p></td></tr>',
      '</tbody></table>',
      '</section>',
    ].join(''),
  });
}

function mockRowSpanningTableGeometry(editor: Editor): void {
  const wrapper = editor.view.dom.querySelector<HTMLElement>('.tableWrapper');
  const table = wrapper?.querySelector<HTMLElement>(':scope > table');
  const rows = Array.from(
    table?.querySelectorAll<HTMLElement>(':scope > tbody > tr') ?? [],
  );
  const spanningCell = rows[0]?.querySelector<HTMLElement>(
    ':scope > td[rowspan="3"]',
  );
  const sideCells = rows.map((row) =>
    row.querySelector<HTMLElement>(':scope > td:last-child'),
  );
  const spanningParagraphs = Array.from(
    spanningCell?.querySelectorAll<HTMLElement>(':scope > p') ?? [],
  );
  if (
    !wrapper ||
    !table ||
    rows.length !== 3 ||
    !spanningCell ||
    sideCells.some((cell) => !cell) ||
    spanningParagraphs.length !== 3
  ) {
    throw new Error('Expected a mounted row-spanning table.');
  }
  setElementBox(wrapper, { top: 0, left: 0, width: 300, height: 120 });
  setElementBox(table, { top: 0, left: 0, width: 300, height: 120 });
  rows.forEach((row, index) => {
    setElementBox(row, {
      top: index * 40,
      left: 0,
      width: 300,
      height: 40,
    });
    setElementBox(sideCells[index], {
      top: index * 40,
      left: 150,
      width: 150,
      height: 40,
    });
  });
  setElementBox(spanningCell, { top: 0, left: 0, width: 150, height: 120 });
  spanningParagraphs.forEach((paragraph, index) => {
    setElementBox(paragraph, {
      top: index * 40,
      left: 8,
      width: 134,
      height: 40,
    });
  });
}

function mockNestedTableGeometry(editor: Editor): void {
  const wrapper = editor.view.dom.querySelector<HTMLElement>('.tableWrapper');
  const outerTable = wrapper?.querySelector<HTMLElement>(':scope > table');
  const outerRow = outerTable?.querySelector<HTMLElement>(
    ':scope > tbody > tr',
  );
  const outerCell = outerRow?.querySelector<HTMLElement>(':scope > td');
  const nestedTable = outerCell?.querySelector<HTMLElement>(
    ':scope > .tableWrapper > table, :scope > table',
  );
  const nestedRows = Array.from(
    nestedTable?.querySelectorAll<HTMLElement>(':scope > tbody > tr') ?? [],
  );
  if (
    !wrapper ||
    !outerTable ||
    !outerRow ||
    !outerCell ||
    !nestedTable ||
    nestedRows.length !== 3
  ) {
    throw new Error('Expected a mounted nested table.');
  }
  setElementBox(wrapper, { top: 0, left: 0, width: 300, height: 120 });
  setElementBox(outerTable, { top: 0, left: 0, width: 300, height: 120 });
  setElementBox(outerRow, { top: 0, left: 0, width: 300, height: 120 });
  setElementBox(outerCell, { top: 0, left: 0, width: 300, height: 120 });
  setElementBox(nestedTable, { top: 0, left: 0, width: 300, height: 120 });
  nestedRows.forEach((row, index) => {
    setElementBox(row, {
      top: index * 40,
      left: 0,
      width: 300,
      height: 40,
    });
  });
}

function mockSplitRowGeometry(editor: Editor): void {
  const wrapper = editor.view.dom.querySelector<HTMLElement>('.tableWrapper');
  const table = wrapper?.querySelector<HTMLElement>(':scope > table');
  const rows = Array.from(
    table?.querySelectorAll<HTMLElement>(':scope > tbody > tr') ?? [],
  );
  const cells = Array.from(rows[1]?.children ?? []).filter(
    (element): element is HTMLElement => element instanceof HTMLElement,
  );
  if (!wrapper || !table || rows.length !== 2 || cells.length !== 2) {
    throw new Error('Expected the split-row table.');
  }
  setElementBox(wrapper, { top: 0, left: 0, width: 300, height: 100 });
  setElementBox(table, { top: 0, left: 0, width: 300, height: 100 });
  setElementBox(rows[0], { top: 0, left: 0, width: 300, height: 20 });
  setElementBox(rows[1], { top: 20, left: 0, width: 300, height: 80 });
  cells.forEach((cell, cellIndex) => {
    setElementBox(cell, {
      top: 20,
      left: cellIndex * 150,
      width: 150,
      height: 80,
    });
    const paragraphs = Array.from(
      cell.querySelectorAll<HTMLElement>(':scope > p'),
    );
    setElementBox(paragraphs[0], {
      top: 27,
      left: cellIndex * 150 + 8,
      width: 134,
      height: 25,
    });
    setElementBox(paragraphs[1], {
      top: 60,
      left: cellIndex * 150 + 8,
      width: 134,
      height: 30,
    });
  });
}

function setElementBox(
  element: HTMLElement | undefined,
  box: { top: number; left: number; width: number; height: number },
): void {
  if (!element) throw new Error('Expected a measured table element.');
  Object.defineProperties(element, {
    offsetHeight: { configurable: true, value: box.height },
    offsetWidth: { configurable: true, value: box.width },
  });
  element.getBoundingClientRect = () =>
    ({
      top: box.top,
      right: box.left + box.width,
      bottom: box.top + box.height,
      left: box.left,
      width: box.width,
      height: box.height,
      x: box.left,
      y: box.top,
      toJSON: () => ({}),
    }) as DOMRect;
}
