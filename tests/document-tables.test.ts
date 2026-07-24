import { Editor } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  applyDocumentPagination,
  DocumentPagination,
  measureDocumentLayoutBlocks,
} from '../src/internal/features/work/work-document-pagination';
import {
  canSetDocumentTableRowRepeatHeader,
  documentTableRowOptions,
  setDocumentTableRowOptions,
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
      setDocumentTableRowOptions(
        editor,
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
    applyDocumentPagination(editor, 1, [
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
