import { Editor } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { documentTableSizing } from '../src/internal/features/work/work-document-table-sizing';

describe('document table sizing', () => {
  test('sizes the complete table from a whole-table node selection', () => {
    const editor = createSizingEditor();
    editor.commands.setNodeSelection(firstTablePosition(editor));

    expect(documentTableSizing(editor.state)).toMatchObject({
      selectedColumnCount: 2,
      selectedRowCount: 3,
    });
    expect(editor.commands.setDocumentTableColumnWidth(150)).toBe(true);
    expect(tableColumnWidths(editor)).toEqual([
      [150, 150],
      [150, 150],
      [150, 150],
    ]);
    expect(editor.state.selection.toJSON()).toMatchObject({ type: 'node' });
    expect(documentTableSizing(editor.state)).toMatchObject({
      selectedRowCount: 3,
    });
    expect(editor.commands.setDocumentTableRowHeight(42, 'exact')).toBe(true);
    expect(tableRowAttributes(editor)).toEqual(
      Array.from({ length: 3 }, () =>
        expect.objectContaining({ rowHeight: 42, rowHeightRule: 'exact' }),
      ),
    );

    editor.destroy();
  });

  test('keeps table layout, physical column widths, and row heights in HTML', () => {
    const editor = createSizingEditor();
    editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);

    expect(documentTableSizing(editor.state)).toMatchObject({
      columnWidth: 120,
      rowHeight: 38,
      rowHeightRule: 'exact',
      layoutMode: 'fixed',
      layoutAlgorithm: 'fixed',
      preferredWidthType: 'pixels',
      preferredWidth: 300,
      alignment: 'left',
      indent: 0,
    });
    expect(editor.getHTML()).toContain('data-office-table-layout="fixed"');
    expect(editor.getHTML()).toContain('colwidth="120"');
    expect(editor.getHTML()).toContain('data-office-row-height="38"');
    expect(editor.getHTML()).toContain('data-office-row-height-rule="exact"');

    editor.destroy();
  });

  test('sets the selected physical column and current row without rebuilding the table', () => {
    const editor = createSizingEditor();
    editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);
    let updateCount = 0;
    editor.on('update', () => {
      updateCount += 1;
    });

    expect(editor.commands.setDocumentTableColumnWidth(160)).toBe(true);
    expect(tableColumnWidths(editor)).toEqual([
      [160, 180],
      [160, 180],
      [160, 180],
    ]);
    expect(documentTableSizing(editor.state)).toMatchObject({
      layoutMode: 'fixed',
      layoutAlgorithm: 'fixed',
      preferredWidthType: 'pixels',
      preferredWidth: 340,
    });

    expect(editor.commands.setDocumentTableRowHeight(44, 'atLeast')).toBe(true);
    expect(
      tableRowAttributes(editor).map(({ rowHeight }) => rowHeight),
    ).toEqual([44, null, null]);
    expect(tableRowAttributes(editor)[0]).toMatchObject({
      rowHeightRule: 'atLeast',
    });
    expect(updateCount).toBe(2);

    editor.destroy();
  });

  test('distributes real column and row dimensions and clears widths for autofit', () => {
    const editor = createSizingEditor();
    editor.commands.setTextSelection(tableCellPositions(editor)[2] + 2);

    expect(editor.commands.distributeDocumentTableColumns(480)).toBe(true);
    expect(tableColumnWidths(editor)).toEqual([
      [240, 240],
      [240, 240],
      [240, 240],
    ]);

    expect(editor.commands.distributeDocumentTableRows(108)).toBe(true);
    expect(tableRowAttributes(editor)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          rowHeight: 36,
          rowHeightRule: 'atLeast',
        }),
      ]),
    );
    expect(
      tableRowAttributes(editor).every(({ rowHeight }) => rowHeight === 36),
    ).toBe(true);

    expect(editor.commands.setDocumentTableLayoutMode('contents')).toBe(true);
    expect(documentTableSizing(editor.state)).toMatchObject({
      layoutMode: 'contents',
      layoutAlgorithm: 'autofit',
      preferredWidthType: 'auto',
      preferredWidth: null,
    });
    expect(
      tableColumnWidths(editor).every((widths) =>
        widths.every((width) => !width),
      ),
    ).toBe(true);
    expect(
      editor.view.dom.querySelector<HTMLElement>('colgroup > col')?.style.width,
    ).toBe('');

    expect(editor.commands.setDocumentTableLayoutMode('fixed', 420)).toBe(true);
    expect(tableColumnWidths(editor)).toEqual([
      [210, 210],
      [210, 210],
      [210, 210],
    ]);
    expect(documentTableSizing(editor.state)).toMatchObject({
      layoutMode: 'fixed',
      layoutAlgorithm: 'fixed',
      preferredWidthType: 'pixels',
      preferredWidth: 420,
    });

    expect(editor.commands.setDocumentTableAlignment('center')).toBe(true);
    expect(
      editor.commands.setDocumentTableCellMargins({
        top: 4,
        right: 10,
        bottom: 6,
        left: 12,
      }),
    ).toBe(true);
    expect(documentTableSizing(editor.state)).toMatchObject({
      alignment: 'center',
      cellMargins: { top: 4, right: 10, bottom: 6, left: 12 },
    });
    expect(editor.getHTML()).toContain('data-office-table-alignment="center"');
    expect(editor.getHTML()).toContain(
      'data-office-table-cell-margin-left="12"',
    );

    editor.destroy();
  });

  test('updates preferred width, alignment, and indent atomically', () => {
    const editor = createSizingEditor();
    editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);
    let updateCount = 0;
    editor.on('update', () => {
      updateCount += 1;
    });

    expect(
      editor.commands.setDocumentTableProperties({
        width: { type: 'percent', value: 62.5 },
        alignment: 'left',
        indent: 18.9,
      }),
    ).toBe(true);
    expect(documentTableSizing(editor.state)).toMatchObject({
      preferredWidthType: 'percent',
      preferredWidth: 62.5,
      alignment: 'left',
      indent: 18.9,
    });
    expect(editor.getHTML()).toContain('data-office-table-width="62.5"');
    expect(editor.getHTML()).toContain('data-office-table-indent="18.9"');
    expect(updateCount).toBe(1);

    expect(
      editor.commands.setDocumentTableProperties({
        width: { type: 'auto', value: null },
        alignment: 'center',
        indent: 18.9,
      }),
    ).toBe(true);
    expect(documentTableSizing(editor.state)).toMatchObject({
      preferredWidthType: 'auto',
      preferredWidth: null,
      alignment: 'center',
      indent: 18.9,
    });
    expect(updateCount).toBe(2);

    expect(
      editor.commands.setDocumentTableProperties({
        width: { type: 'percent', value: 0 },
        alignment: 'left',
        indent: 0,
      }),
    ).toBe(false);
    expect(updateCount).toBe(2);

    editor.destroy();
  });

  test('commits table, row, column, and cell properties as one undoable update', () => {
    const editor = createSizingEditor();
    editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);
    const originalHtml = editor.getHTML();
    let updateCount = 0;
    editor.on('update', () => {
      updateCount += 1;
    });

    expect(
      editor.commands.setDocumentTablePropertyChanges({
        table: {
          width: { type: 'percent', value: 72 },
          alignment: 'left',
          indent: 18.9,
        },
        row: {
          height: 45,
          heightRule: 'exact',
          cantSplit: true,
          repeatHeader: false,
        },
        column: {
          width: 150,
          renderedColumnWidths: [120, 180],
        },
        cell: {
          verticalAlign: 'middle',
          margins: { top: 4, right: 8, bottom: 6, left: 10 },
        },
      }),
    ).toBe(true);

    expect(updateCount).toBe(1);
    expect(documentTableSizing(editor.state)).toMatchObject({
      layoutAlgorithm: 'fixed',
      preferredWidthType: 'percent',
      preferredWidth: 72,
      alignment: 'left',
      indent: 18.9,
      columnWidth: 150,
      rowHeight: 45,
      rowHeightRule: 'exact',
    });
    expect(tableColumnWidths(editor)).toEqual([
      [150, 180],
      [150, 180],
      [150, 180],
    ]);
    expect(tableRowAttributes(editor)[0]).toMatchObject({
      rowHeight: 45,
      rowHeightRule: 'exact',
      cantSplit: true,
      repeatHeader: false,
    });
    expect(tableCellAttributes(editor)[0]).toMatchObject({
      verticalAlign: 'middle',
      margins: { top: 4, right: 8, bottom: 6, left: 10 },
    });

    expect(editor.commands.undo()).toBe(true);
    expect(editor.getHTML()).toBe(originalHtml);

    editor.destroy();
  });

  test('keeps an inserted table when undoing its property dialog update', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content:
        '<section data-document-section="true"><p>Before table</p></section>',
    });
    editor.commands.setTextSelection(3);

    expect(
      editor.commands.insertTable({ rows: 2, cols: 2, withHeaderRow: true }),
    ).toBe(true);
    const insertedHtml = editor.getHTML();

    expect(
      editor.commands.setDocumentTablePropertyChanges({
        table: {
          width: { type: 'percent', value: 80 },
          alignment: 'center',
          indent: 0,
        },
      }),
    ).toBe(true);
    expect(editor.getHTML()).not.toBe(insertedHtml);

    expect(editor.commands.undo()).toBe(true);
    expect(editor.getHTML()).toBe(insertedHtml);

    editor.destroy();
  });

  test('updates physical widths coherently through merged cells', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true">',
        '<table data-office-table-layout="fixed"><tbody>',
        '<tr><td colspan="2" colwidth="100,140"><p>Merged</p></td></tr>',
        '<tr><td colwidth="100"><p>A</p></td><td colwidth="140"><p>B</p></td></tr>',
        '</tbody></table>',
        '</section>',
      ].join(''),
    });
    editor.commands.setTextSelection(tableCellPositions(editor)[1] + 2);

    expect(editor.commands.setDocumentTableColumnWidth(155)).toBe(true);
    expect(tableColumnWidths(editor)).toEqual([
      [155, 140],
      [155, 140],
    ]);

    editor.destroy();
  });

  test('treats a drag-created column width as fixed layout', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true">',
        '<table><tbody><tr>',
        '<td><p>A</p></td><td><p>B</p></td>',
        '</tr></tbody></table>',
        '</section>',
      ].join(''),
    });
    editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);

    expect(editor.commands.setCellAttribute('colwidth', [175])).toBe(true);
    expect(documentTableSizing(editor.state)).toMatchObject({
      layoutMode: 'fixed',
      layoutAlgorithm: 'fixed',
    });

    editor.destroy();
  });

  test('preserves imported autofit geometry when a table is inserted', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content:
        '<section data-document-section="true"><p>Before table</p></section>',
    });

    expect(
      editor.commands.setContent(
        [
          '<section data-document-section="true">',
          '<table data-office-table-layout="window"><tbody><tr>',
          '<td colwidth="135"><p>A</p></td>',
          '<td colwidth="165"><p>B</p></td>',
          '</tr></tbody></table>',
          '</section>',
        ].join(''),
      ),
    ).toBe(true);
    editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);

    expect(documentTableSizing(editor.state)).toMatchObject({
      layoutMode: 'window',
      layoutAlgorithm: 'autofit',
      preferredWidthType: 'percent',
      preferredWidth: 100,
    });

    editor.destroy();
  });

  test('preserves remaining table geometry when an earlier table is removed', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true">',
        '<table data-office-table-layout="fixed"><tbody><tr>',
        '<td colwidth="90"><p>First</p></td>',
        '</tr></tbody></table>',
        '<table data-office-table-layout="window"><tbody><tr>',
        '<td colwidth="135"><p>A</p></td>',
        '<td colwidth="165"><p>B</p></td>',
        '</tr></tbody></table>',
        '</section>',
      ].join(''),
    });

    expect(
      editor.commands.setContent(
        [
          '<section data-document-section="true">',
          '<table data-office-table-layout="window"><tbody><tr>',
          '<td colwidth="135"><p>A</p></td>',
          '<td colwidth="165"><p>B</p></td>',
          '</tr></tbody></table>',
          '</section>',
        ].join(''),
      ),
    ).toBe(true);
    editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);

    expect(documentTableSizing(editor.state)).toMatchObject({
      layoutMode: 'window',
      layoutAlgorithm: 'autofit',
      preferredWidthType: 'percent',
      preferredWidth: 100,
    });

    editor.destroy();
  });

  test('preserves rendered sibling widths when sizing an autofit column', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true">',
        '<table><tbody><tr>',
        '<td><p>A</p></td><td><p>B</p></td><td><p>C</p></td>',
        '</tr></tbody></table>',
        '</section>',
      ].join(''),
    });
    editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);

    expect(
      editor.commands.setDocumentTableColumnWidth(125, [190, 190, 190]),
    ).toBe(true);
    expect(tableColumnWidths(editor)).toEqual([[125, 190, 190]]);
    expect(documentTableSizing(editor.state)).toMatchObject({
      layoutMode: 'fixed',
      layoutAlgorithm: 'fixed',
      preferredWidthType: 'pixels',
      preferredWidth: 505,
    });

    editor.destroy();
  });
});

function createSizingEditor(): Editor {
  return new Editor({
    extensions: createWorkDocumentExtensions(),
    content: [
      '<section data-document-section="true">',
      '<table data-office-table-layout="fixed"><tbody>',
      '<tr data-office-row-height="38" data-office-row-height-rule="exact">',
      '<th colwidth="120"><p>Title A</p></th>',
      '<th colwidth="180"><p>Title B</p></th>',
      '</tr>',
      '<tr><td colwidth="120"><p>Value A</p></td>',
      '<td colwidth="180"><p>Value B</p></td></tr>',
      '<tr><td colwidth="120"><p>Total A</p></td>',
      '<td colwidth="180"><p>Total B</p></td></tr>',
      '</tbody></table>',
      '</section>',
    ].join(''),
  });
}

function tableRowAttributes(editor: Editor): Record<string, unknown>[] {
  const attributes: Record<string, unknown>[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name !== 'tableRow') return true;
    attributes.push(node.attrs);
    return false;
  });
  return attributes;
}

function tableColumnWidths(editor: Editor): Array<Array<number | null>> {
  const rows: Array<Array<number | null>> = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name !== 'tableRow') return true;
    const widths: Array<number | null> = [];
    node.forEach((cell) => {
      const colwidth = Array.isArray(cell.attrs.colwidth)
        ? (cell.attrs.colwidth as number[])
        : [];
      for (let index = 0; index < Number(cell.attrs.colspan ?? 1); index += 1) {
        widths.push(colwidth[index] ?? null);
      }
    });
    rows.push(widths);
    return false;
  });
  return rows;
}

function tableCellAttributes(editor: Editor): Record<string, unknown>[] {
  const attributes: Record<string, unknown>[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name !== 'tableCell' && node.type.name !== 'tableHeader') {
      return true;
    }
    attributes.push(node.attrs);
    return false;
  });
  return attributes;
}

function tableCellPositions(editor: Editor): number[] {
  const positions: number[] = [];
  editor.state.doc.descendants((node, position) => {
    if (node.type.name === 'tableCell' || node.type.name === 'tableHeader') {
      positions.push(position);
      return false;
    }
    return true;
  });
  return positions;
}

function firstTablePosition(editor: Editor): number {
  let position: number | null = null;
  editor.state.doc.descendants((node, offset) => {
    if (position === null && node.type.name === 'table') position = offset;
    return position === null;
  });
  if (position === null) throw new Error('Expected a table node.');
  return position;
}
