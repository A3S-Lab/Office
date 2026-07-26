import { Editor } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { documentTableSizing } from '../src/internal/features/work/work-document-table-sizing';

describe('document table sizing', () => {
  test('keeps table layout, physical column widths, and row heights in HTML', () => {
    const editor = createSizingEditor();
    editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);

    expect(documentTableSizing(editor.state)).toMatchObject({
      columnWidth: 120,
      rowHeight: 38,
      rowHeightRule: 'exact',
      layoutMode: 'fixed',
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
    expect(tableAttributes(editor)).toMatchObject({ layoutMode: 'fixed' });

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
    expect(tableAttributes(editor)).toMatchObject({ layoutMode: 'contents' });
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
    expect(tableAttributes(editor)).toMatchObject({ layoutMode: 'fixed' });

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
    expect(tableAttributes(editor)).toMatchObject({ layoutMode: 'fixed' });

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

function tableAttributes(editor: Editor): Record<string, unknown> {
  let attributes: Record<string, unknown> = {};
  editor.state.doc.descendants((node) => {
    if (node.type.name !== 'table') return true;
    attributes = node.attrs;
    return false;
  });
  return attributes;
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
