import { Editor } from '@tiptap/core';
import { afterEach, expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { DocumentTableInsertPopover } from '../src/internal/features/work/editors/document-table-insert-popover';
import {
  DocumentTableDesignRibbon,
  DocumentTableLayoutRibbon,
} from '../src/internal/features/work/editors/document-table-ribbon';
import { DocumentToolbar } from '../src/internal/features/work/editors/document-toolbar';
import { canInsertDocumentComment } from '../src/internal/features/work/work-document-comments';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

let editor: Editor | null = null;

afterEach(() => {
  const editorElement = editor?.view.dom;
  editor?.destroy();
  editorElement?.remove();
  editor = null;
});

test('chooses table dimensions with the keyboard and preserves selected text', async () => {
  editor = createPlainEditor();
  editor.commands.setTextSelection(documentTextRange(editor, 'selected text'));
  render(<DocumentTableInsertPopover editor={editor} />);

  const trigger = screen.getByRole('button', { name: '插入表格' });
  fireEvent.click(trigger);

  const firstCell = screen.getByRole('button', { name: '1 行 1 列' });
  await waitFor(() => expect(firstCell).toHaveFocus());
  fireEvent.keyDown(firstCell, { key: 'ArrowRight' });
  expect(screen.getByRole('button', { name: '1 行 2 列' })).toHaveFocus();
  fireEvent.keyDown(document.activeElement as HTMLElement, {
    key: 'ArrowDown',
  });
  expect(screen.getByRole('button', { name: '2 行 2 列' })).toHaveFocus();
  expect(screen.getByText('2 × 2 表格')).toBeInTheDocument();

  fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'Enter' });

  expect(screen.queryByRole('dialog', { name: '选择表格大小' })).toBeNull();
  expect(editor.getText()).toContain('Keep this selected text.');
  expect(tableShape(editor)).toEqual([2, 2]);
  await waitFor(() => expect(editor?.isFocused).toBe(true));
});

test('closes the table picker with Escape and restores trigger focus', async () => {
  editor = createPlainEditor();
  render(<DocumentTableInsertPopover editor={editor} />);

  const trigger = screen.getByRole('button', { name: '插入表格' });
  fireEvent.click(trigger);
  const firstCell = screen.getByRole('button', { name: '1 行 1 列' });
  await waitFor(() => expect(firstCell).toHaveFocus());

  fireEvent.keyDown(firstCell, { key: 'Escape' });

  expect(screen.queryByRole('dialog', { name: '选择表格大小' })).toBeNull();
  expect(trigger).toHaveFocus();
});

test('shows Word-style table Design and Layout tabs only inside a table', async () => {
  editor = createMixedEditor();
  const outsidePosition = documentTextRange(editor, 'Outside').from;
  editor.commands.setTextSelection(outsidePosition);
  const view = render(documentToolbar(editor));

  expect(screen.queryByRole('tab', { name: '表格设计' })).toBeNull();
  expect(screen.queryByRole('tab', { name: '表格布局' })).toBeNull();

  editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);
  view.rerender(documentToolbar(editor));

  const tableDesignTab = screen.getByRole('tab', { name: '表格设计' });
  const tableLayoutTab = screen.getByRole('tab', { name: '表格布局' });
  await waitFor(() =>
    expect(tableDesignTab).toHaveAttribute('aria-selected', 'true'),
  );
  fireEvent.click(tableLayoutTab);
  expect(
    screen.getByRole('button', { name: '在下方插入行' }),
  ).toBeInTheDocument();

  editor.commands.setTextSelection(outsidePosition);
  view.rerender(documentToolbar(editor));

  await waitFor(() =>
    expect(screen.queryByRole('tab', { name: '表格设计' })).toBeNull(),
  );
  expect(screen.queryByRole('tab', { name: '表格布局' })).toBeNull();
});

test('only enables adding a comment for an eligible text selection', () => {
  editor = createPlainEditor();
  const inserted: boolean[] = [];
  const view = render(
    documentToolbar(editor, {
      canInsertComment: canInsertDocumentComment(editor),
      onInsertComment: () => inserted.push(true),
    }),
  );

  fireEvent.click(screen.getByRole('tab', { name: '审阅' }));
  const addComment = screen.getByRole('button', { name: '添加批注' });
  expect(addComment).toBeDisabled();
  expect(addComment).toHaveAttribute('title', '请先选择未批注的文字');

  const range = documentTextRange(editor, 'selected text');
  editor.commands.setTextSelection(range);
  view.rerender(
    documentToolbar(editor, {
      canInsertComment: canInsertDocumentComment(editor),
      onInsertComment: () => inserted.push(true),
    }),
  );

  expect(addComment).not.toBeDisabled();
  expect(addComment).toHaveAttribute('title', '添加批注');
  fireEvent.click(addComment);
  expect(inserted).toEqual([true]);

  expect(
    editor.commands.insertDocumentComment({ id: 'existing-comment', range }),
  ).toBe(true);
  view.rerender(
    documentToolbar(editor, {
      canInsertComment: canInsertDocumentComment(editor),
      onInsertComment: () => inserted.push(true),
    }),
  );
  expect(addComment).toBeDisabled();
});

test('edits table rows, columns, header behavior, and removes the table', () => {
  editor = createTableEditor();
  editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);
  const view = render(<DocumentTableLayoutRibbon editor={editor} />);

  fireEvent.click(screen.getByRole('button', { name: '在下方插入行' }));
  expect(tableShape(editor)).toEqual([2, 2, 2]);
  fireEvent.click(screen.getByRole('button', { name: '在上方插入行' }));
  expect(tableShape(editor)).toEqual([2, 2, 2, 2]);
  fireEvent.click(screen.getByRole('button', { name: '删除当前行' }));
  expect(tableShape(editor)).toEqual([2, 2, 2]);

  fireEvent.click(screen.getByRole('button', { name: '在右侧插入列' }));
  expect(tableShape(editor)).toEqual([3, 3, 3]);
  fireEvent.click(screen.getByRole('button', { name: '在左侧插入列' }));
  expect(tableShape(editor)).toEqual([4, 4, 4]);
  fireEvent.click(screen.getByRole('button', { name: '删除当前列' }));
  expect(tableShape(editor)).toEqual([3, 3, 3]);

  editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);
  view.rerender(<DocumentTableDesignRibbon editor={editor} />);
  const originalFirstRow = firstRowCellTypes(editor);
  const toggledCellType =
    originalFirstRow[0] === 'tableHeader' ? 'tableCell' : 'tableHeader';
  fireEvent.click(screen.getByRole('button', { name: '标题行' }));
  expect(firstRowCellTypes(editor)).toEqual(Array(3).fill(toggledCellType));
  fireEvent.click(screen.getByRole('button', { name: '标题行' }));
  expect(firstRowCellTypes(editor)).toEqual(originalFirstRow);

  view.rerender(<DocumentTableLayoutRibbon editor={editor} />);
  const repeatHeaderBefore =
    editor.getAttributes('tableRow').repeatHeader ?? false;
  fireEvent.click(screen.getByRole('button', { name: '跨页重复标题' }));
  expect(editor.getAttributes('tableRow').repeatHeader).toBe(
    !repeatHeaderBefore,
  );
  const cantSplitBefore = editor.getAttributes('tableRow').cantSplit ?? false;
  fireEvent.click(screen.getByRole('button', { name: '整行不跨页' }));
  expect(editor.getAttributes('tableRow').cantSplit).toBe(!cantSplitBefore);

  fireEvent.click(screen.getByRole('button', { name: '删除表格' }));
  expect(tableShape(editor)).toEqual([]);
});

test('merges and splits selected table cells with command-aware controls', () => {
  editor = createTableEditor();
  const [firstCell, secondCell] = tableCellPositions(editor);
  editor.commands.setCellSelection({
    anchorCell: firstCell,
    headCell: secondCell,
  });
  const view = render(<DocumentTableLayoutRibbon editor={editor} />);

  expect(screen.getByRole('button', { name: '合并单元格' })).not.toBeDisabled();
  expect(screen.getByRole('button', { name: '拆分单元格' })).toBeDisabled();
  fireEvent.click(screen.getByRole('button', { name: '合并单元格' }));
  expect(tableShape(editor)).toEqual([1, 2]);
  expect(firstTableCellColspan(editor)).toBe(2);

  view.rerender(<DocumentTableLayoutRibbon editor={editor} />);
  expect(screen.getByRole('button', { name: '拆分单元格' })).not.toBeDisabled();
  fireEvent.click(screen.getByRole('button', { name: '拆分单元格' }));
  expect(tableShape(editor)).toEqual([2, 2]);
});

test('applies table styles, cell shading, and borders as coherent edits', () => {
  editor = createTableEditor();
  editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);
  let updateCount = 0;
  editor.on('update', () => {
    updateCount += 1;
  });
  const view = render(<DocumentTableDesignRibbon editor={editor} />);

  expect(
    screen.getByRole('radio', { name: '应用表格样式：网格' }),
  ).toBeChecked();
  fireEvent.click(
    screen.getByRole('radio', { name: '应用表格样式：蓝色条纹' }),
  );
  expect(updateCount).toBe(1);
  expect(tableCellAttributes(editor)[0]).toMatchObject({
    backgroundColor: '#d9eaf7',
    borderColor: '#9fbad0',
    borderStyle: 'solid',
    borderWidth: 1,
  });
  expect(tableCellAttributes(editor)[2]).toMatchObject({
    backgroundColor: '#f7fbff',
  });

  expect(editor.commands.undo()).toBe(true);
  expect(tableCellAttributes(editor)[0]).toMatchObject({
    backgroundColor: '#f1f4f9',
    borderColor: '#cfd5df',
  });

  const [firstCell, secondCell] = tableCellPositions(editor);
  editor.commands.setCellSelection({
    anchorCell: firstCell,
    headCell: secondCell,
  });
  view.rerender(<DocumentTableDesignRibbon editor={editor} />);
  fireEvent.click(screen.getByRole('button', { name: '单元格底纹' }));
  const fillDialog = screen.getByRole('dialog', { name: '单元格底纹' });
  fireEvent.click(
    within(fillDialog).getByRole('option', { name: '颜色 #fff2cc' }),
  );
  expect(tableCellAttributes(editor).slice(0, 2)).toEqual(
    expect.arrayContaining([
      expect.objectContaining({ backgroundColor: '#fff2cc' }),
      expect.objectContaining({ backgroundColor: '#fff2cc' }),
    ]),
  );

  fireEvent.click(screen.getByRole('combobox', { name: '边框样式' }));
  fireEvent.click(screen.getByRole('option', { name: '无边框' }));
  expect(tableCellAttributes(editor).slice(0, 2)).toEqual(
    expect.arrayContaining([
      expect.objectContaining({ borderStyle: 'none', borderWidth: 0 }),
      expect.objectContaining({ borderStyle: 'none', borderWidth: 0 }),
    ]),
  );
});

test('aligns table cells and applies Word-style sizing from Layout', () => {
  editor = createTableEditor();
  editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);
  expect(editor.commands.setCellAttribute('colwidth', [180])).toBe(true);
  const view = render(<DocumentTableLayoutRibbon editor={editor} />);

  fireEvent.click(screen.getByRole('button', { name: '单元格垂直居中' }));
  expect(tableCellAttributes(editor)[0]).toMatchObject({
    verticalAlign: 'middle',
  });

  fireEvent.click(screen.getByRole('button', { name: '单元格水平居中' }));
  expect(firstTableParagraphAlignment(editor)).toBe('center');

  view.rerender(<DocumentTableLayoutRibbon editor={editor} />);
  fireEvent.change(screen.getByRole('textbox', { name: '列宽（厘米）' }), {
    target: { value: '3.17' },
  });
  expect(Number(tableCellAttributes(editor)[0]?.colwidth?.[0])).toBeCloseTo(
    119.81,
    1,
  );

  fireEvent.click(screen.getByRole('button', { name: '平均分布列' }));
  const distributedWidths = tableCellAttributes(editor).map(
    ({ colwidth }) => colwidth?.[0],
  );
  expect(new Set(distributedWidths).size).toBe(1);

  fireEvent.click(screen.getByRole('button', { name: '平均分布行' }));
  expect(
    tableRowAttributes(editor).every(({ rowHeight }) => rowHeight === 36),
  ).toBe(true);

  view.rerender(<DocumentTableLayoutRibbon editor={editor} />);
  fireEvent.click(screen.getByRole('combobox', { name: '表格自动调整' }));
  fireEvent.click(screen.getByRole('option', { name: '适应内容' }));
  expect(firstTableAttributes(editor)).toMatchObject({
    layoutMode: 'contents',
  });
  expect(tableCellAttributes(editor).every(({ colwidth }) => !colwidth)).toBe(
    true,
  );
});

function createPlainEditor(): Editor {
  const currentEditor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: [
      '<section data-document-section="true">',
      '<p>Keep this selected text.</p>',
      '<p>Continue here.</p>',
      '</section>',
    ].join(''),
  });
  document.body.append(currentEditor.view.dom);
  return currentEditor;
}

function createTableEditor(): Editor {
  return new Editor({
    extensions: createWorkDocumentExtensions(),
    content: [
      '<section data-document-section="true">',
      '<table><tbody>',
      '<tr data-office-repeat-header="true">',
      '<th><p>Title A</p></th><th><p>Title B</p></th>',
      '</tr>',
      '<tr><td><p>Value A</p></td><td><p>Value B</p></td></tr>',
      '</tbody></table>',
      '</section>',
    ].join(''),
  });
}

function createMixedEditor(): Editor {
  return new Editor({
    extensions: createWorkDocumentExtensions(),
    content: [
      '<section data-document-section="true">',
      '<p>Outside</p>',
      '<table><tbody>',
      '<tr><th><p>Title A</p></th><th><p>Title B</p></th></tr>',
      '<tr><td><p>Value A</p></td><td><p>Value B</p></td></tr>',
      '</tbody></table>',
      '</section>',
    ].join(''),
  });
}

function firstTableAttributes(editor: Editor): Record<string, unknown> {
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

function documentToolbar(
  currentEditor: Editor,
  {
    canInsertComment = false,
    onInsertComment = () => undefined,
  }: {
    canInsertComment?: boolean;
    onInsertComment?: () => void;
  } = {},
) {
  const noop = () => undefined;
  return (
    <DocumentToolbar
      editor={currentEditor}
      layoutOpen={false}
      navigationOpen={false}
      pageColor="#ffffff"
      showPageNumbers
      showRulers={false}
      spellcheckEnabled
      viewMode="page"
      zoom={100}
      pageChromeEditor={null}
      pageChromeEditingPart={null}
      pageChromeShowPageNumber
      onRequestImage={noop}
      onPageChromeEditingPartChange={noop}
      onClosePageChrome={noop}
      onTogglePageChromePageNumber={noop}
      onToggleLayout={noop}
      onToggleNavigation={noop}
      onTogglePageNumbers={noop}
      onToggleRulers={noop}
      onPageColorChange={noop}
      onToggleSpellcheck={noop}
      onViewModeChange={noop}
      onZoomChange={noop}
      onInsertSection={noop}
      onInsertNote={noop}
      onInsertCaption={noop}
      onInsertCrossReference={noop}
      citationsOpen={false}
      citationSourceCount={0}
      onToggleCitations={noop}
      onInsertField={noop}
      onRefreshFields={noop}
      canInsertComment={canInsertComment}
      onInsertComment={onInsertComment}
      commentsOpen={false}
      commentCount={0}
      onToggleComments={noop}
      trackChanges={false}
      changesOpen={false}
      changeCount={0}
      findReplaceMode={null}
      onToggleTrackChanges={noop}
      onToggleChanges={noop}
      onOpenFindReplace={noop}
    />
  );
}

function documentTextRange(
  currentEditor: Editor,
  text: string,
): { from: number; to: number } {
  let range: { from: number; to: number } | null = null;
  currentEditor.state.doc.descendants((node, position) => {
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

function tableCellPositions(currentEditor: Editor): number[] {
  const positions: number[] = [];
  currentEditor.state.doc.descendants((node, position) => {
    if (node.type.name === 'tableCell' || node.type.name === 'tableHeader') {
      positions.push(position);
      return false;
    }
    return true;
  });
  return positions;
}

function tableShape(currentEditor: Editor): number[] {
  let shape: number[] = [];
  currentEditor.state.doc.descendants((node) => {
    if (node.type.name !== 'table') return true;
    shape = Array.from(
      { length: node.childCount },
      (_, rowIndex) => node.child(rowIndex).childCount,
    );
    return false;
  });
  return shape;
}

function firstRowCellTypes(currentEditor: Editor): string[] {
  let types: string[] = [];
  currentEditor.state.doc.descendants((node) => {
    if (node.type.name !== 'table') return true;
    const row = node.firstChild;
    if (row) {
      types = Array.from(
        { length: row.childCount },
        (_, index) => row.child(index).type.name,
      );
    }
    return false;
  });
  return types;
}

function firstTableCellColspan(currentEditor: Editor): number | null {
  let colspan: number | null = null;
  currentEditor.state.doc.descendants((node) => {
    if (
      colspan === null &&
      (node.type.name === 'tableCell' || node.type.name === 'tableHeader')
    ) {
      colspan = Number(node.attrs.colspan);
      return false;
    }
    return true;
  });
  return colspan;
}

function tableCellAttributes(currentEditor: Editor): Record<string, unknown>[] {
  const attributes: Record<string, unknown>[] = [];
  currentEditor.state.doc.descendants((node) => {
    if (node.type.name === 'tableCell' || node.type.name === 'tableHeader') {
      attributes.push(node.attrs);
      return false;
    }
    return true;
  });
  return attributes;
}

function firstTableParagraphAlignment(currentEditor: Editor): unknown {
  let alignment: unknown;
  currentEditor.state.doc.descendants((node) => {
    if (alignment !== undefined || node.type.name !== 'paragraph') return true;
    alignment = node.attrs.textAlign;
    return false;
  });
  return alignment;
}
