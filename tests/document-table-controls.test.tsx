import { afterEach, expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { Editor } from '@tiptap/core';
import { DocumentTableInsertPopover } from '../src/internal/features/work/editors/document-table-insert-popover';
import {
  DocumentTableDesignRibbon,
  DocumentTableLayoutRibbon,
} from '../src/internal/features/work/editors/document-table-ribbon';
import { DocumentToolbar } from '../src/internal/features/work/editors/document-toolbar';
import { OfficeTableInsertPopover } from '../src/internal/features/work/editors/office-table-insert-popover';
import { canInsertDocumentComment } from '../src/internal/features/work/work-document-comments';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { documentTableSizing } from '../src/internal/features/work/work-document-table-sizing';

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

test('shares the table-size picker with presentation commands', async () => {
  const dimensions: Array<{ rows: number; columns: number }> = [];
  render(
    <OfficeTableInsertPopover
      label="表格"
      onInsert={(value) => dimensions.push(value)}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '表格' }));
  const target = screen.getByRole('button', { name: '4 行 5 列' });
  fireEvent.mouseEnter(target);
  expect(screen.getByText('4 × 5 表格')).toBeInTheDocument();
  fireEvent.click(target);

  expect(dimensions).toEqual([{ rows: 4, columns: 5 }]);
  expect(screen.queryByRole('dialog', { name: '选择表格大小' })).toBeNull();
});

test('exposes one desktop control per table size without duplicate cells', async () => {
  render(<OfficeTableInsertPopover label="表格" onInsert={() => undefined} />);

  fireEvent.click(screen.getByRole('button', { name: '表格' }));
  await waitFor(() =>
    expect(screen.getByRole('button', { name: '1 行 1 列' })).toHaveFocus(),
  );

  expect(screen.queryAllByRole('cell')).toHaveLength(0);
  expect(screen.getAllByRole('button', { name: /\d+ 行 \d+ 列/ })).toHaveLength(
    80,
  );
});

test('uses touch-sized row and column controls on compact screens', async () => {
  const restoreMatchMedia = mockMatchMedia(true);
  const dimensions: Array<{ rows: number; columns: number }> = [];

  try {
    render(
      <OfficeTableInsertPopover
        label="表格"
        onInsert={(value) => dimensions.push(value)}
      />,
    );

    fireEvent.click(screen.getByRole('button', { name: '表格' }));
    const rows = screen.getByRole('spinbutton', { name: '行数' });
    const columns = screen.getByRole('spinbutton', { name: '列数' });
    await waitFor(() => expect(rows).toHaveFocus());
    expect(screen.queryByRole('button', { name: '1 行 1 列' })).toBeNull();

    fireEvent.change(rows, { target: { value: '3' } });
    fireEvent.click(screen.getByRole('button', { name: '增加列数' }));
    fireEvent.click(screen.getByRole('button', { name: '增加列数' }));
    expect(columns).toHaveValue(3);
    fireEvent.click(screen.getByRole('button', { name: '插入 3 × 3 表格' }));

    expect(dimensions).toEqual([{ rows: 3, columns: 3 }]);
    expect(screen.queryByRole('dialog', { name: '选择表格大小' })).toBeNull();
  } finally {
    restoreMatchMedia();
  }
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
  expect(tableDesignTab).toHaveAttribute('data-contextual', 'true');
  expect(tableLayoutTab).toHaveAttribute('data-contextual', 'true');
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
  expect(addComment).toHaveAttribute('title', '添加批注（Cmd/Ctrl+Alt+M）');
  expect(addComment).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Alt+M Meta+Alt+M',
  );
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

test('applies table styles and cell shading as coherent edits', () => {
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
});

test('applies per-edge borders with a reusable Word-style border pen', () => {
  editor = createTableEditor();
  editor.commands.setNodeSelection(firstTablePosition(editor));
  render(<DocumentTableDesignRibbon editor={editor} />);

  const applyBorders = screen.getByRole('combobox', {
    name: '应用边框',
  });
  fireEvent.click(applyBorders);
  fireEvent.click(screen.getByRole('option', { name: '无框线' }));
  expect(
    tableCellAttributes(editor).every(({ borders }) =>
      Object.values(borders ?? {}).every(
        (border) => border.style === 'none' && border.width === 0,
      ),
    ),
  ).toBe(true);

  fireEvent.click(screen.getByRole('combobox', { name: '边框样式' }));
  fireEvent.click(screen.getByRole('option', { name: '双线' }));
  expect(
    tableCellAttributes(editor).every(({ borders }) =>
      Object.values(borders ?? {}).every((border) => border.style === 'none'),
    ),
  ).toBe(true);

  fireEvent.click(applyBorders);
  fireEvent.click(screen.getByRole('option', { name: '外侧框线' }));
  const [topLeft, topRight, bottomLeft, bottomRight] =
    tableCellAttributes(editor);
  expect(topLeft?.borders).toMatchObject({
    top: { style: 'double', width: 2 },
    right: { style: 'none', width: 0 },
    bottom: { style: 'none', width: 0 },
    left: { style: 'double', width: 2 },
  });
  expect(topRight?.borders).toMatchObject({
    top: { style: 'double', width: 2 },
    right: { style: 'double', width: 2 },
  });
  expect(bottomLeft?.borders).toMatchObject({
    bottom: { style: 'double', width: 2 },
    left: { style: 'double', width: 2 },
  });
  expect(bottomRight?.borders).toMatchObject({
    right: { style: 'double', width: 2 },
    bottom: { style: 'double', width: 2 },
  });

  fireEvent.click(screen.getByRole('combobox', { name: '边框样式' }));
  fireEvent.click(screen.getByRole('option', { name: '虚线' }));
  fireEvent.click(applyBorders);
  fireEvent.click(screen.getByRole('option', { name: '内部横框线' }));
  expect(tableCellAttributes(editor)[0]?.borders).toMatchObject({
    top: { style: 'double' },
    bottom: { style: 'dashed', width: 1 },
  });
  expect(tableCellAttributes(editor)[2]?.borders).toMatchObject({
    top: { style: 'dashed', width: 1 },
    bottom: { style: 'double' },
  });
});

test('applies one table style per radio-gallery keyboard move', () => {
  editor = createTableEditor();
  editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);
  let updateCount = 0;
  editor.on('update', () => {
    updateCount += 1;
  });
  render(<DocumentTableDesignRibbon editor={editor} />);

  const grid = screen.getByRole('radio', {
    name: '应用表格样式：网格',
  });
  const blue = screen.getByRole('radio', {
    name: '应用表格样式：蓝色条纹',
  });
  const clean = screen.getByRole('radio', {
    name: '应用表格样式：简洁',
  });

  grid.focus();
  fireEvent.keyDown(grid, { key: 'ArrowRight' });
  expect(blue).toHaveFocus();
  expect(tableCellAttributes(editor)[0]).toMatchObject({
    backgroundColor: '#d9eaf7',
    borderStyle: 'solid',
  });

  fireEvent.keyDown(blue, { key: 'End' });
  expect(clean).toHaveFocus();
  expect(tableCellAttributes(editor)[0]).toMatchObject({
    borderStyle: 'none',
    borderWidth: 0,
  });

  fireEvent.keyDown(clean, { key: 'Home' });
  expect(grid).toHaveFocus();
  expect(tableCellAttributes(editor)[0]).toMatchObject({
    backgroundColor: '#f1f4f9',
    borderStyle: 'solid',
  });
  expect(updateCount).toBe(3);
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
  fireEvent.blur(screen.getByRole('textbox', { name: '列宽（厘米）' }));
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
  fireEvent.click(screen.getByRole('button', { name: '表格属性' }));
  fireEvent.click(screen.getByRole('radio', { name: '居中' }));
  fireEvent.click(screen.getByRole('button', { name: '确定' }));
  expect(firstTableAttributes(editor)).toMatchObject({
    geometry: expect.objectContaining({ alignment: 'center' }),
  });

  view.rerender(<DocumentTableLayoutRibbon editor={editor} />);
  fireEvent.click(screen.getByRole('button', { name: '单元格边距' }));
  const leftMargin = screen.getByRole('textbox', {
    name: '单元格左边距（厘米）',
  });
  fireEvent.change(leftMargin, { target: { value: '0.5' } });
  fireEvent.blur(leftMargin);
  expect(
    Number(
      (
        firstTableAttributes(editor).geometry as {
          cellMargins?: { left?: number };
        }
      ).cellMargins?.left,
    ),
  ).toBeCloseTo(18.9, 1);

  view.rerender(<DocumentTableLayoutRibbon editor={editor} />);
  fireEvent.click(screen.getByRole('combobox', { name: '表格自动调整' }));
  fireEvent.click(screen.getByRole('option', { name: '适应内容' }));
  expect(firstTableAttributes(editor)).toMatchObject({
    geometry: expect.objectContaining({
      layout: 'autofit',
      width: { type: 'auto', value: null },
    }),
  });
  expect(tableCellAttributes(editor).every(({ colwidth }) => !colwidth)).toBe(
    true,
  );
});

test('edits whole-table width and position with validation and cancel safety', async () => {
  editor = createTableEditor();
  editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);
  render(<DocumentTableLayoutRibbon editor={editor} />);

  const trigger = screen.getByRole('button', { name: '表格属性' });
  trigger.focus();
  fireEvent.click(trigger);

  expect(screen.getByRole('dialog', { name: '表格属性' })).toBeInTheDocument();
  const percentageWidth = screen.getByRole('radio', { name: '百分比' });
  expect(percentageWidth).toBeChecked();
  await waitFor(() => expect(percentageWidth).toHaveFocus());
  expect(
    screen.getByRole('textbox', { name: '表格宽度（百分比）' }),
  ).toHaveValue('100');

  fireEvent.change(
    screen.getByRole('textbox', { name: '表格宽度（百分比）' }),
    { target: { value: '0' } },
  );
  expect(screen.getByRole('button', { name: '确定' })).toBeDisabled();
  expect(screen.getByRole('alert')).toHaveTextContent('请输入 1 到 100');

  fireEvent.change(
    screen.getByRole('textbox', { name: '表格宽度（百分比）' }),
    { target: { value: '62.5' } },
  );
  fireEvent.click(screen.getByRole('radio', { name: '左对齐' }));
  const indent = screen.getByRole('textbox', {
    name: '表格左缩进（厘米）',
  });
  fireEvent.change(indent, { target: { value: '' } });
  expect(screen.getByRole('button', { name: '确定' })).toBeDisabled();
  expect(screen.getByRole('alert')).toHaveTextContent('请输入 0 到 30');

  fireEvent.click(screen.getByRole('radio', { name: '居中' }));
  expect(indent).toBeDisabled();
  expect(screen.queryByRole('alert')).toBeNull();
  expect(screen.getByRole('button', { name: '确定' })).toBeEnabled();

  fireEvent.click(screen.getByRole('radio', { name: '左对齐' }));
  expect(indent).toBeEnabled();
  expect(screen.getByRole('button', { name: '确定' })).toBeDisabled();
  fireEvent.change(indent, {
    target: { value: '0.5' },
  });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(firstTableAttributes(editor)).toMatchObject({
    geometry: expect.objectContaining({
      width: { type: 'percent', value: 62.5 },
      alignment: 'left',
      indent: expect.closeTo(18.9, 1),
    }),
  });
  expect(screen.queryByRole('dialog', { name: '表格属性' })).toBeNull();
  await waitFor(() => expect(trigger).toHaveFocus());

  fireEvent.click(trigger);
  fireEvent.click(screen.getByRole('radio', { name: '居中' }));
  expect(
    screen.getByRole('textbox', { name: '表格左缩进（厘米）' }),
  ).toBeDisabled();
  fireEvent.click(screen.getByRole('button', { name: '取消' }));
  expect(firstTableAttributes(editor)).toMatchObject({
    geometry: expect.objectContaining({
      alignment: 'left',
      indent: expect.closeTo(18.9, 1),
    }),
  });

  fireEvent.click(trigger);
  const width = screen.getByRole('textbox', { name: '表格宽度（百分比）' });
  fireEvent.change(width, { target: { value: '80' } });
  fireEvent.keyDown(width, { key: 'Escape' });
  expect(screen.queryByRole('dialog', { name: '表格属性' })).toBeNull();
  expect(firstTableAttributes(editor)).toMatchObject({
    geometry: expect.objectContaining({
      width: { type: 'percent', value: 62.5 },
    }),
  });
  await waitFor(() => expect(trigger).toHaveFocus());
});

test('applies table, row, column, and cell tabs as one undoable dialog action', () => {
  editor = createTableEditor({ partialCellMargins: true });
  editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);
  const originalHtml = editor.getHTML();
  let updateCount = 0;
  editor.on('update', () => {
    updateCount += 1;
  });
  render(<DocumentTableLayoutRibbon editor={editor} />);

  fireEvent.click(screen.getByRole('button', { name: '表格属性' }));
  expect(
    screen.getByRole('tablist', { name: '表格属性分类' }),
  ).toBeInTheDocument();

  fireEvent.click(screen.getByRole('tab', { name: '行' }));
  fireEvent.click(screen.getByRole('checkbox', { name: '指定行高' }));
  fireEvent.change(screen.getByRole('textbox', { name: '当前行高（厘米）' }), {
    target: { value: '1.2' },
  });
  fireEvent.click(screen.getByRole('combobox', { name: '行高规则' }));
  fireEvent.click(screen.getByRole('option', { name: '固定值' }));
  fireEvent.click(screen.getByRole('checkbox', { name: '允许跨页断行' }));
  fireEvent.click(
    screen.getByRole('checkbox', { name: '在各页顶端重复标题行' }),
  );

  fireEvent.click(screen.getByRole('tab', { name: '列' }));
  fireEvent.change(screen.getByRole('textbox', { name: '当前列宽（厘米）' }), {
    target: { value: '4' },
  });

  fireEvent.click(screen.getByRole('tab', { name: '单元格' }));
  fireEvent.click(screen.getByRole('radio', { name: '居中' }));
  expect(
    screen.getByRole('checkbox', { name: '使用表格默认边距' }),
  ).not.toBeChecked();
  fireEvent.change(
    screen.getByRole('textbox', { name: '当前单元格左边距（厘米）' }),
    { target: { value: '0.4' } },
  );

  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(updateCount).toBe(1);
  expect(tableRowAttributes(editor)[0]).toMatchObject({
    rowHeight: expect.closeTo(45.35, 1),
    rowHeightRule: 'exact',
    cantSplit: true,
    repeatHeader: false,
  });
  expect(tableCellAttributes(editor)[0]).toMatchObject({
    colwidth: [expect.closeTo(151.18, 1)],
    verticalAlign: 'middle',
    margins: {
      top: 8,
      right: 16,
      left: expect.closeTo(15.12, 1),
    },
  });

  expect(editor.commands.undo()).toBe(true);
  expect(editor.getHTML()).toBe(originalHtml);
});

test('supports centimeter and automatic preferred table widths', async () => {
  editor = createTableEditor();
  editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);
  render(<DocumentTableLayoutRibbon editor={editor} />);

  const trigger = screen.getByRole('button', { name: '表格属性' });
  fireEvent.click(trigger);
  fireEvent.click(screen.getByRole('radio', { name: '厘米' }));
  const width = screen.getByRole('textbox', { name: '表格宽度（厘米）' });
  expect(width).toHaveValue('15');
  fireEvent.change(width, { target: { value: '10' } });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(firstTableAttributes(editor)).toMatchObject({
    geometry: expect.objectContaining({
      width: {
        type: 'pixels',
        value: expect.closeTo(377.95, 1),
      },
    }),
  });

  fireEvent.click(trigger);
  fireEvent.click(screen.getByRole('radio', { name: '自动' }));
  expect(screen.queryByRole('textbox', { name: /表格宽度/ })).toBeNull();
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(firstTableAttributes(editor)).toMatchObject({
    geometry: expect.objectContaining({
      width: { type: 'auto', value: null },
    }),
  });
});

test('authors a percentage width for the selected table column', () => {
  editor = createTableEditor();
  editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);
  render(<DocumentTableLayoutRibbon editor={editor} />);

  fireEvent.click(screen.getByRole('button', { name: '表格属性' }));
  fireEvent.click(screen.getByRole('tab', { name: '列' }));
  fireEvent.click(screen.getByRole('radio', { name: '百分比' }));
  const width = screen.getByRole('textbox', { name: '当前列宽（百分比）' });
  fireEvent.change(width, { target: { value: '35' } });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(tableCellAttributes(editor)[0]).toMatchObject({
    columnWidthPercentages: [35],
  });
  expect(tableCellAttributes(editor)[1]).toMatchObject({
    columnWidthPercentages: [50],
  });
  expect(documentTableSizing(editor.state)).toMatchObject({
    columnWidthType: 'percent',
    columnWidth: 35,
  });
});

test('shows the rendered column width for an autofit table', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: [
      '<section data-document-section="true">',
      '<table data-office-table-layout="autofit" ',
      'data-office-table-width-type="percent" data-office-table-width="62.5">',
      '<tbody><tr><td colwidth="6.67"><p>A</p></td>',
      '<td colwidth="6.67"><p>B</p></td></tr></tbody></table>',
      '</section>',
    ].join(''),
  });
  editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);
  const table = editor.view.dom.querySelector<HTMLTableElement>('table');
  const columns = Array.from(
    editor.view.dom.querySelectorAll<HTMLElement>('colgroup > col'),
  );
  if (!table || columns.length !== 2) {
    throw new Error('Expected one two-column table.');
  }
  table.getBoundingClientRect = () => elementRect(320, 48);
  for (const column of columns) {
    column.getBoundingClientRect = () => elementRect(160, 0);
  }

  render(<DocumentTableLayoutRibbon editor={editor} />);

  expect(screen.getByRole('textbox', { name: '列宽（厘米）' })).toHaveValue(
    '4.23',
  );
  expect(
    screen.getByRole('combobox', { name: '表格自动调整' }),
  ).toHaveTextContent('当前宽度');
});

test('disables row and column distribution for a one-cell table', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: [
      '<section data-document-section="true">',
      '<table><tbody><tr><td><p>Only cell</p></td></tr></tbody></table>',
      '</section>',
    ].join(''),
  });
  editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);
  render(<DocumentTableLayoutRibbon editor={editor} />);

  expect(screen.getByRole('button', { name: '平均分布行' })).toBeDisabled();
  expect(screen.getByRole('button', { name: '平均分布列' })).toBeDisabled();
});

test('cancels a dirty table dimension without changing the selected cells', () => {
  editor = createTableEditor();
  editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);
  expect(editor.commands.setCellAttribute('colwidth', [180])).toBe(true);
  const selectionBefore = editor.state.selection.toJSON();
  render(<DocumentTableLayoutRibbon editor={editor} />);

  const width = screen.getByRole('textbox', { name: '列宽（厘米）' });
  const committedWidth = width.getAttribute('value');
  fireEvent.change(width, { target: { value: '8.5' } });
  expect(width).toHaveAttribute('data-office-escape-consumer', 'true');

  fireEvent.keyDown(width, { key: 'Escape' });

  expect(width).toHaveValue(committedWidth);
  expect(width).not.toHaveAttribute('data-office-escape-consumer');
  expect(Number(tableCellAttributes(editor)[0]?.colwidth?.[0])).toBe(180);
  expect(editor.state.selection.toJSON()).toEqual(selectionBefore);
});

test('preserves a dirty table dimension while its live baseline changes', () => {
  editor = createTableEditor();
  editor.commands.setTextSelection(tableCellPositions(editor)[0] + 2);
  expect(editor.commands.setCellAttribute('colwidth', [180])).toBe(true);
  const view = render(<DocumentTableLayoutRibbon editor={editor} />);

  const width = screen.getByRole('textbox', { name: '列宽（厘米）' });
  fireEvent.change(width, { target: { value: '8.5' } });
  expect(width).toHaveValue('8.5');

  expect(editor.commands.setDocumentTableColumnWidth(200)).toBe(true);
  view.rerender(<DocumentTableLayoutRibbon editor={editor} />);

  expect(width).toHaveValue('8.5');
  fireEvent.keyDown(width, { key: 'Escape' });
  expect(width).toHaveValue('5.29');
  expect(Number(tableCellAttributes(editor)[0]?.colwidth?.[0])).toBe(200);
});

test('keeps whole-table controls truthful for a table node selection', () => {
  editor = createTableEditor();
  editor.commands.setNodeSelection(firstTablePosition(editor));
  const view = render(<DocumentTableDesignRibbon editor={editor} />);

  fireEvent.click(
    screen.getByRole('radio', { name: '应用表格样式：绿色条纹' }),
  );
  expect(
    tableCellAttributes(editor).every(
      ({ borderColor }) => borderColor === '#abc8be',
    ),
  ).toBe(true);

  view.rerender(<DocumentTableLayoutRibbon editor={editor} />);
  expect(screen.getByRole('button', { name: '在上方插入行' })).toBeDisabled();
  expect(screen.getByRole('button', { name: '跨页重复标题' })).toBeDisabled();
  const keepRowsTogether = screen.getByRole('button', {
    name: '整行不跨页',
  });
  expect(keepRowsTogether).toBeEnabled();
  fireEvent.click(keepRowsTogether);
  expect(
    tableRowAttributes(editor).every(({ cantSplit }) => cantSplit === true),
  ).toBe(true);

  const width = screen.getByRole('textbox', { name: '列宽（厘米）' });
  fireEvent.change(width, { target: { value: '3.2' } });
  fireEvent.blur(width);
  expect(
    tableCellAttributes(editor).every(({ colwidth }) =>
      Array.isArray(colwidth)
        ? colwidth.every((value) => Number(value) > 120)
        : false,
    ),
  ).toBe(true);
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

function createTableEditor({
  partialCellMargins = false,
}: {
  partialCellMargins?: boolean;
} = {}): Editor {
  const firstCellMargins = partialCellMargins
    ? ' data-office-cell-margin-top="8" data-office-cell-margin-right="16"'
    : '';
  return new Editor({
    extensions: createWorkDocumentExtensions(),
    content: [
      '<section data-document-section="true">',
      '<table><tbody>',
      '<tr data-office-repeat-header="true">',
      `<th${firstCellMargins}><p>Title A</p></th><th><p>Title B</p></th>`,
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
      layout={{
        pageSize: 'a4',
        orientation: 'portrait',
        margins: { top: 25, right: 23, bottom: 25, left: 23 },
        columns: { count: 1, spacing: 12, separator: false },
        breakAfter: 'nextPage',
      }}
      layoutOpen={false}
      navigationOpen={false}
      pageColor="#ffffff"
      showPageNumbers
      showHiddenText={false}
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
      onLayoutChange={noop}
      onOpenLayout={noop}
      onToggleNavigation={noop}
      onToggleHiddenText={noop}
      onTogglePageNumbers={noop}
      onToggleRulers={noop}
      onPageColorChange={noop}
      onToggleSpellcheck={noop}
      onViewModeChange={noop}
      onZoomChange={noop}
      onZoomFit={noop}
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

function firstTablePosition(currentEditor: Editor): number {
  let position: number | null = null;
  currentEditor.state.doc.descendants((node, offset) => {
    if (position === null && node.type.name === 'table') position = offset;
    return position === null;
  });
  if (position === null) throw new Error('Expected a table node.');
  return position;
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

function elementRect(width: number, height: number): DOMRect {
  return {
    bottom: height,
    height,
    left: 0,
    right: width,
    top: 0,
    width,
    x: 0,
    y: 0,
    toJSON: () => ({}),
  } as DOMRect;
}

function mockMatchMedia(matches: boolean): () => void {
  const originalDescriptor = Object.getOwnPropertyDescriptor(
    window,
    'matchMedia',
  );
  Object.defineProperty(window, 'matchMedia', {
    configurable: true,
    value: () => ({
      matches,
      media: '(max-width: 520px)',
      onchange: null,
      addEventListener: () => undefined,
      removeEventListener: () => undefined,
      addListener: () => undefined,
      removeListener: () => undefined,
      dispatchEvent: () => true,
    }),
  });
  return () => {
    if (originalDescriptor) {
      Object.defineProperty(window, 'matchMedia', originalDescriptor);
    } else {
      Reflect.deleteProperty(window, 'matchMedia');
    }
  };
}
