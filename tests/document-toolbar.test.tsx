import { afterEach, expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { Editor } from '@tiptap/core';
import { clearDocumentFormatClipboard } from '../src/internal/features/work/editors/document-format-clipboard';
import { createDocumentPageChromeEditorExtensions } from '../src/internal/features/work/editors/document-page-chrome-editor';
import { DocumentToolbar } from '../src/internal/features/work/editors/document-toolbar';
import { collectDocumentChanges } from '../src/internal/features/work/work-document-changes';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import type { WorkDocumentSectionLayout } from '../src/internal/features/work/work-types';

const toolbarLayout: WorkDocumentSectionLayout = {
  pageSize: 'a4',
  orientation: 'portrait',
  margins: { top: 25, right: 23, bottom: 25, left: 23 },
  columns: { count: 1, spacing: 12, separator: false },
  breakAfter: 'nextPage',
};

interface ToolbarCalls {
  captions: string[];
  fields: string[];
  files: string[];
  notes: string[];
  pageColors: string[];
  pageChromeParts: string[];
  viewModes: string[];
  zoomFits: string[];
  zooms: number[];
  closePageChrome: number;
  comments: number;
  crossReferences: number;
  imageRequests: number;
  insertComments: number;
  layout: number;
  navigation: number;
  pageChromePageNumbers: number;
  pageNumbers: number;
  refreshFields: number;
  rulers: number;
  sections: number;
  spellcheck: number;
  toggleChanges: number;
  toggleCitations: number;
  trackChanges: number;
}

let editor: Editor | null = null;

afterEach(() => {
  const editorElement = editor?.view.dom;
  const editorRoot = editorElement?.closest('.work-document-editor');
  editor?.destroy();
  if (editorRoot) editorRoot.remove();
  else editorElement?.remove();
  editor = null;
  clearDocumentFormatClipboard();
});

test('wires every Insert and Page Layout action to document state or its owner', async () => {
  editor = createEditor();
  const calls = createCalls();
  render(toolbar(editor, calls));

  const quickAccess = screen.getByRole('toolbar', {
    name: '快速访问工具栏',
  });
  expect(
    within(quickAccess).getByRole('button', { name: '撤销' }),
  ).toHaveAttribute('aria-keyshortcuts', 'Control+Z Meta+Z');
  expect(
    within(quickAccess).getByRole('button', { name: '重做' }),
  ).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+Z Meta+Shift+Z Control+Y Meta+Y',
  );

  fireEvent.click(screen.getByRole('button', { name: '文件' }));
  fireEvent.click(screen.getByRole('menuitem', { name: '保存副本' }));
  expect(calls.files).toEqual(['save-copy']);

  fireEvent.click(screen.getByRole('tab', { name: '插入' }));
  fireEvent.click(screen.getByRole('button', { name: '插入图片' }));
  expect(calls.imageRequests).toBe(1);
  fireEvent.click(screen.getByRole('button', { name: '页眉' }));
  fireEvent.click(screen.getByRole('button', { name: '页脚' }));
  expect(calls.pageChromeParts).toEqual(['header', 'footer']);

  const firstBreakCount = nodeCount(editor, 'pageBreak');
  fireEvent.click(screen.getByRole('button', { name: '插入分页符' }));
  expect(nodeCount(editor, 'pageBreak')).toBe(firstBreakCount + 1);
  expect(screen.getByRole('button', { name: '插入分页符' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Enter Meta+Enter',
  );

  editor.commands.setTextSelection(textRange(editor, 'Toolbar text'));
  const addLink = screen.getByRole('button', { name: '添加链接' });
  expect(addLink).toHaveAttribute('aria-keyshortcuts', 'Control+K Meta+K');
  fireEvent.click(addLink);
  const linkDialog = await screen.findByRole('dialog', { name: '添加链接' });
  fireEvent.change(
    within(linkDialog).getByRole('textbox', { name: '链接地址' }),
    {
      target: { value: 'https://a3s.dev/office' },
    },
  );
  fireEvent.click(within(linkDialog).getByRole('button', { name: '添加链接' }));
  await waitFor(() =>
    expect(editor?.getHTML()).toContain('href="https://a3s.dev/office"'),
  );

  editor.commands.setTextSelection(textRange(editor, 'Toolbar text'));
  fireEvent.click(screen.getByRole('button', { name: '添加书签' }));
  const bookmarkDialog = await screen.findByRole('dialog', {
    name: '添加书签',
  });
  expect(
    within(bookmarkDialog).getByRole('textbox', { name: '书签名称' }),
  ).toHaveFocus();
  fireEvent.change(
    within(bookmarkDialog).getByRole('textbox', { name: '书签名称' }),
    { target: { value: 'Toolbar_target' } },
  );
  fireEvent.click(
    within(bookmarkDialog).getByRole('button', { name: '添加书签' }),
  );
  await waitFor(() =>
    expect(nodeCount(editor as Editor, 'documentBookmarkBoundary')).toBe(2),
  );

  const fieldSelect = screen.getByRole('combobox', {
    name: '插入页码或日期',
  });
  fireEvent.click(fieldSelect);
  expect(screen.getByRole('option', { name: '页码或日期' })).toBeDisabled();
  await waitFor(() =>
    expect(screen.getByRole('option', { name: '页码' })).toHaveFocus(),
  );
  fireEvent.click(screen.getByRole('option', { name: '当前日期' }));
  expect(calls.fields).toEqual(['date']);
  fireEvent.click(screen.getByRole('button', { name: '页码' }));
  expect(calls.pageNumbers).toBe(1);

  fireEvent.click(screen.getByRole('tab', { name: '页面布局' }));
  expect(screen.getByRole('button', { name: '段落间距' })).toBeEnabled();
  expect(screen.getByRole('button', { name: '段落分页' })).toBeEnabled();
  fireEvent.click(screen.getByRole('button', { name: '页面设置' }));
  expect(calls.layout).toBe(1);

  fireEvent.click(screen.getByRole('button', { name: '页面颜色' }));
  const colorDialog = screen.getByRole('dialog', { name: '页面颜色' });
  expect(within(colorDialog).getByText('主题颜色')).toBeVisible();
  expect(within(colorDialog).getByText('标准色')).toBeVisible();
  fireEvent.click(
    within(colorDialog).getByRole('option', { name: '颜色 #fff2cc' }),
  );
  expect(calls.pageColors).toEqual(['#fff2cc']);

  const secondBreakCount = nodeCount(editor, 'pageBreak');
  fireEvent.click(screen.getByRole('button', { name: '插入分页符' }));
  fireEvent.click(screen.getByRole('button', { name: '插入分节符' }));
  expect(nodeCount(editor, 'pageBreak')).toBe(secondBreakCount + 1);
  expect(calls.sections).toBe(1);
});

test('orders Insert and Page Layout groups like WPS Writer', () => {
  editor = createEditor();
  render(toolbar(editor, createCalls()));

  fireEvent.click(screen.getByRole('tab', { name: '插入' }));
  expect(activeRibbonGroupLabels()).toEqual([
    '页面',
    '表格',
    '插图',
    '链接',
    '页眉和页脚',
    '文本',
  ]);
  expect(screen.getByRole('button', { name: '页码' })).toBeVisible();

  fireEvent.click(screen.getByRole('tab', { name: '页面布局' }));
  expect(activeRibbonGroupLabels()).toEqual(['页面设置', '段落', '页面背景']);
  expect(screen.queryByRole('button', { name: '页码' })).toBeNull();
});

test('orders References, Review, and View groups like WPS Writer', () => {
  editor = createEditor();
  render(toolbar(editor, createCalls()));

  fireEvent.click(screen.getByRole('tab', { name: '引用' }));
  expect(activeRibbonGroupLabels()).toEqual([
    '脚注',
    '题注',
    '引文和书目',
    '更新',
  ]);

  fireEvent.click(screen.getByRole('tab', { name: '审阅' }));
  expect(activeRibbonGroupLabels()).toEqual(['校对', '批注', '修订', '更改']);

  fireEvent.click(screen.getByRole('tab', { name: '视图' }));
  expect(activeRibbonGroupLabels()).toEqual(['文档视图', '显示', '缩放 100%']);
});

test('wires every References, Review, and View action without silent buttons', () => {
  editor = createEditor();
  editor.commands.insertDocumentField('date');
  const calls = createCalls();
  render(toolbar(editor, calls));

  fireEvent.click(screen.getByRole('tab', { name: '引用' }));
  for (const label of ['插入脚注', '插入尾注']) {
    fireEvent.click(screen.getByRole('button', { name: label }));
  }
  for (const label of ['插入图片题注', '插入表格题注']) {
    fireEvent.click(screen.getByRole('button', { name: label }));
  }
  fireEvent.click(screen.getByRole('button', { name: '插入交叉引用' }));
  fireEvent.click(screen.getByRole('button', { name: '文献库（2）' }));
  const refreshFields = screen.getByRole('button', {
    name: '更新页码和日期',
  });
  expect(refreshFields).toHaveAttribute('aria-keyshortcuts', 'F9');
  fireEvent.click(refreshFields);
  expect(calls.notes).toEqual(['footnote', 'endnote']);
  expect(calls.captions).toEqual(['figure', 'table']);
  expect(calls.crossReferences).toBe(1);
  expect(calls.toggleCitations).toBe(1);
  expect(calls.refreshFields).toBe(1);

  fireEvent.click(screen.getByRole('tab', { name: '审阅' }));
  expect(screen.getByRole('button', { name: '拼写检查' })).toHaveAttribute(
    'aria-keyshortcuts',
    'F7',
  );
  expect(screen.getByRole('button', { name: '添加批注' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Alt+M Meta+Alt+M',
  );
  expect(screen.getByRole('button', { name: '修订模式' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+E Meta+Shift+E',
  );
  for (const label of [
    '拼写检查',
    '添加批注',
    '查看批注（3）',
    '修订模式',
    '查看修订（4）',
  ]) {
    fireEvent.click(screen.getByRole('button', { name: label }));
  }
  expect(calls.spellcheck).toBe(1);
  expect(calls.insertComments).toBe(1);
  expect(calls.comments).toBe(1);
  expect(calls.trackChanges).toBe(1);
  expect(calls.toggleChanges).toBe(1);

  fireEvent.click(screen.getByRole('tab', { name: '视图' }));
  fireEvent.click(screen.getByRole('button', { name: '标尺' }));
  fireEvent.click(screen.getByRole('button', { name: '导航窗格' }));
  fireEvent.click(screen.getByRole('button', { name: '页面视图' }));
  fireEvent.click(screen.getByRole('button', { name: '网页视图' }));
  for (const label of ['缩小文档', '缩放至 100%', '单页', '页宽', '放大文档']) {
    fireEvent.click(screen.getByRole('button', { name: label }));
  }
  expect(calls.rulers).toBe(1);
  expect(calls.navigation).toBe(1);
  expect(calls.viewModes).toEqual(['page', 'web']);
  expect(calls.zooms).toEqual([90, 100, 110]);
  expect(calls.zoomFits).toEqual(['page', 'width']);
});

test('navigates, accepts, and rejects tracked changes from the Review ribbon', () => {
  editor = createEditorWithChanges();
  const changes = collectDocumentChanges(editor.state.doc);
  editor.commands.setTextSelection({
    from: changes[0].from,
    to: changes[0].to,
  });
  render(toolbar(editor, createCalls()));

  fireEvent.click(screen.getByRole('tab', { name: '审阅' }));
  fireEvent.click(screen.getByRole('button', { name: '下一处修订' }));
  expect(editor.state.selection.from).toBe(changes[1].from);
  expect(editor.state.selection.to).toBe(changes[1].to);

  fireEvent.click(screen.getByRole('button', { name: '拒绝修订' }));
  expect(
    collectDocumentChanges(editor.state.doc).map((change) => change.id),
  ).toEqual(['change-one']);

  fireEvent.click(screen.getByRole('button', { name: '接受修订' }));
  expect(collectDocumentChanges(editor.state.doc)).toHaveLength(0);
  expect(editor.state.selection.empty).toBe(true);
});

test('only enables field refresh when the document contains fields', () => {
  editor = createEditor();
  const calls = createCalls();
  const view = render(toolbar(editor, calls));
  fireEvent.click(screen.getByRole('tab', { name: '引用' }));

  const refresh = screen.getByRole('button', { name: '更新页码和日期' });
  expect(refresh).toBeDisabled();
  fireEvent.click(refresh);
  expect(calls.refreshFields).toBe(0);

  editor.commands.insertDocumentField('page');
  view.rerender(toolbar(editor, calls));
  expect(refresh).toBeEnabled();
  fireEvent.click(refresh);
  expect(calls.refreshFields).toBe(1);
});

test('keeps body link shortcuts out of non-document editing surfaces', async () => {
  editor = createEditor();
  const root = document.createElement('section');
  root.className = 'work-document-editor';
  const auxiliaryInput = document.createElement('input');
  root.append(editor.view.dom, auxiliaryInput);
  document.body.append(root);
  render(toolbar(editor, createCalls()));

  fireEvent.keyDown(auxiliaryInput, { key: 'k', ctrlKey: true });
  expect(screen.queryByRole('dialog', { name: '添加链接' })).toBeNull();

  fireEvent.keyDown(editor.view.dom, { key: 'k', ctrlKey: true });
  const dialog = await screen.findByRole('dialog', { name: '添加链接' });
  expect(dialog).toBeVisible();
  fireEvent.keyDown(dialog, { key: 'Escape' });
  await waitFor(() =>
    expect(screen.queryByRole('dialog', { name: '添加链接' })).toBeNull(),
  );
});

test('routes history shortcuts from ribbon controls without stealing native text undo', () => {
  editor = createEditor();
  const root = document.createElement('section');
  root.className = 'work-document-editor';
  const ribbonButton = document.createElement('button');
  ribbonButton.textContent = '表格属性';
  const ribbonToggle = document.createElement('input');
  ribbonToggle.type = 'checkbox';
  const fileName = document.createElement('input');
  root.append(editor.view.dom, ribbonButton, ribbonToggle, fileName);
  document.body.append(root);
  render(toolbar(editor, createCalls()));

  const originalHtml = editor.getHTML();
  editor.chain().focus().insertContent('已修改').run();
  const changedHtml = editor.getHTML();
  expect(changedHtml).not.toBe(originalHtml);

  fireEvent.keyDown(fileName, { key: 'z', ctrlKey: true });
  expect(editor.getHTML()).toBe(changedHtml);

  fireEvent.keyDown(ribbonToggle, { key: 'z', ctrlKey: true });
  expect(editor.getHTML()).toBe(originalHtml);

  fireEvent.keyDown(ribbonButton, { key: 'y', ctrlKey: true });
  expect(editor.getHTML()).toBe(changedHtml);
});

test('routes WPS Writer formatting and review shortcuts inside the document', () => {
  editor = createEditor();
  const calls = createCalls();
  const root = document.createElement('section');
  root.className = 'work-document-editor';
  const auxiliaryInput = document.createElement('input');
  root.append(editor.view.dom, auxiliaryInput);
  document.body.append(root);
  render(toolbar(editor, calls));
  editor.commands.setTextSelection(textRange(editor, 'Toolbar text'));

  fireEvent.keyDown(auxiliaryInput, { key: 'e', ctrlKey: true });
  expect(editor.getAttributes('paragraph').textAlign).toBeNull();

  fireEvent.keyDown(editor.view.dom, { key: 'e', ctrlKey: true });
  expect(editor.getAttributes('paragraph').textAlign).toBe('center');

  fireEvent.keyDown(editor.view.dom, {
    key: '>',
    code: 'Period',
    ctrlKey: true,
    shiftKey: true,
  });
  expect(editor.getAttributes('textStyle').fontSize).toBe('12pt');

  fireEvent.keyDown(editor.view.dom, { key: '1', ctrlKey: true });
  expect(editor.getAttributes('paragraph').lineHeight).toBe('1');

  fireEvent.keyDown(editor.view.dom, {
    key: '2',
    code: 'Digit2',
    altKey: true,
    ctrlKey: true,
  });
  expect(editor.isActive('heading', { level: 2 })).toBe(true);

  fireEvent.keyDown(editor.view.dom, {
    key: 'e',
    ctrlKey: true,
    shiftKey: true,
  });
  fireEvent.keyDown(editor.view.dom, {
    key: 'm',
    altKey: true,
    ctrlKey: true,
  });
  fireEvent.keyDown(editor.view.dom, { key: 'F7' });
  expect(calls.trackChanges).toBe(1);
  expect(calls.insertComments).toBe(1);
  expect(calls.spellcheck).toBe(1);
});

test('opens and applies advanced font spacing from the scoped shortcut and launcher', async () => {
  editor = createEditor();
  const root = document.createElement('section');
  root.className = 'work-document-editor';
  const auxiliaryInput = document.createElement('input');
  root.append(editor.view.dom, auxiliaryInput);
  document.body.append(root);
  render(toolbar(editor, createCalls()));
  const selection = textRange(editor, 'Toolbar text');
  editor.commands.setTextSelection(selection);

  const launcher = screen.getByRole('button', {
    name: '字体高级设置',
  });
  expect(launcher).toHaveAttribute('aria-keyshortcuts', 'Control+D Meta+D');

  fireEvent.keyDown(auxiliaryInput, { key: 'd', ctrlKey: true });
  expect(screen.queryByRole('dialog', { name: '字体高级设置' })).toBeNull();

  fireEvent.keyDown(editor.view.dom, { key: 'd', ctrlKey: true });
  const dialog = await screen.findByRole('dialog', {
    name: '字体高级设置',
  });
  expect(editor.state.selection.toJSON()).toEqual({
    type: 'text',
    anchor: selection.from,
    head: selection.to,
  });
  fireEvent.click(within(dialog).getByRole('combobox', { name: '字符间距' }));
  fireEvent.click(await screen.findByRole('option', { name: '加宽' }));
  fireEvent.change(
    within(dialog).getByRole('textbox', { name: '间距值（磅）' }),
    { target: { value: '1.5' } },
  );
  fireEvent.click(within(dialog).getByRole('button', { name: '应用' }));

  await waitFor(() =>
    expect(screen.queryByRole('dialog', { name: '字体高级设置' })).toBeNull(),
  );
  expect(editor.getAttributes('textStyle').characterSpacingTwips).toBe(30);
  expect(editor.getHTML()).toContain('letter-spacing: 1.5pt');
  expect(editor.view.dom).toHaveFocus();

  expect(editor.commands.undo()).toBe(true);
  editor.commands.setTextSelection(selection);
  expect(
    editor.getAttributes('textStyle').characterSpacingTwips,
  ).toBeUndefined();

  fireEvent.click(launcher);
  const reopened = await screen.findByRole('dialog', {
    name: '字体高级设置',
  });
  fireEvent.keyDown(
    within(reopened).getByRole('combobox', { name: '字符间距' }),
    { key: 'd', ctrlKey: true },
  );
  expect(screen.getAllByRole('dialog', { name: '字体高级设置' })).toHaveLength(
    1,
  );
  fireEvent.click(within(reopened).getByRole('button', { name: '取消' }));
});

test('routes the advanced font shortcut to the active page-chrome editor', async () => {
  editor = createEditor();
  const pageChromeEditor = new Editor({
    extensions: createDocumentPageChromeEditorExtensions(),
    content: '<p>Header spacing</p>',
  });
  const root = document.createElement('section');
  root.className = 'work-document-editor';
  root.append(editor.view.dom, pageChromeEditor.view.dom);
  document.body.append(root);
  render(toolbar(editor, createCalls(), 100, false, pageChromeEditor));
  pageChromeEditor.commands.setTextSelection(
    textRange(pageChromeEditor, 'Header spacing'),
  );

  fireEvent.keyDown(pageChromeEditor.view.dom, { key: 'd', ctrlKey: true });
  const dialog = await screen.findByRole('dialog', {
    name: '字体高级设置',
  });
  fireEvent.click(within(dialog).getByRole('combobox', { name: '字符间距' }));
  fireEvent.click(await screen.findByRole('option', { name: '紧缩' }));
  fireEvent.change(
    within(dialog).getByRole('textbox', { name: '间距值（磅）' }),
    { target: { value: '0.75' } },
  );
  fireEvent.click(within(dialog).getByRole('button', { name: '应用' }));

  await waitFor(() =>
    expect(pageChromeEditor.getHTML()).toContain('letter-spacing: -0.75pt'),
  );
  expect(
    pageChromeEditor.getAttributes('textStyle').characterSpacingTwips,
  ).toBe(-15);
  pageChromeEditor.destroy();
});

test('routes WPS copy-format and paste-format shortcuts inside the document', () => {
  editor = createEditor();
  const root = document.createElement('section');
  root.className = 'work-document-editor';
  const auxiliaryInput = document.createElement('input');
  root.append(editor.view.dom, auxiliaryInput);
  document.body.append(root);
  render(toolbar(editor, createCalls()));

  editor.commands.setTextSelection(textRange(editor, 'Toolbar'));
  editor.commands.setBold();
  fireEvent.keyDown(editor.view.dom, {
    key: 'c',
    ctrlKey: true,
    shiftKey: true,
  });

  editor.commands.setTextSelection(textRange(editor, 'text'));
  fireEvent.keyDown(auxiliaryInput, {
    key: 'v',
    ctrlKey: true,
    shiftKey: true,
  });
  expect(editor.getHTML()).not.toContain('<strong>text</strong>');

  fireEvent.keyDown(editor.view.dom, {
    key: 'v',
    ctrlKey: true,
    shiftKey: true,
  });
  expect(editor.getHTML()).toContain('<strong>text</strong>');
});

test('keeps quick access undo and redo connected to document history', () => {
  editor = createEditor();
  const calls = createCalls();
  const view = render(toolbar(editor, calls));
  const originalHtml = editor.getHTML();

  editor.chain().focus().insertContent('Quick access change').run();
  const changedHtml = editor.getHTML();
  expect(changedHtml).not.toBe(originalHtml);
  view.rerender(toolbar(editor, calls));

  let quickAccess = screen.getByRole('toolbar', {
    name: '快速访问工具栏',
  });
  fireEvent.click(within(quickAccess).getByRole('button', { name: '撤销' }));
  expect(editor.getHTML()).toBe(originalHtml);
  view.rerender(toolbar(editor, calls));

  quickAccess = screen.getByRole('toolbar', { name: '快速访问工具栏' });
  fireEvent.click(within(quickAccess).getByRole('button', { name: '重做' }));
  expect(editor.getHTML()).toBe(changedHtml);
});

test('disables document zoom buttons at the supported boundaries', () => {
  editor = createEditor();
  const calls = createCalls();
  const view = render(toolbar(editor, calls, 50));
  fireEvent.click(screen.getByRole('tab', { name: '视图' }));

  expect(screen.getByRole('button', { name: '缩小文档' })).toBeDisabled();
  expect(screen.getByRole('button', { name: '放大文档' })).toBeEnabled();

  view.rerender(toolbar(editor, calls, 200));
  expect(screen.getByRole('button', { name: '缩小文档' })).toBeEnabled();
  expect(screen.getByRole('button', { name: '放大文档' })).toBeDisabled();
});

test('can start with the ribbon panel collapsed without disabling editing tools', () => {
  editor = createEditor();
  render(toolbar(editor, createCalls(), 100, true));

  const ribbon = screen.getByRole('region', { name: '文字功能区' });
  expect(ribbon).toHaveAttribute('data-collapsed', 'true');
  expect(screen.getByRole('button', { name: '展开功能区' })).toHaveAttribute(
    'aria-expanded',
    'false',
  );

  fireEvent.click(screen.getByRole('button', { name: '展开功能区' }));
  expect(ribbon).not.toHaveAttribute('data-collapsed');
  expect(screen.getByRole('button', { name: '折叠功能区' })).toHaveAttribute(
    'aria-expanded',
    'true',
  );
});

function toolbar(
  currentEditor: Editor,
  calls: ToolbarCalls,
  zoom = 100,
  defaultRibbonCollapsed = false,
  currentPageChromeEditor: Editor | null = null,
) {
  return (
    <DocumentToolbar
      editor={currentEditor}
      defaultRibbonCollapsed={defaultRibbonCollapsed}
      fileActions={[
        {
          id: 'save-copy',
          label: '保存副本',
          onSelect: () => calls.files.push('save-copy'),
        },
      ]}
      layout={toolbarLayout}
      layoutOpen={false}
      navigationOpen={false}
      pageColor="#ffffff"
      showPageNumbers={false}
      showRulers={false}
      spellcheckEnabled
      viewMode="page"
      zoom={zoom}
      pageChromeEditor={currentPageChromeEditor}
      pageChromeEditingPart={currentPageChromeEditor ? 'header' : null}
      pageChromeShowPageNumber={false}
      onRequestImage={() => {
        calls.imageRequests += 1;
      }}
      onPageChromeEditingPartChange={(part) => calls.pageChromeParts.push(part)}
      onClosePageChrome={() => {
        calls.closePageChrome += 1;
      }}
      onTogglePageChromePageNumber={() => {
        calls.pageChromePageNumbers += 1;
      }}
      onToggleLayout={() => {
        calls.layout += 1;
      }}
      onLayoutChange={() => undefined}
      onOpenLayout={() => undefined}
      onToggleNavigation={() => {
        calls.navigation += 1;
      }}
      onTogglePageNumbers={() => {
        calls.pageNumbers += 1;
      }}
      onToggleRulers={() => {
        calls.rulers += 1;
      }}
      onPageColorChange={(color) => calls.pageColors.push(color)}
      onToggleSpellcheck={() => {
        calls.spellcheck += 1;
      }}
      onViewModeChange={(mode) => calls.viewModes.push(mode)}
      onZoomChange={(zoom) => calls.zooms.push(zoom)}
      onZoomFit={(fit) => calls.zoomFits.push(fit)}
      onInsertSection={() => {
        calls.sections += 1;
      }}
      onInsertNote={(kind) => calls.notes.push(kind)}
      onInsertCaption={(kind) => calls.captions.push(kind)}
      onInsertCrossReference={() => {
        calls.crossReferences += 1;
      }}
      citationsOpen={false}
      citationSourceCount={2}
      onToggleCitations={() => {
        calls.toggleCitations += 1;
      }}
      onInsertField={(kind) => calls.fields.push(kind)}
      onRefreshFields={() => {
        calls.refreshFields += 1;
      }}
      canInsertComment
      onInsertComment={() => {
        calls.insertComments += 1;
      }}
      commentsOpen={false}
      commentCount={3}
      onToggleComments={() => {
        calls.comments += 1;
      }}
      trackChanges={false}
      changesOpen={false}
      changeCount={4}
      findReplaceMode={null}
      onToggleTrackChanges={() => {
        calls.trackChanges += 1;
      }}
      onToggleChanges={() => {
        calls.toggleChanges += 1;
      }}
      onOpenFindReplace={() => undefined}
    />
  );
}

function createCalls(): ToolbarCalls {
  return {
    captions: [],
    fields: [],
    files: [],
    notes: [],
    pageColors: [],
    pageChromeParts: [],
    viewModes: [],
    zoomFits: [],
    zooms: [],
    closePageChrome: 0,
    comments: 0,
    crossReferences: 0,
    imageRequests: 0,
    insertComments: 0,
    layout: 0,
    navigation: 0,
    pageChromePageNumbers: 0,
    pageNumbers: 0,
    refreshFields: 0,
    rulers: 0,
    sections: 0,
    spellcheck: 0,
    toggleChanges: 0,
    toggleCitations: 0,
    trackChanges: 0,
  };
}

function createEditor(): Editor {
  const currentEditor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content:
      '<section data-document-section="true"><p>Toolbar text</p></section>',
  });
  document.body.append(currentEditor.view.dom);
  return currentEditor;
}

function createEditorWithChanges(): Editor {
  const currentEditor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content:
      '<section data-document-section="true"><p><ins data-document-change="true" data-change-kind="insertion" data-change-id="change-one">First</ins> plain <ins data-document-change="true" data-change-kind="insertion" data-change-id="change-two">Second</ins></p></section>',
  });
  document.body.append(currentEditor.view.dom);
  return currentEditor;
}

function nodeCount(currentEditor: Editor, type: string): number {
  let count = 0;
  currentEditor.state.doc.descendants((node) => {
    if (node.type.name === type) count += 1;
  });
  return count;
}

function activeRibbonGroupLabels(): string[] {
  return Array.from(
    document.querySelectorAll<HTMLElement>(
      '.document-toolbar .work-office-ribbon-group',
    ),
    (group) => group.getAttribute('aria-label') ?? '',
  );
}

function textRange(
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
