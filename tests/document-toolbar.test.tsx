import { Editor } from '@tiptap/core';
import { afterEach, expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { DocumentToolbar } from '../src/internal/features/work/editors/document-toolbar';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

interface ToolbarCalls {
  captions: string[];
  fields: string[];
  files: string[];
  notes: string[];
  pageColors: string[];
  viewModes: string[];
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
  editor?.destroy();
  editorElement?.remove();
  editor = null;
});

test('wires every Insert and Page Layout action to document state or its owner', async () => {
  editor = createEditor();
  const calls = createCalls();
  render(toolbar(editor, calls));

  fireEvent.click(screen.getByRole('button', { name: '文件' }));
  fireEvent.click(screen.getByRole('menuitem', { name: '保存副本' }));
  expect(calls.files).toEqual(['save-copy']);

  fireEvent.click(screen.getByRole('tab', { name: '插入' }));
  fireEvent.click(screen.getByRole('button', { name: '插入图片' }));
  expect(calls.imageRequests).toBe(1);

  const firstBreakCount = nodeCount(editor, 'pageBreak');
  fireEvent.click(screen.getByRole('button', { name: '插入分页符' }));
  expect(nodeCount(editor, 'pageBreak')).toBe(firstBreakCount + 1);

  editor.commands.setTextSelection(textRange(editor, 'Toolbar text'));
  fireEvent.click(screen.getByRole('button', { name: '添加链接' }));
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

  fireEvent.click(screen.getByRole('combobox', { name: '插入页码或日期' }));
  fireEvent.click(screen.getByRole('option', { name: '当前日期' }));
  expect(calls.fields).toEqual(['date']);

  fireEvent.click(screen.getByRole('tab', { name: '页面布局' }));
  expect(screen.getByRole('button', { name: '段落间距' })).toBeEnabled();
  expect(screen.getByRole('button', { name: '段落分页' })).toBeEnabled();
  fireEvent.click(screen.getByRole('button', { name: '页面设置' }));
  fireEvent.click(screen.getByRole('button', { name: '显示页码' }));
  expect(calls.layout).toBe(1);
  expect(calls.pageNumbers).toBe(1);

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

test('wires every References, Review, and View action without silent buttons', () => {
  editor = createEditor();
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
  fireEvent.click(screen.getByRole('button', { name: '更新页码和日期' }));
  expect(calls.notes).toEqual(['footnote', 'endnote']);
  expect(calls.captions).toEqual(['figure', 'table']);
  expect(calls.crossReferences).toBe(1);
  expect(calls.toggleCitations).toBe(1);
  expect(calls.refreshFields).toBe(1);

  fireEvent.click(screen.getByRole('tab', { name: '审阅' }));
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
  for (const label of [
    '缩小文档',
    '缩放至 75%',
    '缩放至 100%',
    '缩放至 125%',
    '放大文档',
  ]) {
    fireEvent.click(screen.getByRole('button', { name: label }));
  }
  expect(calls.rulers).toBe(1);
  expect(calls.navigation).toBe(1);
  expect(calls.viewModes).toEqual(['page', 'web']);
  expect(calls.zooms).toEqual([90, 75, 100, 125, 110]);
});

function toolbar(currentEditor: Editor, calls: ToolbarCalls) {
  return (
    <DocumentToolbar
      editor={currentEditor}
      fileActions={[
        {
          id: 'save-copy',
          label: '保存副本',
          onSelect: () => calls.files.push('save-copy'),
        },
      ]}
      layoutOpen={false}
      navigationOpen={false}
      pageColor="#ffffff"
      showPageNumbers={false}
      showRulers={false}
      spellcheckEnabled
      viewMode="page"
      zoom={100}
      pageChromeEditor={null}
      pageChromeEditingPart={null}
      pageChromeShowPageNumber={false}
      onRequestImage={() => {
        calls.imageRequests += 1;
      }}
      onPageChromeEditingPartChange={() => undefined}
      onClosePageChrome={() => {
        calls.closePageChrome += 1;
      }}
      onTogglePageChromePageNumber={() => {
        calls.pageChromePageNumbers += 1;
      }}
      onToggleLayout={() => {
        calls.layout += 1;
      }}
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
    viewModes: [],
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

function nodeCount(currentEditor: Editor, type: string): number {
  let count = 0;
  currentEditor.state.doc.descendants((node) => {
    if (node.type.name === type) count += 1;
  });
  return count;
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
