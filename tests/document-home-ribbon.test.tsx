import { Editor } from '@tiptap/core';
import { afterEach, expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import {
  changeDocumentFontSize,
  documentFontFamilyOptions,
} from '../src/internal/features/work/editors/document-formatting-options';
import { DocumentHomeRibbon } from '../src/internal/features/work/editors/document-home-ribbon';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { MAX_DOCUMENT_INDENT_LEVEL } from '../src/internal/features/work/work-document-paragraph-formatting';
import {
  OFFICE_DOCUMENT_LAYOUT_FONT_FAMILY,
  OFFICE_DOCUMENT_LAYOUT_FONT_ID,
} from '../src/internal/features/work/work-document-fonts';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

test('switches superscript and subscript without stacking both marks', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p><sub>water</sub> and <sup>power</sup></p>',
  });
  editor.commands.setTextSelection(textRange(editor, 'water'));
  const view = render(
    <DocumentHomeRibbon editor={editor} onFindText={() => undefined} />,
  );

  expect(screen.getByRole('button', { name: '下标' })).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  expect(screen.getByRole('button', { name: '上标' })).toHaveAttribute(
    'aria-pressed',
    'false',
  );

  fireEvent.click(screen.getByRole('button', { name: '上标' }));
  expect(editor.getHTML()).toContain('<sup>water</sup>');
  expect(editor.getHTML()).not.toContain('<sub>water</sub>');
  view.rerender(
    <DocumentHomeRibbon editor={editor} onFindText={() => undefined} />,
  );
  expect(screen.getByRole('button', { name: '下标' })).toHaveAttribute(
    'aria-pressed',
    'false',
  );
  expect(screen.getByRole('button', { name: '上标' })).toHaveAttribute(
    'aria-pressed',
    'true',
  );

  fireEvent.click(screen.getByRole('button', { name: '上标' }));
  expect(editor.getHTML()).not.toContain('<sup>water</sup>');

  editor.commands.setTextSelection(textRange(editor, 'power'));
  view.rerender(
    <DocumentHomeRibbon editor={editor} onFindText={() => undefined} />,
  );
  fireEvent.click(screen.getByRole('button', { name: '下标' }));
  expect(editor.getHTML()).toContain('<sub>power</sub>');
  expect(editor.getHTML()).not.toContain('<sup>power</sup>');
});

test('applies italic formatting without replacing the selected text', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>Italic text stays intact</p>',
  });
  editor.commands.setTextSelection(textRange(editor, 'Italic text'));
  render(
    <DocumentHomeRibbon
      editor={editor}
      findReplaceMode={null}
      onFindText={() => undefined}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '斜体' }));

  expect(editor.getHTML()).toContain('<em>Italic text</em> stays intact');
  expect(editor.getText()).toBe('Italic text stays intact');
  expect(screen.getByRole('button', { name: '斜体' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+I Meta+I',
  );
});

test('previews every font option with the font it applies', async () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>Font preview</p>',
  });
  render(
    <DocumentHomeRibbon
      editor={editor}
      findReplaceMode={null}
      onFindText={() => undefined}
    />,
  );

  fireEvent.click(screen.getByRole('combobox', { name: '字体' }));
  await waitFor(() =>
    expect(screen.getByRole('option', { name: '默认字体' })).toHaveFocus(),
  );
  for (const option of documentFontFamilyOptions) {
    if (!('previewStyle' in option)) continue;
    const menuOption = screen.getByRole('option', { name: option.label });
    expect(
      menuOption.querySelector('.work-office-select-option-copy'),
    ).toHaveStyle({ fontFamily: option.previewStyle.fontFamily });
  }
});

test('groups common fonts and includes host-provided layout fonts', async () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>Font catalog</p>',
  });
  render(
    <DocumentHomeRibbon
      editor={editor}
      findReplaceMode={null}
      layoutFonts={[
        {
          id: OFFICE_DOCUMENT_LAYOUT_FONT_ID,
          family: OFFICE_DOCUMENT_LAYOUT_FONT_FAMILY,
          url: '/fonts/noto-sans-hans.otf',
        },
        {
          id: 'brand-sans-regular',
          family: 'Brand Sans',
          url: '/fonts/brand-sans.ttf',
        },
      ]}
      onFindText={() => undefined}
    />,
  );

  fireEvent.click(screen.getByRole('combobox', { name: '字体' }));
  const options = screen.getAllByRole('option');
  expect(options.length).toBeGreaterThan(20);
  expect(screen.getByText('内置字体')).toBeInTheDocument();
  expect(screen.getByText('中文字体')).toBeInTheDocument();
  expect(screen.getByText('西文字体')).toBeInTheDocument();
  expect(screen.getByText('等宽字体')).toBeInTheDocument();
  expect(screen.getByRole('option', { name: '思源黑体' })).toBeInTheDocument();
  expect(
    screen.getByRole('option', { name: 'Brand Sans' }),
  ).toBeInTheDocument();
  expect(screen.getByRole('option', { name: '苹方' })).toBeInTheDocument();
  expect(screen.getByRole('option', { name: 'Calibri' })).toBeInTheDocument();
  expect(screen.getByRole('option', { name: 'Menlo' })).toBeInTheDocument();
});

test('keeps imported font family and size visible instead of reporting defaults', async () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content:
      '<p style="line-height: 1.2"><span style="font-family: Calibri; font-size: 11pt">Imported style</span></p>',
  });
  editor.commands.setTextSelection(textRange(editor, 'Imported style'));
  render(
    <DocumentHomeRibbon
      editor={editor}
      findReplaceMode={null}
      onFindText={() => undefined}
    />,
  );

  const family = screen.getByRole('combobox', { name: '字体' });
  const size = screen.getByRole('combobox', { name: '字号' });
  expect(family).toHaveTextContent('Calibri');
  expect(family.querySelector('span')).toHaveStyle({ fontFamily: 'Calibri' });
  expect(size).toHaveTextContent('11');
  expect(screen.getByRole('combobox', { name: '行距' })).toHaveTextContent(
    '1.2 倍',
  );

  fireEvent.click(family);
  const importedFamily = screen.getByRole('option', { name: 'Calibri' });
  expect(importedFamily).toHaveAttribute('aria-selected', 'true');
  await waitFor(() => expect(importedFamily).toHaveFocus());
});

test('wires every direct character-format action to the TipTap selection', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>Format this text</p>',
  });
  editor.commands.setTextSelection(textRange(editor, 'Format this'));
  render(
    <DocumentHomeRibbon
      editor={editor}
      findReplaceMode={null}
      onFindText={() => undefined}
    />,
  );
  const font = screen.getByRole('region', { name: '字体' });

  for (const label of ['加粗', '斜体', '下划线', '删除线']) {
    fireEvent.click(within(font).getByRole('button', { name: label }));
  }
  fireEvent.click(within(font).getByRole('combobox', { name: '字体' }));
  fireEvent.click(screen.getByRole('option', { name: 'Arial' }));
  fireEvent.click(within(font).getByRole('combobox', { name: '字号' }));
  fireEvent.click(screen.getByRole('option', { name: '14' }));
  fireEvent.click(within(font).getByRole('button', { name: '文字颜色' }));
  fireEvent.click(screen.getByRole('option', { name: '颜色 #0070c0' }));
  fireEvent.click(within(font).getByRole('button', { name: '突出显示' }));

  expect(textMarkNames(editor, 'Format this')).toEqual(
    new Set([
      'bold',
      'highlight',
      'italic',
      'strike',
      'textStyle',
      'underline',
    ]),
  );
  expect(editor.getAttributes('textStyle')).toMatchObject({
    color: '#0070c0',
    fontFamily: 'Arial, sans-serif',
    fontSize: '14pt',
  });
  expect(editor.getHTML()).toContain('font-family: Arial, sans-serif');
  expect(editor.getHTML()).toContain('font-size: 14pt');
  expect(editor.getHTML()).toContain('color: #0070c0');

  fireEvent.click(within(font).getByRole('button', { name: '清除格式' }));
  expect(textMarkNames(editor, 'Format this')).toEqual(new Set());
  expect(editor.getText()).toBe('Format this text');
});

test('steps font size and keeps Undo and Redo connected to editor history', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>History text</p>',
  });
  editor.commands.setTextSelection(textRange(editor, 'History text'));
  const view = render(
    <DocumentHomeRibbon
      editor={editor}
      findReplaceMode={null}
      onFindText={() => undefined}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '增大字号' }));
  expect(editor.getAttributes('textStyle').fontSize).toBe('12pt');
  fireEvent.click(screen.getByRole('button', { name: '减小字号' }));
  expect(editor.getAttributes('textStyle').fontSize).toBe('10.5pt');
  fireEvent.click(screen.getByRole('button', { name: '加粗' }));
  expect(editor.getHTML()).toContain('<strong>History text</strong>');

  view.rerender(
    <DocumentHomeRibbon
      editor={editor}
      findReplaceMode={null}
      onFindText={() => undefined}
    />,
  );
  fireEvent.click(screen.getByRole('button', { name: '撤销' }));
  expect(editor.getHTML()).not.toContain('<strong>History text</strong>');

  view.rerender(
    <DocumentHomeRibbon
      editor={editor}
      findReplaceMode={null}
      onFindText={() => undefined}
    />,
  );
  fireEvent.click(screen.getByRole('button', { name: '重做' }));
  expect(editor.getHTML()).toContain('<strong>History text</strong>');
});

test('disables font-size and indent commands at their real boundaries', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>Boundary controls</p>',
  });
  editor.commands.setTextSelection(textRange(editor, 'Boundary controls'));
  editor.chain().setFontSize('72pt').run();
  const view = render(
    <DocumentHomeRibbon
      editor={editor}
      findReplaceMode={null}
      onFindText={() => undefined}
    />,
  );

  expect(screen.getByRole('button', { name: '增大字号' })).toBeDisabled();
  expect(screen.getByRole('button', { name: '减小字号' })).toBeEnabled();
  expect(screen.getByRole('button', { name: '减少缩进' })).toBeDisabled();
  expect(changeDocumentFontSize(editor, 1)).toBe(false);
  expect(editor.commands.changeDocumentIndent(-1)).toBe(false);

  editor.chain().setFontSize('9pt').run();
  editor.commands.setDocumentIndentLevel(MAX_DOCUMENT_INDENT_LEVEL);
  view.rerender(
    <DocumentHomeRibbon
      editor={editor}
      findReplaceMode={null}
      onFindText={() => undefined}
    />,
  );

  expect(screen.getByRole('button', { name: '增大字号' })).toBeEnabled();
  expect(screen.getByRole('button', { name: '减小字号' })).toBeDisabled();
  expect(screen.getByRole('button', { name: '增加缩进' })).toBeDisabled();
  expect(editor.commands.changeDocumentIndent(1)).toBe(false);
});

test('wires paragraph alignment, direction, spacing, and indent controls', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>Paragraph controls</p>',
  });
  editor.commands.setTextSelection(textRange(editor, 'Paragraph').from);
  const findModes: boolean[] = [];
  render(
    <DocumentHomeRibbon
      editor={editor}
      findReplaceMode={null}
      onFindText={(replace) => findModes.push(replace)}
    />,
  );
  const paragraph = screen.getByRole('region', { name: '段落' });

  fireEvent.click(within(paragraph).getByRole('button', { name: '居中' }));
  expect(editor.getAttributes('paragraph').textAlign).toBe('center');
  fireEvent.click(within(paragraph).getByRole('button', { name: '右对齐' }));
  expect(editor.getAttributes('paragraph').textAlign).toBe('right');
  fireEvent.click(within(paragraph).getByRole('button', { name: '左对齐' }));
  expect(editor.getAttributes('paragraph').textAlign).toBe('left');
  fireEvent.click(within(paragraph).getByRole('button', { name: '两端对齐' }));
  expect(editor.getAttributes('paragraph').textAlign).toBe('justify');
  fireEvent.click(within(paragraph).getByRole('button', { name: '从右向左' }));
  expect(editor.getHTML()).toContain('dir="rtl"');
  fireEvent.click(within(paragraph).getByRole('button', { name: '从左向右' }));
  expect(editor.getHTML()).toContain('dir="ltr"');
  fireEvent.click(within(paragraph).getByRole('button', { name: '增加缩进' }));
  expect(editor.getHTML()).toContain('data-office-indent-level="1"');
  fireEvent.click(within(paragraph).getByRole('button', { name: '减少缩进' }));
  expect(editor.getHTML()).not.toContain('data-office-indent-level');
  fireEvent.click(within(paragraph).getByRole('button', { name: '增加缩进' }));
  fireEvent.click(within(paragraph).getByRole('combobox', { name: '行距' }));
  fireEvent.click(screen.getByRole('option', { name: '1.5 倍' }));

  expect(editor.getHTML()).toContain('dir="ltr"');
  expect(editor.getHTML()).toContain('data-office-indent-level="1"');
  expect(editor.getHTML()).toContain('line-height: 1.5');
  expect(editor.getText()).toBe('Paragraph controls');

  fireEvent.click(screen.getByRole('button', { name: '查找' }));
  fireEvent.click(screen.getByRole('button', { name: '替换' }));
  expect(findModes).toEqual([false, true]);
});

test('shows paragraph styles and applies the active style idempotently', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>Project brief</p>',
  });
  const view = render(
    <DocumentHomeRibbon
      editor={editor}
      findReplaceMode={null}
      onFindText={() => undefined}
    />,
  );
  const gallery = screen.getByRole('radiogroup', { name: '段落样式库' });
  const paragraph = within(gallery).getByRole('radio', {
    name: '应用样式：正文',
  });

  expect(paragraph).toBeChecked();
  fireEvent.click(
    within(gallery).getByRole('radio', { name: '应用样式：标题 1' }),
  );
  expect(editor.getHTML()).toContain('<h1>Project brief</h1>');
  const headingHtml = editor.getHTML();

  view.rerender(
    <DocumentHomeRibbon
      editor={editor}
      findReplaceMode={null}
      onFindText={() => undefined}
    />,
  );
  const activeHeading = within(
    screen.getByRole('radiogroup', { name: '段落样式库' }),
  ).getByRole('radio', { name: '应用样式：标题 1' });
  expect(activeHeading).toBeChecked();

  fireEvent.click(activeHeading);
  expect(editor.getHTML()).toBe(headingHtml);

  fireEvent.click(
    within(screen.getByRole('radiogroup', { name: '段落样式库' })).getByRole(
      'radio',
      { name: '应用样式：正文' },
    ),
  );
  expect(editor.getHTML()).toContain('<p>Project brief</p>');
  expect(editor.getHTML()).not.toContain('<h1>');
});

test('supports arrow-key selection in the paragraph style gallery', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>Keyboard styles</p>',
  });
  render(
    <DocumentHomeRibbon
      editor={editor}
      findReplaceMode={null}
      onFindText={() => undefined}
    />,
  );
  const gallery = screen.getByRole('radiogroup', { name: '段落样式库' });
  const paragraph = within(gallery).getByRole('radio', {
    name: '应用样式：正文',
  });
  const heading = within(gallery).getByRole('radio', {
    name: '应用样式：标题 1',
  });

  paragraph.focus();
  fireEvent.keyDown(paragraph, { key: 'ArrowRight' });

  expect(document.activeElement).toBe(heading);
  expect(editor.getHTML()).toContain('<h1>Keyboard styles</h1>');
});

test('uses a keyboard-operated bullet library without toggling the active style off', async () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>List item</p>',
  });
  render(
    <DocumentHomeRibbon
      editor={editor}
      findReplaceMode={null}
      onFindText={() => undefined}
    />,
  );

  const trigger = screen.getByRole('button', { name: '项目符号库' });
  fireEvent.click(trigger);
  let library = screen.getByRole('dialog', { name: '项目符号库' });
  const disc = within(library).getByRole('menuitemradio', {
    name: '实心圆点',
  });
  const circle = within(library).getByRole('menuitemradio', {
    name: '空心圆点',
  });

  await waitFor(() => expect(disc).toHaveFocus());
  fireEvent.keyDown(disc, { key: 'ArrowRight' });
  expect(circle).toHaveFocus();
  fireEvent.keyDown(circle, { key: 'Enter' });
  expect(editor.getHTML()).toContain('data-office-bullet-style="circle"');
  expect(
    screen.queryByRole('dialog', { name: '项目符号库' }),
  ).not.toBeInTheDocument();
  expect(trigger).toHaveFocus();

  fireEvent.click(trigger);
  library = screen.getByRole('dialog', { name: '项目符号库' });
  const activeCircle = within(library).getByRole('menuitemradio', {
    name: '空心圆点',
  });
  expect(activeCircle).toHaveAttribute('aria-checked', 'true');
  fireEvent.click(activeCircle);
  expect(editor.getHTML()).toContain('data-office-bullet-style="circle"');
});

test('keeps the list gallery roving tab stop aligned with arrow-key focus', async () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>List item</p>',
  });
  render(
    <DocumentHomeRibbon
      editor={editor}
      findReplaceMode={null}
      onFindText={() => undefined}
    />,
  );

  const trigger = screen.getByRole('button', { name: '项目符号库' });
  fireEvent.click(trigger);
  const library = screen.getByRole('dialog', { name: '项目符号库' });
  const disc = within(library).getByRole('menuitemradio', {
    name: '实心圆点',
  });
  const circle = within(library).getByRole('menuitemradio', {
    name: '空心圆点',
  });
  await waitFor(() => expect(disc).toHaveFocus());

  fireEvent.keyDown(disc, { key: 'ArrowRight' });
  expect(circle).toHaveFocus();
  expect(circle).toHaveAttribute('tabindex', '0');
  expect(disc).toHaveAttribute('tabindex', '-1');

  fireEvent.keyDown(circle, { key: 'Tab' });
  await waitFor(() =>
    expect(
      screen.queryByRole('dialog', { name: '项目符号库' }),
    ).not.toBeInTheDocument(),
  );
  expect(screen.getByRole('button', { name: '编号' })).toHaveFocus();
});

test('edits numbering style, start value, and continuation from the ribbon', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: [
      '<ol start="3" type="A">',
      '<li><p>First</p></li><li><p>Second</p></li>',
      '</ol>',
      '<p>Break</p>',
      '<ol><li><p>Third</p></li></ol>',
    ].join(''),
  });
  editor.commands.setTextSelection(textRange(editor, 'Third').from);
  render(
    <DocumentHomeRibbon
      editor={editor}
      findReplaceMode={null}
      onFindText={() => undefined}
    />,
  );

  const trigger = screen.getByRole('button', { name: '编号库' });
  fireEvent.click(trigger);
  let library = screen.getByRole('dialog', { name: '编号库' });
  fireEvent.click(
    within(library).getByRole('menuitemradio', { name: '大写罗马数字' }),
  );
  expect(editor.getHTML()).toContain('<ol type="I">');

  fireEvent.click(trigger);
  library = screen.getByRole('dialog', { name: '编号库' });
  const start = within(library).getByRole('textbox', { name: '起始编号' });
  fireEvent.change(start, { target: { value: '7' } });
  fireEvent.click(within(library).getByRole('button', { name: '应用起始值' }));
  expect(editor.getHTML()).toContain('<ol start="7" type="I">');

  fireEvent.click(trigger);
  library = screen.getByRole('dialog', { name: '编号库' });
  fireEvent.click(
    within(library).getByRole('button', { name: '继续前一列表' }),
  );
  expect(editor.getHTML()).toContain('<ol start="5" type="A">');
});

function textRange(editor: Editor, text: string): { from: number; to: number } {
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

function textMarkNames(editor: Editor, text: string): Set<string> {
  const range = textRange(editor, text);
  let names = new Set<string>();
  editor.state.doc.nodesBetween(range.from, range.to, (node) => {
    if (node.isText) names = new Set(node.marks.map((mark) => mark.type.name));
  });
  return names;
}
