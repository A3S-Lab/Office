import { Editor } from '@tiptap/core';
import { afterEach, expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { DocumentHomeRibbon } from '../src/internal/features/work/editors/document-home-ribbon';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

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
