import { Editor } from '@tiptap/core';
import { afterEach, expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
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
