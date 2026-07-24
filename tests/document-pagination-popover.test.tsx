import { Editor } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import { fireEvent, render, screen } from '@testing-library/react';
import { afterEach, expect, test } from '@rstest/core';
import { DocumentPaginationPopover } from '../src/internal/features/work/editors/document-pagination-popover';
import { DocumentParagraphFormatting } from '../src/internal/features/work/work-document-paragraph-formatting';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

test('edits typed paragraph pagination properties from an accessible popover', () => {
  editor = new Editor({
    extensions: [StarterKit, DocumentParagraphFormatting],
    content: '<p>A3S Office</p>',
  });

  render(<DocumentPaginationPopover editor={editor} />);
  fireEvent.click(screen.getByRole('button', { name: '段落分页' }));

  expect(
    screen.getByRole('dialog', { name: '段落分页选项' }),
  ).toBeInTheDocument();
  expect(
    screen.getByRole('checkbox', { name: '段落不跨页' }),
  ).not.toBeChecked();
  expect(
    screen.getByRole('checkbox', { name: '避免页首、页尾单行' }),
  ).toBeChecked();

  fireEvent.click(screen.getByRole('checkbox', { name: '段落不跨页' }));
  fireEvent.click(screen.getByRole('checkbox', { name: '与下一段同页' }));
  fireEvent.click(screen.getByRole('checkbox', { name: '段前另起一页' }));
  fireEvent.click(screen.getByRole('checkbox', { name: '避免页首、页尾单行' }));

  expect(editor.getHTML()).toContain('data-office-keep-lines="true"');
  expect(editor.getHTML()).toContain('data-office-keep-with-next="true"');
  expect(editor.getHTML()).toContain('data-office-page-break-before="true"');
  expect(editor.getHTML()).toContain('data-office-widow-control="false"');
});
