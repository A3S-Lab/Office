import { Editor } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import { fireEvent, render, screen } from '@testing-library/react';
import { expect, test } from '@rstest/core';
import { DocumentParagraphSpacingPopover } from '../src/internal/features/work/editors/document-paragraph-spacing-popover';
import { DocumentParagraphFormatting } from '../src/internal/features/work/work-document-paragraph-formatting';

test('edits and clears paragraph spacing from an accessible popover', () => {
  const editor = new Editor({
    extensions: [StarterKit, DocumentParagraphFormatting],
    content: '<p>A3S Office</p>',
  });

  render(<DocumentParagraphSpacingPopover editor={editor} />);

  fireEvent.click(screen.getByRole('button', { name: '段落间距' }));
  expect(
    screen.getByRole('dialog', { name: '段落间距选项' }),
  ).toBeInTheDocument();

  fireEvent.change(screen.getByRole('textbox', { name: '段前间距（磅）' }), {
    target: { value: '12' },
  });
  fireEvent.change(screen.getByRole('textbox', { name: '段后间距（磅）' }), {
    target: { value: '6' },
  });

  expect(editor.getHTML()).toContain('data-office-space-before="12"');
  expect(editor.getHTML()).toContain('data-office-space-after="6"');
  expect(editor.getHTML()).toContain('margin-top: 12pt');
  expect(editor.getHTML()).toContain('margin-bottom: 6pt');

  fireEvent.click(screen.getByRole('button', { name: '恢复默认间距' }));
  expect(editor.getHTML()).not.toContain('data-office-space-before');
  expect(editor.getHTML()).not.toContain('data-office-space-after');

  editor.destroy();
});
