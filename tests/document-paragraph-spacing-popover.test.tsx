import { Editor } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { expect, test } from '@rstest/core';
import { DocumentParagraphSpacingPopover } from '../src/internal/features/work/editors/document-paragraph-spacing-popover';
import { DocumentParagraphFormatting } from '../src/internal/features/work/work-document-paragraph-formatting';

test('edits and clears paragraph spacing from an accessible popover', async () => {
  const editor = new Editor({
    extensions: [StarterKit, DocumentParagraphFormatting],
    content: '<p>A3S Office</p>',
  });

  render(<DocumentParagraphSpacingPopover editor={editor} />);

  fireEvent.click(screen.getByRole('button', { name: '段落间距' }));
  expect(
    screen.getByRole('dialog', { name: '段落间距选项' }),
  ).toBeInTheDocument();
  await waitFor(() =>
    expect(
      screen.getByRole('textbox', { name: '段前间距（磅）' }),
    ).toHaveFocus(),
  );

  const before = screen.getByRole('textbox', { name: '段前间距（磅）' });
  fireEvent.change(before, { target: { value: '12.' } });
  expect(before).toHaveValue('12.');
  expect(editor.getHTML()).not.toContain('data-office-space-before');
  fireEvent.change(before, { target: { value: '12.5' } });
  fireEvent.keyDown(before, { key: 'Enter' });

  const after = screen.getByRole('textbox', { name: '段后间距（磅）' });
  fireEvent.change(after, { target: { value: '6' } });
  fireEvent.keyDown(after, { key: 'Enter' });

  expect(editor.getHTML()).toContain('data-office-space-before="12.5"');
  expect(editor.getHTML()).toContain('data-office-space-after="6"');
  expect(editor.getHTML()).toContain('margin-top: 12.5pt');
  expect(editor.getHTML()).toContain('margin-bottom: 6pt');

  fireEvent.click(screen.getByRole('button', { name: '恢复默认间距' }));
  expect(editor.getHTML()).not.toContain('data-office-space-before');
  expect(editor.getHTML()).not.toContain('data-office-space-after');

  fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'Escape' });
  expect(screen.queryByRole('dialog', { name: '段落间距选项' })).toBeNull();
  expect(screen.getByRole('button', { name: '段落间距' })).toHaveFocus();

  editor.destroy();
});

test('cancels a dirty spacing draft before Escape closes the popover', async () => {
  const editor = new Editor({
    extensions: [StarterKit, DocumentParagraphFormatting],
    content: '<p>A3S Office</p>',
  });

  render(<DocumentParagraphSpacingPopover editor={editor} />);
  const trigger = screen.getByRole('button', { name: '段落间距' });
  fireEvent.click(trigger);
  const before = screen.getByRole('textbox', { name: '段前间距（磅）' });
  await waitFor(() => expect(before).toHaveFocus());

  fireEvent.change(before, { target: { value: '18' } });
  fireEvent.keyDown(before, { key: 'Escape' });

  expect(
    screen.getByRole('dialog', { name: '段落间距选项' }),
  ).toBeInTheDocument();
  expect(before).toHaveValue('');
  expect(editor.getHTML()).not.toContain('data-office-space-before');

  fireEvent.keyDown(before, { key: 'Escape' });
  expect(screen.queryByRole('dialog', { name: '段落间距选项' })).toBeNull();
  expect(trigger).toHaveFocus();
  expect(editor.getHTML()).toBe('<p>A3S Office</p>');

  editor.destroy();
});
