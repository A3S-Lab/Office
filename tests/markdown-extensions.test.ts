import { Editor } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import {
  createWorkMarkdownExtensions,
  markdownTaskCheckboxLabel,
} from '../src/internal/features/work/work-markdown-extensions';

describe('Markdown extensions', () => {
  test('composes the Markdown schema through one extension registry', () => {
    const editor = new Editor({
      extensions: createWorkMarkdownExtensions(),
      content: '# Initial',
      contentType: 'markdown',
    });

    expect(editor.extensionManager.extensions.map(({ name }) => name)).toEqual(
      expect.arrayContaining([
        'workMarkdown',
        'markdown',
        'taskList',
        'taskItem',
        'table',
      ]),
    );
    editor.destroy();
  });

  test('replaces controlled Markdown through a typed editor command', () => {
    const editor = new Editor({
      extensions: createWorkMarkdownExtensions(),
      content: '# Initial',
      contentType: 'markdown',
    });

    expect(
      editor.commands.setWorkMarkdown('## Updated\n\n- [x] Ready', {
        emitUpdate: false,
      }),
    ).toBe(true);
    expect(editor.getMarkdown()).toContain('## Updated');
    expect(editor.getMarkdown()).toContain('- [x] Ready');
    expect(editor.can().undo()).toBe(false);
    editor.destroy();
  });

  test('keeps task labels accessible', () => {
    expect(
      markdownTaskCheckboxLabel({
        attrs: { checked: true },
        textContent: '发布',
      }),
    ).toBe('已完成：发布');
  });
});
