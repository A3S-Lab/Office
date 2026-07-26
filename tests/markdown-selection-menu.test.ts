import { Editor } from '@tiptap/core';
import { afterEach, describe, expect, test } from '@rstest/core';
import { createWorkMarkdownExtensions } from '../src/internal/features/work/work-markdown-extensions';
import {
  createWorkMarkdownSourceSelectionAction,
  createWorkMarkdownSourceSelectionSnapshot,
  createWorkMarkdownVisualSelectionAction,
  createWorkMarkdownVisualSelectionSnapshot,
} from '../src/internal/features/work/work-markdown-selection-menu';
import type { WorkMarkdownContent } from '../src/internal/features/work/work-types';

const initialMarkdown = 'Before target text after.\n\nSecond paragraph.';
const content: WorkMarkdownContent = {
  type: 'markdown',
  markdown: initialMarkdown,
};

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

describe('Markdown selection menu context', () => {
  test('provides source selection context and keeps its live target between insertions', () => {
    const start = initialMarkdown.indexOf('target text');
    const snapshot = createWorkMarkdownSourceSelectionSnapshot(
      initialMarkdown,
      { start, end: start + 'target text'.length, direction: 'none' },
      content,
    );
    expect(snapshot).toMatchObject({
      selection: {
        surface: 'source',
        text: 'target text',
        beforeText: expect.stringContaining('Before '),
        afterText: expect.stringContaining(' after.'),
      },
      document: {
        content,
        markdown: initialMarkdown,
      },
    });
    if (!snapshot) throw new Error('Expected a source selection snapshot.');

    let markdown = initialMarkdown;
    const action = createWorkMarkdownSourceSelectionAction(
      snapshot,
      () => markdown,
      (edit) => {
        markdown = edit.markdown;
        return true;
      },
    );

    expect(action.context.commands.insertBefore('expanded ')).toEqual({
      applied: true,
    });
    expect(action.context.commands.insertAfter(' polished')).toEqual({
      applied: true,
    });
    expect(markdown).toContain('Before expanded target text polished after.');
    action.dispose();
    expect(action.context.commands.replaceText('replacement')).toEqual({
      applied: false,
      reason: 'editor-unavailable',
    });
  });

  test('rejects a source replacement after the selected text changes', () => {
    const start = initialMarkdown.indexOf('target text');
    const snapshot = createWorkMarkdownSourceSelectionSnapshot(
      initialMarkdown,
      { start, end: start + 'target text'.length, direction: 'none' },
      content,
    );
    if (!snapshot) throw new Error('Expected a source selection snapshot.');
    let markdown = initialMarkdown;
    const action = createWorkMarkdownSourceSelectionAction(
      snapshot,
      () => markdown,
      (edit) => {
        markdown = edit.markdown;
        return true;
      },
    );

    markdown = markdown.replace('target text', 'changed text');
    expect(action.context.commands.replaceText('replacement')).toEqual({
      applied: false,
      reason: 'stale-selection',
    });
    expect(markdown).toContain('Before changed text after.');
    action.dispose();
  });

  test('maps a visual target through unrelated edits and rejects target changes', () => {
    editor = new Editor({
      extensions: createWorkMarkdownExtensions(),
      content: initialMarkdown,
      contentType: 'markdown',
    });
    const range = textRange(editor, 'target text');
    editor.commands.setTextSelection(range);
    const snapshot = createWorkMarkdownVisualSelectionSnapshot(editor, content);
    if (!snapshot) throw new Error('Expected a visual selection snapshot.');
    expect(snapshot.selection.surface).toBe('visual');
    const action = createWorkMarkdownVisualSelectionAction(editor, snapshot);

    expect(editor.commands.insertContentAt(range.from, 'new ')).toBe(true);
    expect(action.context.commands.replaceText('replacement')).toEqual({
      applied: true,
    });
    expect(editor.getText()).toContain('Before new replacement after.');
    action.dispose();

    editor.commands.setTextSelection(textRange(editor, 'replacement'));
    const changedSnapshot = createWorkMarkdownVisualSelectionSnapshot(editor, {
      ...content,
      markdown: editor.getMarkdown(),
    });
    if (!changedSnapshot) {
      throw new Error('Expected a second visual selection snapshot.');
    }
    const changedAction = createWorkMarkdownVisualSelectionAction(
      editor,
      changedSnapshot,
    );
    expect(
      editor.commands.insertContentAt(
        textRange(editor, 'replacement'),
        'changed text',
      ),
    ).toBe(true);
    expect(changedAction.context.commands.replaceText('final text')).toEqual({
      applied: false,
      reason: 'stale-selection',
    });
    changedAction.dispose();
  });
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
  if (!range) throw new Error(`Unable to find "${text}" in the document.`);
  return range;
}
