import { Editor } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import {
  DocumentFindReplacePanel,
  documentTextMatches,
} from '../src/internal/features/work/editors/document-find-replace-panel';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

test('finds text across adjacent marked runs and moves through matches', () => {
  const { editor, element } = createEditor(
    '<p>Al<strong>pha</strong> beta Alpha</p>',
  );

  try {
    expect(documentTextMatches(editor, 'alpha')).toEqual([
      { from: 1, to: 6 },
      { from: 12, to: 17 },
    ]);
    render(
      <DocumentFindReplacePanel
        editor={editor}
        mode="find"
        onModeChange={() => undefined}
        onReplaceText={() => false}
        onClose={() => undefined}
      />,
    );

    fireEvent.change(screen.getByRole('textbox', { name: '查找内容' }), {
      target: { value: 'alpha' },
    });
    expect(screen.getByText('2 个匹配')).toBeVisible();

    fireEvent.click(screen.getByRole('button', { name: '下一个匹配' }));
    expect(
      editor.state.doc.textBetween(
        editor.state.selection.from,
        editor.state.selection.to,
      ),
    ).toBe('Alpha');
  } finally {
    editor.destroy();
    element.remove();
  }
});

test('replaces the current match and exposes the task-pane mode switch', () => {
  const { editor, element } = createEditor('<p>Alpha beta Alpha</p>');
  const modes: string[] = [];

  try {
    render(
      <DocumentFindReplacePanel
        editor={editor}
        mode="replace"
        onModeChange={(mode) => modes.push(mode)}
        onReplaceText={(from, to, replacement) =>
          editor
            .chain()
            .setTextSelection({ from, to })
            .insertContent(replacement)
            .run()
        }
        onClose={() => undefined}
      />,
    );

    fireEvent.change(screen.getByRole('textbox', { name: '查找内容' }), {
      target: { value: 'Alpha' },
    });
    fireEvent.change(screen.getByRole('textbox', { name: '替换为' }), {
      target: { value: 'Omega' },
    });
    fireEvent.click(screen.getByRole('button', { name: '替换' }));
    expect(editor.getText()).toBe('Omega beta Alpha');
    expect(screen.getByText('已替换当前匹配')).toBeVisible();

    fireEvent.click(screen.getByRole('tab', { name: '查找' }));
    expect(modes).toEqual(['find']);
  } finally {
    editor.destroy();
    element.remove();
  }
});

function createEditor(content: string): {
  editor: Editor;
  element: HTMLDivElement;
} {
  const element = document.createElement('div');
  document.body.append(element);
  return {
    editor: new Editor({
      element,
      extensions: createWorkDocumentExtensions(),
      content,
    }),
    element,
  };
}
