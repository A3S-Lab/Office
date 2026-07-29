import { Editor } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import {
  DocumentFindReplacePanel,
  documentTextMatches,
} from '../src/internal/features/work/editors/document-find-replace-panel';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

test('finds text across adjacent marked runs and moves through matches', async () => {
  const { editor, element } = createEditor(
    '<p>Al<strong>pha</strong> beta Alpha</p>',
  );

  try {
    expect(documentTextMatches(editor, 'alpha')).toEqual([
      { from: 1, to: 6 },
      { from: 12, to: 17 },
    ]);
    const view = render(
      <DocumentFindReplacePanel
        editor={editor}
        mode="find"
        onModeChange={() => undefined}
        onReplaceText={() => false}
        onClose={() => undefined}
      />,
    );

    const query = screen.getByRole('textbox', { name: '查找内容' });
    fireEvent.change(query, {
      target: { value: 'alpha' },
    });
    expect(screen.getByText('2 个匹配')).toBeVisible();
    await waitFor(() =>
      expect(
        new Set(
          [
            ...editor.view.dom.querySelectorAll<HTMLElement>(
              '.work-document-find-match',
            ),
          ].map((match) => match.dataset.documentFindIndex),
        ).size,
      ).toBe(2),
    );
    expect(
      editor.view.dom.querySelectorAll('.work-document-find-match.active'),
    ).toHaveLength(0);

    query.focus();
    fireEvent.keyDown(query, { key: 'Enter' });
    expect(query).toHaveFocus();
    expect(editor.state.selection).toMatchObject({ from: 1, to: 6 });
    expect(
      new Set(
        [
          ...editor.view.dom.querySelectorAll<HTMLElement>(
            '.work-document-find-match.active',
          ),
        ].map((match) => match.dataset.documentFindIndex),
      ),
    ).toEqual(new Set(['0']));

    fireEvent.keyDown(query, { key: 'Enter' });
    expect(query).toHaveFocus();
    expect(editor.state.selection).toMatchObject({ from: 12, to: 17 });
    expect(
      new Set(
        [
          ...editor.view.dom.querySelectorAll<HTMLElement>(
            '.work-document-find-match.active',
          ),
        ].map((match) => match.dataset.documentFindIndex),
      ),
    ).toEqual(new Set(['1']));
    expect(
      editor.state.doc.textBetween(
        editor.state.selection.from,
        editor.state.selection.to,
      ),
    ).toBe('Alpha');
    view.unmount();
    expect(
      editor.view.dom.querySelectorAll('.work-document-find-match'),
    ).toHaveLength(0);
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
    const replacement = screen.getByRole('textbox', { name: '替换为' });
    fireEvent.change(replacement, {
      target: { value: 'Omega' },
    });
    fireEvent.click(screen.getByRole('button', { name: '替换' }));
    expect(editor.getText()).toBe('Omega beta Alpha');
    expect(screen.getByText('已替换当前匹配')).toBeVisible();

    replacement.focus();
    fireEvent.keyDown(replacement, { key: 'Enter' });
    expect(editor.getText()).toBe('Omega beta Omega');
    expect(replacement).toHaveFocus();
    expect(screen.getByText('没有匹配内容')).toBeVisible();

    fireEvent.click(screen.getByRole('tab', { name: '查找' }));
    expect(modes).toEqual(['find']);
  } finally {
    editor.destroy();
    element.remove();
  }
});

test('moves backward from an inactive search to the final match', () => {
  const { editor, element } = createEditor('<p>Alpha beta Alpha</p>');

  try {
    render(
      <DocumentFindReplacePanel
        editor={editor}
        mode="find"
        onModeChange={() => undefined}
        onReplaceText={() => false}
        onClose={() => undefined}
      />,
    );

    const query = screen.getByRole('textbox', { name: '查找内容' });
    fireEvent.change(query, { target: { value: 'Alpha' } });
    fireEvent.keyDown(query, { key: 'Enter', shiftKey: true });

    expect(query).toHaveFocus();
    expect(editor.state.selection).toMatchObject({ from: 12, to: 17 });
    expect(screen.getByText('第 2 个，共 2 个')).toBeVisible();
  } finally {
    editor.destroy();
    element.remove();
  }
});

test('refocuses and selects the query when the command is invoked again', async () => {
  const { editor, element } = createEditor('<p>Alpha beta Alpha</p>');

  try {
    const view = render(
      <>
        <button type="button">Outside</button>
        <DocumentFindReplacePanel
          editor={editor}
          mode="find"
          focusRequest={0}
          onModeChange={() => undefined}
          onReplaceText={() => false}
          onClose={() => undefined}
        />
      </>,
    );
    const query = screen.getByRole('textbox', { name: '查找内容' });
    fireEvent.change(query, { target: { value: 'Alpha' } });
    screen.getByRole('button', { name: 'Outside' }).focus();
    expect(query).not.toHaveFocus();

    view.rerender(
      <>
        <button type="button">Outside</button>
        <DocumentFindReplacePanel
          editor={editor}
          mode="find"
          focusRequest={1}
          onModeChange={() => undefined}
          onReplaceText={() => false}
          onClose={() => undefined}
        />
      </>,
    );

    await waitFor(() => expect(query).toHaveFocus());
    expect(query.selectionStart).toBe(0);
    expect(query.selectionEnd).toBe(query.value.length);
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
