import { expect, test } from '@rstest/core';
import { act, renderHook } from '@testing-library/react';
import type { OfficeEditorRuntime } from '../src/internal/features/work/editors/office-editor-extension';
import { useOfficeEditorKeyboardShortcuts } from '../src/internal/features/work/editors/use-office-editor-keyboard-shortcuts';

test('runs post-command behavior only for handled editor shortcuts', () => {
  const handled: string[] = [];
  const editor = {
    handleKeyDown: (event: KeyboardEvent) => event.key === 'v',
  } as OfficeEditorRuntime<unknown, unknown>;

  renderHook(() =>
    useOfficeEditorKeyboardShortcuts(editor, {
      onHandled: (event) => handled.push(event.key),
    }),
  );

  act(() => {
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'c' }));
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'v' }));
  });

  expect(handled).toEqual(['v']);
});

test('does not route host-page shortcuts into a scoped editor', () => {
  const routed: string[] = [];
  const editor = {
    handleKeyDown: (event: KeyboardEvent) => {
      routed.push((event.target as HTMLElement).dataset.location ?? 'unknown');
      return true;
    },
  } as OfficeEditorRuntime<unknown, unknown>;
  const scope = document.createElement('section');
  const inside = document.createElement('button');
  const outside = document.createElement('button');
  inside.dataset.location = 'inside';
  outside.dataset.location = 'outside';
  scope.append(inside);
  document.body.append(scope, outside);

  try {
    renderHook(() =>
      useOfficeEditorKeyboardShortcuts(editor, {
        scopeRef: { current: scope },
      }),
    );
    act(() => {
      inside.dispatchEvent(
        new KeyboardEvent('keydown', { bubbles: true, key: 'z' }),
      );
      outside.dispatchEvent(
        new KeyboardEvent('keydown', { bubbles: true, key: 'z' }),
      );
    });

    expect(routed).toEqual(['inside']);
  } finally {
    scope.remove();
    outside.remove();
  }
});
