import { expect, test } from '@rstest/core';
import { renderHook } from '@testing-library/react';
import { usePlaygroundDocumentSuggestionFixture } from '../playground/src/collaboration-suggestion-fixture';

test('recreates the native suggestion fixture after it is disabled', () => {
  const { result, rerender, unmount } = renderHook(
    ({ enabled }: { enabled: boolean }) =>
      usePlaygroundDocumentSuggestionFixture(enabled),
    { initialProps: { enabled: true } },
  );

  expect(result.current?.nativeStage).toBe('ready');
  const firstEditor = result.current?.editor;

  rerender({ enabled: false });
  expect(result.current).toBeUndefined();

  rerender({ enabled: true });
  expect(result.current?.nativeStage).toBe('ready');
  expect(result.current?.editor).not.toBe(firstEditor);

  unmount();
});
