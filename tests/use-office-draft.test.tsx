import { expect, test } from '@rstest/core';
import { act, renderHook } from '@testing-library/react';
import { useOfficeDraft } from '../src/internal/features/work/editors/use-office-draft';

test('preserves a dirty office draft while syncing its saved baseline', () => {
  const { result } = renderHook(() =>
    useOfficeDraft(() => ({ name: '已保存' })),
  );

  expect(result.current.dirty).toBe(false);
  act(() => result.current.setDraft({ name: '编辑中' }));
  expect(result.current.dirty).toBe(true);

  act(() => result.current.syncDraft({ name: '外部更新' }));
  expect(result.current.draft.name).toBe('编辑中');
  expect(result.current.dirty).toBe(true);

  act(() => result.current.cancelDraft());
  expect(result.current.draft.name).toBe('外部更新');
  expect(result.current.dirty).toBe(false);
});

test('can replace or force-sync an office draft as a clean value', () => {
  const { result } = renderHook(() => useOfficeDraft(() => 1));

  act(() => result.current.setDraft(2));
  act(() => result.current.syncDraft(3, true));
  expect(result.current.draft).toBe(3);
  expect(result.current.dirty).toBe(false);

  act(() => result.current.replaceDraft(4));
  expect(result.current.draft).toBe(4);
  expect(result.current.dirty).toBe(false);
});
