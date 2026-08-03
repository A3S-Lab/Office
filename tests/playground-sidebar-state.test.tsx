import { expect, test } from '@rstest/core';
import { act, renderHook } from '@testing-library/react';
import { usePlaygroundSidebarState } from '../playground/src/use-playground-sidebar-state';

test('keeps persistent and modal sidebar preferences independent across breakpoints', () => {
  const { result, rerender } = renderHook(
    ({ modal }: { modal: boolean }) => usePlaygroundSidebarState(modal),
    { initialProps: { modal: false } },
  );

  expect(result.current.open).toBe(true);

  rerender({ modal: true });
  expect(result.current.open).toBe(false);

  act(() => result.current.setOpen(true));
  expect(result.current.open).toBe(true);

  rerender({ modal: false });
  expect(result.current.open).toBe(true);

  rerender({ modal: true });
  expect(result.current.open).toBe(false);

  rerender({ modal: false });
  act(() => result.current.setOpen(false));
  expect(result.current.open).toBe(false);

  rerender({ modal: true });
  expect(result.current.open).toBe(false);

  act(() => result.current.setOpen(true));
  expect(result.current.open).toBe(true);
  act(() => result.current.closeAll());
  expect(result.current.open).toBe(false);

  rerender({ modal: false });
  expect(result.current.open).toBe(false);
});
