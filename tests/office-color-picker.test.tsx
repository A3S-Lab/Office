import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { OfficeColorPicker } from '../src/internal/features/work/editors/office-color-picker';

test('opens on the current color and supports spatial keyboard navigation', async () => {
  const selected: string[] = [];
  render(
    <OfficeColorPicker
      ariaLabel="文字颜色"
      value="#111827"
      onValueChange={(value) => selected.push(value)}
    />,
  );

  const trigger = screen.getByRole('button', { name: '文字颜色' });
  fireEvent.click(trigger);
  const first = screen.getByRole('option', { name: '颜色 #111827' });
  const second = screen.getByRole('option', { name: '颜色 #374151' });
  await waitFor(() => expect(first).toHaveFocus());

  fireEvent.keyDown(first, { key: 'ArrowRight' });
  expect(second).toHaveFocus();
  fireEvent.click(second);

  expect(selected).toEqual(['#374151']);
  expect(screen.queryByRole('dialog', { name: '文字颜色' })).toBeNull();
  expect(trigger).toHaveFocus();
});
