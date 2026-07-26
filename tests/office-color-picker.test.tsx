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
  expect(screen.getByText('主题颜色')).toBeVisible();
  expect(screen.getByText('标准色')).toBeVisible();
  const first = screen.getByRole('option', { name: '颜色 #111827' });
  const second = screen.getByRole('option', { name: '颜色 #1f2937' });
  await waitFor(() => expect(first).toHaveFocus());

  fireEvent.keyDown(first, { key: 'ArrowRight' });
  expect(second).toHaveFocus();
  fireEvent.click(second);

  expect(selected).toEqual(['#1f2937']);
  expect(screen.queryByRole('dialog', { name: '文字颜色' })).toBeNull();
  expect(trigger).toHaveFocus();
});

test('shows a readable custom-color preview before applying a valid value', () => {
  const selected: string[] = [];
  render(
    <OfficeColorPicker
      ariaLabel="页面颜色"
      value="#ffffff"
      triggerLabel="页面颜色"
      onValueChange={(value) => selected.push(value)}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '页面颜色' }));
  const input = screen.getByRole('textbox', { name: '自定义颜色值' });
  const apply = screen.getByRole('button', { name: '应用自定义颜色' });
  fireEvent.change(input, { target: { value: '#fff2cc' } });

  expect(screen.getByTestId('custom-color-preview')).toHaveStyle({
    backgroundColor: '#fff2cc',
  });
  fireEvent.click(apply);
  expect(selected).toEqual(['#fff2cc']);
});
