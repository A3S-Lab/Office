import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useState } from 'react';
import { OfficeSelect } from '../src/internal/features/work/editors/office-select';

test('opens and moves to a matching option from printable keyboard input', async () => {
  function Fixture() {
    const [value, setValue] = useState('default');
    return (
      <OfficeSelect
        ariaLabel="字体"
        value={value}
        options={[
          { value: 'default', label: '默认字体' },
          { value: 'simsun', label: '宋体' },
          { value: 'arial', label: 'Arial' },
          { value: 'aptos', label: 'Aptos', disabled: true },
        ]}
        onValueChange={setValue}
      />
    );
  }

  render(<Fixture />);
  const trigger = screen.getByRole('combobox', { name: '字体' });
  fireEvent.keyDown(trigger, { key: 'a' });
  await waitFor(() =>
    expect(screen.getByRole('option', { name: 'Arial' })).toHaveFocus(),
  );

  fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'Enter' });
  expect(trigger).toHaveTextContent('Arial');
  expect(trigger).toHaveFocus();

  fireEvent.keyDown(trigger, { key: 'a' });
  await waitFor(() =>
    expect(screen.getByRole('option', { name: 'Arial' })).toHaveFocus(),
  );
  fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'a' });
  expect(screen.getByRole('option', { name: 'Arial' })).toHaveFocus();
});

test('groups long option lists and searches their alternate names', async () => {
  render(
    <OfficeSelect
      ariaLabel="字体"
      value="default"
      options={[
        { value: 'default', label: '默认字体' },
        {
          value: 'pingfang',
          group: '中文字体',
          label: '苹方',
          meta: '系统',
          searchText: 'PingFang SC 苹方',
        },
        {
          value: 'arial',
          group: '西文字体',
          label: 'Arial',
          meta: '系统',
        },
      ]}
      onValueChange={() => undefined}
    />,
  );

  const trigger = screen.getByRole('combobox', { name: '字体' });
  fireEvent.keyDown(trigger, { key: 'p' });

  await waitFor(() =>
    expect(screen.getByRole('option', { name: '苹方' })).toHaveFocus(),
  );
  expect(screen.getByText('中文字体')).toHaveClass(
    'work-office-select-group-label',
  );
  expect(screen.getByText('西文字体')).toHaveClass(
    'work-office-select-group-label',
  );
  expect(screen.getAllByText('系统')).toHaveLength(2);
});
