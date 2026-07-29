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
