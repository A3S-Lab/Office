import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useState } from 'react';
import {
  CommittedOfficeTextArea,
  CommittedOfficeTextField,
  OfficeFileInput,
} from '../src/internal/features/work/editors/office-text-field';

test('starts file consumers before reset and supports same-file reselection', async () => {
  const file = new File(['image'], 'diagram.png', { type: 'image/png' });
  const finishReadings: Array<() => void> = [];
  const selected: File[] = [];

  render(
    <OfficeFileInput
      aria-label="测试图片文件"
      onFileSelect={(nextFile) => {
        selected.push(nextFile);
        return new Promise<void>((resolve) => {
          finishReadings.push(resolve);
        });
      }}
    />,
  );

  const input = screen.getByLabelText('测试图片文件') as HTMLInputElement;
  Object.defineProperty(input, 'files', {
    configurable: true,
    value: [file],
  });
  let inputValue = 'C:\\fakepath\\diagram.png';
  const releasedValues: string[] = [];
  Object.defineProperty(input, 'value', {
    configurable: true,
    get: () => inputValue,
    set: (value: string) => {
      inputValue = value;
      releasedValues.push(value);
    },
  });

  fireEvent.change(input);

  expect(selected).toEqual([file]);
  expect(inputValue).toBe('');
  expect(releasedValues).toEqual(['']);

  inputValue = 'C:\\fakepath\\diagram.png';
  fireEvent.change(input);

  expect(selected).toEqual([file, file]);
  expect(inputValue).toBe('');
  expect(releasedValues).toEqual(['', '']);

  for (const finishReading of finishReadings) finishReading();
  await waitFor(() => expect(inputValue).toBe(''));
  expect(releasedValues).toEqual(['', '']);
});

test('keeps multiline drafts local until they are valid and committed', () => {
  const commits: number[][] = [];
  let escapes = 0;

  function Fixture() {
    const [values, setValues] = useState([1, 2]);
    return (
      <fieldset
        aria-label="Committed textarea boundary"
        onKeyDown={() => escapes++}
      >
        <CommittedOfficeTextArea
          aria-label="系列数据"
          value={values}
          formatValue={(items) => items.join(', ')}
          parseValue={(draft) => {
            const items = draft
              .split(/[\s,]+/)
              .filter(Boolean)
              .map(Number);
            return items.every(Number.isFinite) ? items : null;
          }}
          onValueCommit={(items) => {
            commits.push(items);
            setValues(items);
          }}
        />
      </fieldset>
    );
  }

  render(<Fixture />);
  const input = screen.getByRole('textbox', { name: '系列数据' });

  fireEvent.change(input, { target: { value: '1.5, 2.' } });
  expect(input).toHaveValue('1.5, 2.');
  expect(commits).toEqual([]);
  fireEvent.blur(input);
  expect(input).toHaveValue('1.5, 2');
  expect(commits).toEqual([[1.5, 2]]);

  fireEvent.change(input, { target: { value: '1.5, invalid' } });
  expect(input).toHaveAttribute('aria-invalid', 'true');
  expect(input).toHaveAttribute('data-office-escape-consumer', 'true');
  fireEvent.keyDown(input, { key: 'Escape' });
  expect(input).toHaveValue('1.5, 2');
  expect(input).not.toHaveAttribute('data-office-escape-consumer');
  expect(escapes).toBe(0);
  expect(commits).toEqual([[1.5, 2]]);

  fireEvent.change(input, { target: { value: '3\n4' } });
  fireEvent.keyDown(input, { key: 'Enter', metaKey: true });
  expect(input).toHaveValue('3, 4');
  expect(commits).toEqual([
    [1.5, 2],
    [3, 4],
  ]);
});

test('commits single-line drafts with Enter and restores them with Escape', () => {
  const commits: number[][] = [];

  function Fixture() {
    const [values, setValues] = useState([1, 2]);
    return (
      <CommittedOfficeTextField
        aria-label="误差值"
        value={values}
        formatValue={(items) => items.join(', ')}
        parseValue={(draft) => {
          const items = draft.split(',').map((item) => Number(item.trim()));
          return items.every(Number.isFinite) ? items : null;
        }}
        onValueCommit={(items) => {
          commits.push(items);
          setValues(items);
        }}
      />
    );
  }

  render(<Fixture />);
  const input = screen.getByRole('textbox', { name: '误差值' });
  fireEvent.change(input, { target: { value: '1, invalid' } });
  expect(input).toHaveAttribute('aria-invalid', 'true');
  fireEvent.keyDown(input, { key: 'Escape' });
  expect(input).toHaveValue('1, 2');
  expect(commits).toEqual([]);

  fireEvent.change(input, { target: { value: '2.5, 3.' } });
  fireEvent.keyDown(input, { key: 'Enter' });
  expect(input).toHaveValue('2.5, 3');
  expect(commits).toEqual([[2.5, 3]]);
});
