import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { useState } from 'react';
import {
  CommittedOfficeNumberField,
  OfficeNumberField,
} from '../src/internal/features/work/editors/office-number-field';

test('commits typed values only on Enter or blur', () => {
  const commits: string[] = [];

  function Fixture() {
    const [value, setValue] = useState('1');
    return (
      <OfficeNumberField
        ariaLabel="段前间距"
        value={value}
        step={0.5}
        onValueChange={setValue}
        onCommit={(next) => commits.push(next)}
      />
    );
  }

  render(<Fixture />);
  const input = screen.getByRole('textbox', { name: '段前间距' });
  fireEvent.change(input, { target: { value: '12.' } });
  expect(input).toHaveValue('12.');
  expect(commits).toEqual([]);

  fireEvent.keyDown(input, { key: 'Enter', isComposing: true });
  expect(commits).toEqual([]);

  fireEvent.keyDown(input, { key: 'Enter' });
  expect(commits).toEqual(['12.']);

  fireEvent.change(input, { target: { value: '12.5' } });
  fireEvent.blur(input);
  expect(commits).toEqual(['12.', '12.5']);
});

test('commits keyboard and button step changes immediately', () => {
  const commits: string[] = [];

  function Fixture() {
    const [value, setValue] = useState('1');
    return (
      <OfficeNumberField
        ariaLabel="列宽"
        value={value}
        min={0.5}
        step={0.5}
        onValueChange={setValue}
        onCommit={(next) => commits.push(next)}
      />
    );
  }

  render(<Fixture />);
  const input = screen.getByRole('textbox', { name: '列宽' });
  fireEvent.keyDown(input, { key: 'ArrowUp' });
  expect(input).toHaveValue('1.5');
  fireEvent.click(screen.getByRole('button', { name: '减少列宽' }));
  expect(input).toHaveValue('1');
  expect(commits).toEqual(['1.5', '1']);
});

test('reverts a draft with Escape without leaking the key to its menu', () => {
  const cancels: string[] = [];
  render(
    <fieldset
      aria-label="Number field shortcut boundary"
      onKeyDown={() => cancels.push('parent')}
    >
      <OfficeNumberField
        ariaLabel="缩放比例"
        value="125."
        escapeConsumer
        onValueChange={() => undefined}
        onCancel={() => cancels.push('cancel')}
      />
    </fieldset>,
  );

  fireEvent.keyDown(screen.getByRole('textbox', { name: '缩放比例' }), {
    key: 'Escape',
  });

  expect(cancels).toEqual(['cancel']);
});

test('lets a clean committed field pass Escape to its containing surface', () => {
  const calls: string[] = [];

  function Fixture() {
    const [value, setValue] = useState(25);
    return (
      <fieldset
        aria-label="Task pane shortcut boundary"
        onKeyDown={() => calls.push('surface')}
      >
        <CommittedOfficeNumberField
          ariaLabel="上页边距"
          value={value}
          min={5}
          max={60}
          normalizeValue={(rawValue) => {
            const number = Number(rawValue);
            return Number.isFinite(number) ? number : null;
          }}
          onValueCommit={setValue}
        />
      </fieldset>
    );
  }

  render(<Fixture />);
  const input = screen.getByRole('textbox', { name: '上页边距' });

  fireEvent.keyDown(input, { key: 'Escape' });
  expect(calls).toEqual(['surface']);

  fireEvent.change(input, { target: { value: '32' } });
  fireEvent.keyDown(input, { key: 'Escape' });
  expect(input).toHaveValue('25');
  expect(calls).toEqual(['surface']);
});

test('keeps incomplete committed fields local until a valid commit', () => {
  const commits: number[] = [];

  function Fixture() {
    const [value, setValue] = useState(25);
    return (
      <CommittedOfficeNumberField
        ariaLabel="上页边距"
        value={value}
        min={5}
        max={60}
        step={0.5}
        normalizeValue={(rawValue) => {
          if (!rawValue.trim()) return null;
          const number = Number(rawValue);
          if (!Number.isFinite(number)) return null;
          return Math.min(60, Math.max(5, Math.round(number * 10) / 10));
        }}
        onValueCommit={(nextValue) => {
          commits.push(nextValue);
          setValue(nextValue);
        }}
      />
    );
  }

  render(<Fixture />);
  const input = screen.getByRole('textbox', { name: '上页边距' });
  fireEvent.change(input, { target: { value: '' } });
  expect(input).toHaveValue('');
  expect(input).toHaveAttribute('data-office-escape-consumer', 'true');
  expect(commits).toEqual([]);

  fireEvent.blur(input);
  expect(input).toHaveValue('25');
  expect(input).not.toHaveAttribute('data-office-escape-consumer');
  expect(commits).toEqual([]);

  fireEvent.change(input, { target: { value: '42.' } });
  expect(input).toHaveValue('42.');
  expect(commits).toEqual([]);
  fireEvent.keyDown(input, { key: 'Enter' });
  expect(input).toHaveValue('42');
  expect(commits).toEqual([42]);

  fireEvent.change(input, { target: { value: '18' } });
  fireEvent.keyDown(input, { key: 'Escape' });
  expect(input).toHaveValue('42');
  expect(input).not.toHaveAttribute('data-office-escape-consumer');
  expect(commits).toEqual([42]);
});

test('supports an explicit empty commit for optional measurements', () => {
  const commits: Array<number | undefined> = [];

  function Fixture() {
    const [value, setValue] = useState<number | undefined>(10);
    return (
      <CommittedOfficeNumberField
        ariaLabel="坐标轴最小值"
        value={value}
        placeholder="自动"
        normalizeValue={(rawValue) => {
          if (!rawValue.trim()) return undefined;
          const number = Number(rawValue);
          return Number.isFinite(number) ? number : null;
        }}
        onValueCommit={(nextValue) => {
          commits.push(nextValue);
          setValue(nextValue);
        }}
      />
    );
  }

  render(<Fixture />);
  const input = screen.getByRole('textbox', { name: '坐标轴最小值' });
  fireEvent.change(input, { target: { value: '' } });
  expect(commits).toEqual([]);
  fireEvent.blur(input);
  expect(input).toHaveValue('');
  expect(commits).toEqual([undefined]);
});

test('commits Enter without submitting the surrounding settings form', () => {
  const calls: string[] = [];
  render(
    <form
      onSubmit={(event) => {
        event.preventDefault();
        calls.push('submit');
      }}
    >
      <OfficeNumberField
        ariaLabel="最大迭代次数"
        value="120"
        onValueChange={() => undefined}
        onCommit={() => calls.push('commit')}
      />
    </form>,
  );

  fireEvent.keyDown(screen.getByRole('textbox', { name: '最大迭代次数' }), {
    key: 'Enter',
  });

  expect(calls).toEqual(['commit']);
});

test('does not commit twice when Enter moves focus and synchronously blurs', () => {
  const commits: string[] = [];
  let nextControl: HTMLButtonElement | null = null;

  render(
    <>
      <OfficeNumberField
        ariaLabel="列宽"
        value="3.2"
        onValueChange={() => undefined}
        onCommit={(value) => {
          commits.push(value);
          nextControl?.focus();
        }}
      />
      <button
        ref={(element) => {
          nextControl = element;
        }}
        type="button"
      >
        返回文档
      </button>
    </>,
  );

  const input = screen.getByRole('textbox', { name: '列宽' });
  input.focus();
  fireEvent.keyDown(input, { key: 'Enter' });

  expect(commits).toEqual(['3.2']);
  expect(screen.getByRole('button', { name: '返回文档' })).toHaveFocus();
});
