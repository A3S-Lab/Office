import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { useState } from 'react';
import { DocumentColumnsPanel } from '../src/internal/features/work/editors/document-columns-panel';
import type { WorkDocumentColumns } from '../src/internal/features/work/work-types';

test('keeps incomplete column values local until Enter or blur', () => {
  const changes: WorkDocumentColumns[] = [];

  function Fixture() {
    const [columns, setColumns] = useState<WorkDocumentColumns>({
      count: 2,
      spacing: 12,
      separator: false,
    });
    return (
      <DocumentColumnsPanel
        columns={columns}
        onChange={(next) => {
          changes.push(next);
          setColumns(next);
        }}
      />
    );
  }

  render(<Fixture />);
  const count = screen.getByRole('textbox', { name: '分栏数量' });
  fireEvent.change(count, { target: { value: '' } });
  expect(count).toHaveValue('');
  expect(changes).toEqual([]);

  fireEvent.change(count, { target: { value: '3.' } });
  expect(changes).toEqual([]);
  fireEvent.keyDown(count, { key: 'Enter' });
  expect(count).toHaveValue('3');
  expect(changes.at(-1)?.count).toBe(3);

  const spacing = screen.getByRole('textbox', { name: '分栏间距' });
  fireEvent.change(spacing, { target: { value: '8.' } });
  fireEvent.keyDown(spacing, { key: 'Escape' });
  expect(spacing).toHaveValue('12');
  expect(changes).toHaveLength(1);
});

test('commits custom column widths as one normalized document change', () => {
  const changes: WorkDocumentColumns[] = [];

  function Fixture() {
    const [columns, setColumns] = useState<WorkDocumentColumns>({
      count: 2,
      spacing: 12,
      separator: false,
      custom: [
        { widthPercent: 50, spacing: 12 },
        { widthPercent: 50, spacing: 0 },
      ],
    });
    return (
      <DocumentColumnsPanel
        columns={columns}
        onChange={(next) => {
          changes.push(next);
          setColumns(next);
        }}
      />
    );
  }

  render(<Fixture />);
  const width = screen.getByRole('textbox', {
    name: '第 1 栏宽度百分比',
  });
  fireEvent.change(width, { target: { value: '' } });
  expect(changes).toEqual([]);
  fireEvent.blur(width);
  expect(width).toHaveValue('50');
  expect(changes).toEqual([]);

  fireEvent.change(width, { target: { value: '62.5' } });
  fireEvent.blur(width);
  expect(changes).toHaveLength(1);
  expect(changes[0].custom?.map((column) => column.widthPercent)).toEqual([
    62.5, 37.5,
  ]);
});
