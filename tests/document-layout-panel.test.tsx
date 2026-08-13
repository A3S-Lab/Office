import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { useState } from 'react';
import { DocumentLayoutPanel } from '../src/internal/features/work/editors/document-layout-panel';
import type { WorkDocumentSectionLayout } from '../src/internal/features/work/work-types';

const layout: WorkDocumentSectionLayout = {
  pageSize: 'a4',
  orientation: 'portrait',
  margins: { top: 25, right: 23, bottom: 25, left: 23 },
  columns: { count: 1, spacing: 12, separator: false },
  breakAfter: 'nextPage',
  pageNumberStart: 1,
};

test('separates page, columns, and header/footer settings into focused tabs', async () => {
  render(
    <DocumentLayoutPanel
      layout={layout}
      sectionIndex={0}
      sectionCount={2}
      onChange={() => undefined}
      onInsertSection={() => undefined}
      onMergeSection={() => undefined}
      onClose={() => undefined}
    />,
  );

  const pageTab = screen.getByRole('tab', { name: '页面' });
  const columnsTab = screen.getByRole('tab', { name: '分栏与分节' });
  const headerFooterTab = screen.getByRole('tab', { name: '页眉页脚' });

  expect(pageTab).toHaveAttribute('aria-selected', 'true');
  expect(screen.getByRole('combobox', { name: '纸张大小' })).toBeVisible();
  expect(screen.getByLabelText('上页边距')).toBeVisible();
  expect(screen.queryByLabelText('分栏数量')).toBeNull();
  expect(screen.queryByRole('textbox', { name: '默认页页眉' })).toBeNull();

  fireEvent.click(columnsTab);
  expect(columnsTab).toHaveAttribute('aria-selected', 'true');
  expect(screen.getByLabelText('分栏数量')).toBeVisible();
  expect(screen.getByRole('combobox', { name: '分节方式' })).toBeVisible();
  expect(screen.queryByRole('combobox', { name: '纸张大小' })).toBeNull();

  fireEvent.keyDown(columnsTab, { key: 'ArrowRight' });
  expect(headerFooterTab).toHaveAttribute('aria-selected', 'true');
  expect(
    await screen.findByRole('textbox', { name: '默认页页眉' }),
  ).toBeVisible();
  expect(screen.getByRole('textbox', { name: '默认页页脚' })).toBeVisible();
  expect(screen.getByLabelText('起始页码')).toBeVisible();
});

test('keeps page changes controlled inside the active tab', () => {
  const changes: WorkDocumentSectionLayout[] = [];
  render(
    <DocumentLayoutPanel
      layout={layout}
      sectionIndex={1}
      sectionCount={2}
      onChange={(nextLayout) => changes.push(nextLayout)}
      onInsertSection={() => undefined}
      onMergeSection={() => undefined}
      onClose={() => undefined}
    />,
  );

  fireEvent.click(screen.getByRole('combobox', { name: '页面方向' }));
  fireEvent.click(screen.getByRole('option', { name: '横向' }));

  expect(changes).toEqual([{ ...layout, orientation: 'landscape' }]);
});

test('commits page measurements only after a complete value', () => {
  const changes: WorkDocumentSectionLayout[] = [];

  function Fixture() {
    const [current, setCurrent] = useState(layout);
    return (
      <DocumentLayoutPanel
        layout={current}
        sectionIndex={0}
        sectionCount={1}
        onChange={(next) => {
          changes.push(next);
          setCurrent(next);
        }}
        onInsertSection={() => undefined}
        onMergeSection={() => undefined}
        onClose={() => undefined}
      />
    );
  }

  render(<Fixture />);
  const topMargin = screen.getByRole('textbox', { name: '上页边距' });
  fireEvent.change(topMargin, { target: { value: '' } });
  expect(topMargin).toHaveValue('');
  expect(changes).toEqual([]);

  fireEvent.blur(topMargin);
  expect(topMargin).toHaveValue('25');
  expect(changes).toEqual([]);

  fireEvent.change(topMargin, { target: { value: '42.' } });
  expect(changes).toEqual([]);
  fireEvent.keyDown(topMargin, { key: 'Enter' });
  expect(topMargin).toHaveValue('42');
  expect(changes.at(-1)?.margins.top).toBe(42);

  fireEvent.change(topMargin, { target: { value: '16' } });
  fireEvent.keyDown(topMargin, { key: 'Escape' });
  expect(topMargin).toHaveValue('42');
  expect(changes).toHaveLength(1);
});

test('keeps an incomplete starting page number out of the section model', () => {
  const changes: WorkDocumentSectionLayout[] = [];

  function Fixture() {
    const [current, setCurrent] = useState(layout);
    return (
      <DocumentLayoutPanel
        layout={current}
        sectionIndex={0}
        sectionCount={1}
        onChange={(next) => {
          changes.push(next);
          setCurrent(next);
        }}
        onInsertSection={() => undefined}
        onMergeSection={() => undefined}
        onClose={() => undefined}
      />
    );
  }

  render(<Fixture />);
  fireEvent.click(screen.getByRole('tab', { name: '页眉页脚' }));
  const start = screen.getByRole('textbox', { name: '起始页码' });
  fireEvent.change(start, { target: { value: '' } });
  expect(changes).toEqual([]);
  fireEvent.blur(start);
  expect(start).toHaveValue('1');
  expect(changes).toEqual([]);

  fireEvent.change(start, { target: { value: '12.6' } });
  fireEvent.keyDown(start, { key: 'Enter' });
  expect(start).toHaveValue('13');
  expect(changes.at(-1)?.pageNumberStart).toBe(13);
});

test('edits exact page chrome, overlap, mirror, and gutter geometry', () => {
  const changes: WorkDocumentSectionLayout[] = [];

  function Fixture() {
    const [current, setCurrent] = useState(layout);
    return (
      <DocumentLayoutPanel
        layout={current}
        sectionIndex={0}
        sectionCount={2}
        onChange={(next) => {
          changes.push(next);
          setCurrent(next);
        }}
        onInsertSection={() => undefined}
        onMergeSection={() => undefined}
        onClose={() => undefined}
      />
    );
  }

  render(<Fixture />);
  const headerDistance = screen.getByRole('textbox', {
    name: '页眉距顶端',
  });
  fireEvent.change(headerDistance, { target: { value: '10' } });
  fireEvent.keyDown(headerDistance, { key: 'Enter' });
  expect(changes.at(-1)?.pageMargins?.header).toBe(567);

  fireEvent.click(screen.getByRole('combobox', { name: '上边距与页眉关系' }));
  fireEvent.click(screen.getByRole('option', { name: '允许重叠' }));
  expect(changes.at(-1)?.pageMargins?.top).toBe(-1_417);

  fireEvent.click(screen.getByRole('combobox', { name: '镜像页边距' }));
  fireEvent.click(screen.getByRole('option', { name: '镜像页边距' }));
  expect(changes.at(-1)?.pageMargins).toMatchObject({
    mirrorMargins: true,
    gutterAtTop: false,
  });

  fireEvent.click(screen.getByRole('combobox', { name: '装订线位置' }));
  fireEvent.click(screen.getByRole('option', { name: '顶部' }));
  expect(changes.at(-1)?.pageMargins).toMatchObject({
    mirrorMargins: false,
    gutterAtTop: true,
    gutterOnRight: false,
  });
});
