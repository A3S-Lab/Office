import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
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
