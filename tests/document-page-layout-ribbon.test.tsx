import { afterEach, expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { Editor } from '@tiptap/core';
import type { DocumentLayoutPanelTab } from '../src/internal/features/work/editors/document-layout-panel';
import { DocumentPageLayoutRibbon } from '../src/internal/features/work/editors/document-page-layout-ribbon';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import type { WorkDocumentSectionLayout } from '../src/internal/features/work/work-types';

const layout: WorkDocumentSectionLayout = {
  pageSize: 'a4',
  orientation: 'portrait',
  margins: { top: 25, right: 23, bottom: 25, left: 23 },
  columns: { count: 1, spacing: 12, separator: false },
  breakAfter: 'nextPage',
};

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

test('executes WPS page-setup presets and routes advanced settings', async () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>Layout target</p>',
  });
  const changes: WorkDocumentSectionLayout[] = [];
  const advancedTargets: DocumentLayoutPanelTab[] = [];
  let layoutToggles = 0;
  let sections = 0;
  render(
    <DocumentPageLayoutRibbon
      editor={editor}
      layout={layout}
      layoutOpen={false}
      pageColor="#ffffff"
      onLayoutChange={(next) => changes.push(next)}
      onOpenLayout={(target) => advancedTargets.push(target)}
      onToggleLayout={() => {
        layoutToggles += 1;
      }}
      onPageColorChange={() => undefined}
      onInsertSection={() => {
        sections += 1;
      }}
    />,
  );

  fireEvent.click(screen.getByRole('combobox', { name: '页边距' }));
  fireEvent.click(await screen.findByRole('option', { name: '普通' }));
  expect(changes.at(-1)?.margins).toEqual({
    top: 25.4,
    right: 25.4,
    bottom: 25.4,
    left: 25.4,
  });

  fireEvent.click(screen.getByRole('combobox', { name: '页边距' }));
  fireEvent.click(await screen.findByRole('option', { name: '窄' }));
  expect(changes.at(-1)?.margins).toEqual({
    top: 12.7,
    right: 12.7,
    bottom: 12.7,
    left: 12.7,
  });

  fireEvent.click(screen.getByRole('combobox', { name: '页边距' }));
  fireEvent.click(await screen.findByRole('option', { name: '宽' }));
  expect(changes.at(-1)?.margins).toEqual({
    top: 25.4,
    right: 50.8,
    bottom: 25.4,
    left: 50.8,
  });

  fireEvent.click(screen.getByRole('combobox', { name: '页面方向' }));
  fireEvent.click(await screen.findByRole('option', { name: '横向' }));
  expect(changes.at(-1)?.orientation).toBe('landscape');

  fireEvent.click(screen.getByRole('combobox', { name: '纸张大小' }));
  fireEvent.click(await screen.findByRole('option', { name: 'Letter' }));
  expect(changes.at(-1)?.pageSize).toBe('letter');

  fireEvent.click(screen.getByRole('combobox', { name: '分栏' }));
  fireEvent.click(await screen.findByRole('option', { name: '两栏' }));
  expect(changes.at(-1)?.columns).toMatchObject({ count: 2 });

  fireEvent.click(screen.getByRole('combobox', { name: '页边距' }));
  fireEvent.click(await screen.findByRole('option', { name: '自定义页边距' }));
  fireEvent.click(screen.getByRole('combobox', { name: '分栏' }));
  fireEvent.click(await screen.findByRole('option', { name: '更多分栏' }));
  expect(advancedTargets).toEqual(['page', 'columns']);

  fireEvent.click(screen.getByRole('button', { name: '页面设置' }));
  fireEvent.click(screen.getByRole('button', { name: '插入分节符' }));
  expect(layoutToggles).toBe(1);
  expect(sections).toBe(1);
});
