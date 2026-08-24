import type { PluginRegistry } from '@embedpdf/react-pdf-viewer';
import { expect, rstest, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import type {
  PdfEditorCanCommands,
  PdfEditorCommands,
} from '../src/internal/features/work/editors/pdf-editor-extensions';
import { PdfPageOrganizerDialog } from '../src/internal/features/work/editors/pdf-page-organizer-dialog';

test('selects pages and routes one rotate intent through the typed command', async () => {
  const calls: string[] = [];
  const onClose = rstest.fn();
  render(
    <PdfPageOrganizerDialog
      busy={false}
      can={canCommands()}
      commands={commands(calls)}
      currentPage={2}
      diagnostics={[]}
      error={null}
      registry={emptyRegistry()}
      restoreFocusTarget={() => null}
      totalPages={4}
      onClose={onClose}
      onDismissError={() => undefined}
    />,
  );

  expect(screen.getByRole('dialog', { name: '组织 PDF 页面' })).toBeTruthy();
  expect(screen.getByText('已选择 1 / 4 页')).toBeTruthy();
  fireEvent.click(screen.getByRole('button', { name: '选择第 3 页' }), {
    metaKey: true,
  });
  expect(screen.getByText('已选择 2 / 4 页')).toBeTruthy();
  fireEvent.click(screen.getByRole('button', { name: '向右旋转所选页' }));
  await waitFor(() => expect(calls).toEqual(['rotate:1,2:90']));
  expect(onClose).toHaveBeenCalledTimes(1);
});

test('keeps delete, reorder, extract, split, insert, and merge discoverable', () => {
  render(
    <PdfPageOrganizerDialog
      busy={false}
      can={canCommands()}
      commands={commands([])}
      currentPage={1}
      diagnostics={[]}
      error={null}
      registry={emptyRegistry()}
      restoreFocusTarget={() => null}
      totalPages={3}
      onClose={() => undefined}
      onDismissError={() => undefined}
    />,
  );

  for (const name of [
    '插入空白页',
    '合并另一个 PDF',
    '前移所选页',
    '后移所选页',
    '向左旋转所选页',
    '向右旋转所选页',
    '抽取所选页',
    '在所选页后拆分',
    '删除所选页',
  ]) {
    expect(screen.getByRole('button', { name })).toBeTruthy();
  }
  expect(screen.getByText(/签名、加密、表单、目录或标记结构/)).toBeTruthy();
});

test('shows stable fail-closed diagnostics without closing the dialog', () => {
  const onDismissError = rstest.fn();
  render(
    <PdfPageOrganizerDialog
      busy={false}
      can={canCommands()}
      commands={commands([])}
      currentPage={1}
      diagnostics={[]}
      error={{
        code: 'pdf.pages.signed-source',
        message: 'Rewriting this PDF would invalidate its signature.',
      }}
      registry={emptyRegistry()}
      restoreFocusTarget={() => null}
      totalPages={2}
      onClose={() => undefined}
      onDismissError={onDismissError}
    />,
  );
  expect(screen.getByRole('alert')).toHaveTextContent(
    'Rewriting this PDF would invalidate its signature.',
  );
  fireEvent.click(screen.getByRole('button', { name: '关闭错误提示' }));
  expect(onDismissError).toHaveBeenCalledTimes(1);
});

function commands(calls: string[]): PdfEditorCommands {
  return {
    deletePages: async (indexes) => {
      calls.push(`delete:${indexes.join(',')}`);
      return true;
    },
    extractPages: async (indexes) => {
      calls.push(`extract:${indexes.join(',')}`);
      return true;
    },
    insertBlankPage: async (index) => {
      calls.push(`insert:${index}`);
      return true;
    },
    mergePages: async (index, source) => {
      calls.push(`merge:${index}:${source.size}`);
      return true;
    },
    openPageOrganizer: () => undefined,
    reorderPages: async (order) => {
      calls.push(`reorder:${order.join(',')}`);
      return true;
    },
    rotatePages: async (indexes, degrees) => {
      calls.push(`rotate:${indexes.join(',')}:${degrees}`);
      return true;
    },
    splitPages: async (indexes) => {
      calls.push(`split:${indexes.join(',')}`);
      return true;
    },
  } as PdfEditorCommands;
}

function canCommands(): PdfEditorCanCommands {
  const allowed = () => true;
  return {
    deletePages: allowed,
    extractPages: allowed,
    insertBlankPage: allowed,
    mergePages: allowed,
    openPageOrganizer: allowed,
    reorderPages: allowed,
    rotatePages: allowed,
    splitPages: allowed,
  } as PdfEditorCanCommands;
}

function emptyRegistry(): PluginRegistry {
  return {
    getPlugin: () => undefined,
    pluginsReady: () => Promise.resolve(),
  } as unknown as PluginRegistry;
}
