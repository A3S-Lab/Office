import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { createRef } from 'react';
import { PdfToolbar } from '../src/internal/features/work/editors/pdf-toolbar';
import type { PdfAnnotationController } from '../src/internal/features/work/editors/pdf-annotation-controller';
import type {
  PdfViewerController,
  PdfViewerControllerState,
} from '../src/internal/features/work/editors/pdf-viewer-controller';

test('keeps PDF navigation, search, zoom, history, and save in one toolbar', () => {
  const calls: string[] = [];
  const controller = createController(calls);
  const searchRef = createRef<HTMLInputElement>();

  render(
    <PdfToolbar
      annotation={createAnnotationController(calls)}
      controller={controller}
      editable
      onSave={() => calls.push('save')}
      saveLabel="保存"
      saveState="idle"
      searchInputRef={searchRef}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '撤销' }));
  fireEvent.click(screen.getByRole('button', { name: '高亮' }));
  fireEvent.click(screen.getByRole('button', { name: '下一页' }));
  fireEvent.change(screen.getByRole('textbox', { name: '页码' }), {
    target: { value: '7' },
  });
  fireEvent.blur(screen.getByRole('textbox', { name: '页码' }));
  fireEvent.click(screen.getByRole('button', { name: '放大' }));
  fireEvent.click(screen.getByRole('button', { name: '整页' }));

  fireEvent.change(screen.getByRole('searchbox', { name: '在 PDF 中搜索' }), {
    target: { value: '架构' },
  });
  fireEvent.submit(
    screen
      .getByRole('searchbox', { name: '在 PDF 中搜索' })
      .closest('form') as HTMLFormElement,
  );
  fireEvent.click(screen.getByRole('button', { name: '保存' }));

  expect(calls).toEqual([
    'undo',
    'annotation:highlight',
    'next-page',
    'page:7',
    'zoom-in',
    'fit-page',
    'search:架构',
    'save',
  ]);
  expect(screen.getByLabelText('PDF 缩放比例')).toHaveTextContent('125%');
  expect(screen.getByText('/ 8')).toBeInTheDocument();
});

function createAnnotationController(calls: string[]): PdfAnnotationController {
  return {
    state: {
      activeToolId: null,
      available: true,
      hasPendingChanges: false,
      selectedCount: 0,
    },
    deleteSelection: () => calls.push('delete-annotation'),
    selectTool: (toolId) => calls.push(`annotation:${toolId ?? 'pointer'}`),
  };
}

function createController(calls: string[]): PdfViewerController {
  const state: PdfViewerControllerState = {
    canRedo: false,
    canUndo: true,
    currentPage: 2,
    documentOpen: true,
    error: null,
    features: {
      export: true,
      history: true,
      navigation: true,
      search: true,
      zoom: true,
    },
    ready: true,
    search: {
      active: false,
      activeResultIndex: -1,
      error: false,
      loading: false,
      query: '',
      total: 0,
    },
    totalPages: 8,
    zoomMode: 'fit-width',
    zoomPercent: 125,
  };

  return {
    state,
    clearSearch: () => calls.push('clear-search'),
    fitPage: () => calls.push('fit-page'),
    fitWidth: () => calls.push('fit-width'),
    goToPage: (page) => calls.push(`page:${page}`),
    nextPage: () => calls.push('next-page'),
    nextSearchResult: () => calls.push('next-result'),
    previousPage: () => calls.push('previous-page'),
    previousSearchResult: () => calls.push('previous-result'),
    redo: () => calls.push('redo'),
    saveAsCopy: () => Promise.resolve(new Blob()),
    search: (query) => calls.push(`search:${query}`),
    undo: () => calls.push('undo'),
    zoomIn: () => calls.push('zoom-in'),
    zoomOut: () => calls.push('zoom-out'),
  };
}
