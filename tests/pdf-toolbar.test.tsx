import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, within } from '@testing-library/react';
import { createRef } from 'react';
import type {
  PdfEditorCanCommands,
  PdfEditorCommands,
} from '../src/internal/features/work/editors/pdf-editor-extensions';
import { PdfToolbar } from '../src/internal/features/work/editors/pdf-toolbar';
import type { PdfAnnotationController } from '../src/internal/features/work/editors/pdf-annotation-controller';
import type {
  PdfViewerController,
  PdfViewerControllerState,
} from '../src/internal/features/work/editors/pdf-viewer-controller';

test('keeps PDF navigation, search, zoom, history, and save in one toolbar', () => {
  const calls: string[] = [];
  const controller = createController(calls);
  const annotation = createAnnotationController(calls);
  const searchRef = createRef<HTMLInputElement>();
  const can = createCanCommands(controller);
  const commands = createCommands(controller, annotation, calls);

  render(
    <PdfToolbar
      annotationState={annotation.state}
      can={can}
      commands={commands}
      editable
      saveLabel="保存"
      saveState="idle"
      searchInputRef={searchRef}
      state={controller.state}
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

test('keeps compact PDF actions reachable from the more-tools menu', () => {
  const calls: string[] = [];
  const controller = createController(calls);
  const annotation = createAnnotationController(calls);

  render(
    <PdfToolbar
      annotationState={annotation.state}
      can={createCanCommands(controller)}
      commands={createCommands(controller, annotation, calls)}
      editable
      saveLabel="保存"
      saveState="idle"
      searchInputRef={createRef<HTMLInputElement>()}
      state={controller.state}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '更多 PDF 工具' }));
  let menu = screen.getByRole('menu', { name: '更多 PDF 工具' });
  fireEvent.click(within(menu).getByRole('menuitem', { name: '下划线批注' }));

  fireEvent.click(screen.getByRole('button', { name: '更多 PDF 工具' }));
  menu = screen.getByRole('menu', { name: '更多 PDF 工具' });
  fireEvent.click(within(menu).getByRole('menuitem', { name: '页宽' }));

  expect(calls).toEqual(['annotation:underline', 'fit-width']);
});

function createCanCommands(
  controller: PdfViewerController,
): PdfEditorCanCommands {
  const ready = () => controller.state.ready && controller.state.documentOpen;
  return {
    clearSearch: ready,
    deleteAnnotationSelection: () => false,
    fitPage: ready,
    fitWidth: ready,
    goToPage: (page) =>
      ready() &&
      Number.isInteger(page) &&
      page >= 1 &&
      page <= controller.state.totalPages,
    nextPage: () =>
      ready() && controller.state.currentPage < controller.state.totalPages,
    nextSearchResult: () => ready() && controller.state.search.total > 0,
    previousPage: () => ready() && controller.state.currentPage > 1,
    previousSearchResult: () => ready() && controller.state.search.total > 0,
    redo: () => ready() && controller.state.canRedo,
    save: ready,
    search: ready,
    setAnnotationColor: ready,
    selectAnnotationTool: ready,
    undo: () => ready() && controller.state.canUndo,
    zoomIn: ready,
    zoomOut: ready,
  };
}

function createCommands(
  controller: PdfViewerController,
  annotation: PdfAnnotationController,
  calls: string[],
): PdfEditorCommands {
  return {
    clearSearch: controller.clearSearch,
    deleteAnnotationSelection: annotation.deleteSelection,
    fitPage: controller.fitPage,
    fitWidth: controller.fitWidth,
    goToPage: controller.goToPage,
    nextPage: controller.nextPage,
    nextSearchResult: controller.nextSearchResult,
    previousPage: controller.previousPage,
    previousSearchResult: controller.previousSearchResult,
    redo: controller.redo,
    save: async () => {
      calls.push('save');
    },
    search: controller.search,
    setAnnotationColor: annotation.setAnnotationColor,
    selectAnnotationTool: annotation.selectTool,
    undo: controller.undo,
    zoomIn: controller.zoomIn,
    zoomOut: controller.zoomOut,
  };
}

function createAnnotationController(calls: string[]): PdfAnnotationController {
  return {
    state: {
      activeToolId: null,
      annotationColor: '#ffd966',
      available: true,
      hasPendingChanges: false,
      selectedCount: 0,
    },
    deleteSelection: () => calls.push('delete-annotation'),
    setAnnotationColor: (color) => calls.push(`annotation-color:${color}`),
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
