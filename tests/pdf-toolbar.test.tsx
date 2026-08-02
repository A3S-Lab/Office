import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { createRef } from 'react';
import type { PdfAnnotationController } from '../src/internal/features/work/editors/pdf-annotation-controller';
import type {
  PdfEditorCanCommands,
  PdfEditorCommands,
} from '../src/internal/features/work/editors/pdf-editor-extensions';
import { PdfToolbar } from '../src/internal/features/work/editors/pdf-toolbar';
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
  fireEvent.click(screen.getByRole('button', { name: '批注样式' }));
  fireEvent.click(screen.getByRole('radio', { name: '透明度 50%' }));
  fireEvent.click(screen.getByRole('radio', { name: '线宽 4' }));

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
    'annotation-opacity:0.5',
    'annotation-stroke-width:4',
    'search:架构',
    'save',
  ]);
  expect(screen.getByLabelText('PDF 缩放比例')).toHaveTextContent('125%');
  expect(screen.getByText('/ 8')).toBeInTheDocument();
});

test('keeps the compact page-navigation trigger inside the page controls', () => {
  const controller = createController([]);
  const annotation = createAnnotationController([]);
  const toggleRef = createRef<HTMLButtonElement>();
  let openCount = 0;

  render(
    <PdfToolbar
      annotationState={annotation.state}
      can={createCanCommands(controller)}
      commands={createCommands(controller, annotation, [])}
      editable
      pageNavigation={{
        controlsId: 'pdf-page-navigation',
        expanded: false,
        onOpen: () => {
          openCount += 1;
        },
        toggleRef,
      }}
      saveLabel="保存"
      saveState="idle"
      searchInputRef={createRef<HTMLInputElement>()}
      state={controller.state}
    />,
  );

  const toggle = screen.getByRole('button', {
    name: '打开 PDF 页面导航',
  });
  expect(toggle.closest('.work-pdf-page-controls')).not.toBeNull();
  expect(toggleRef.current).toBe(toggle);
  expect(toggle).toHaveAttribute('aria-controls', 'pdf-page-navigation');
  expect(toggle).toHaveAttribute('aria-expanded', 'false');
  expect(toggle).toHaveTextContent('第 2 页');

  fireEvent.click(toggle);
  expect(openCount).toBe(1);
});

test('keeps read-only PDF chrome free of edit-only commands', () => {
  const controller = createController([]);
  const annotation = createAnnotationController([]);

  render(
    <PdfToolbar
      annotationState={annotation.state}
      can={createCanCommands(controller)}
      commands={createCommands(controller, annotation, [])}
      editable={false}
      saveLabel="保存"
      saveState="idle"
      searchInputRef={createRef<HTMLInputElement>()}
      state={controller.state}
    />,
  );

  expect(screen.queryByRole('button', { name: '保存' })).toBeNull();
  expect(screen.queryByRole('button', { name: '撤销' })).toBeNull();
  expect(screen.queryByRole('button', { name: '重做' })).toBeNull();
  expect(screen.queryByRole('button', { name: '选择' })).toBeNull();
  expect(screen.queryByRole('button', { name: '高亮' })).toBeNull();

  fireEvent.click(screen.getByRole('button', { name: '更多 PDF 工具' }));
  const menu = screen.getByRole('menu', { name: '更多 PDF 工具' });
  expect(within(menu).queryByRole('menuitem', { name: '撤销' })).toBeNull();
  expect(
    within(menu).queryByRole('menuitemradio', { name: '下划线批注' }),
  ).toBeNull();
  expect(within(menu).getByRole('menuitem', { name: '下一页' })).toBeEnabled();
  expect(within(menu).getByRole('menuitem', { name: '放大' })).toBeEnabled();
});

test('advertises only shortcuts implemented by the PDF command surface', () => {
  const controller = createController([]);
  const annotation = createAnnotationController([]);

  render(
    <PdfToolbar
      annotationState={annotation.state}
      can={createCanCommands(controller)}
      commands={createCommands(controller, annotation, [])}
      editable
      saveLabel="保存"
      saveState="idle"
      searchInputRef={createRef<HTMLInputElement>()}
      state={controller.state}
    />,
  );

  const shortcuts = [
    ['保存', 'Control+S Meta+S'],
    ['撤销', 'Control+Z Meta+Z'],
    ['重做', 'Control+Shift+Z Meta+Shift+Z Control+Y Meta+Y'],
    ['在 PDF 中搜索', 'Control+F Meta+F'],
    ['缩小', 'Control+- Meta+-'],
    ['放大', 'Control+= Meta+= Control+Shift++ Meta+Shift++'],
    ['整页', 'Control+0 Meta+0'],
    ['删除所选批注', 'Delete Backspace'],
  ] as const;
  for (const [name, shortcut] of shortcuts) {
    expect(
      screen.getByRole(name === '在 PDF 中搜索' ? 'searchbox' : 'button', {
        name,
      }),
    ).toHaveAttribute('aria-keyshortcuts', shortcut);
  }

  fireEvent.click(screen.getByRole('button', { name: '更多 PDF 工具' }));
  const menu = screen.getByRole('menu', { name: '更多 PDF 工具' });
  expect(within(menu).getByRole('menuitem', { name: '撤销' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Z Meta+Z',
  );
  expect(
    within(menu).getByRole('menuitemradio', { name: '整页' }),
  ).toHaveAttribute('aria-keyshortcuts', 'Control+0 Meta+0');
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
  fireEvent.click(
    within(menu).getByRole('menuitemradio', { name: '下划线批注' }),
  );

  fireEvent.click(screen.getByRole('button', { name: '更多 PDF 工具' }));
  menu = screen.getByRole('menu', { name: '更多 PDF 工具' });
  fireEvent.click(within(menu).getByRole('menuitemradio', { name: '页宽' }));

  fireEvent.click(screen.getByRole('button', { name: '更多 PDF 工具' }));
  menu = screen.getByRole('menu', { name: '更多 PDF 工具' });
  fireEvent.click(within(menu).getByRole('menuitem', { name: '上一页' }));

  fireEvent.click(screen.getByRole('button', { name: '更多 PDF 工具' }));
  menu = screen.getByRole('menu', { name: '更多 PDF 工具' });
  fireEvent.click(within(menu).getByRole('menuitem', { name: '放大' }));

  fireEvent.click(screen.getByRole('button', { name: '更多 PDF 工具' }));
  menu = screen.getByRole('menu', { name: '更多 PDF 工具' });
  fireEvent.click(within(menu).getByRole('menuitem', { name: '撤销' }));

  expect(calls).toEqual([
    'annotation:underline',
    'fit-width',
    'previous-page',
    'zoom-in',
    'undo',
  ]);
});

test('keeps secondary phone annotation controls in the PDF overflow menu', () => {
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

  const openMenu = () => {
    fireEvent.click(screen.getByRole('button', { name: '更多 PDF 工具' }));
    return screen.getByRole('menu', { name: '更多 PDF 工具' });
  };

  let menu = openMenu();
  fireEvent.click(within(menu).getByRole('menuitemradio', { name: '选择' }));
  menu = openMenu();
  fireEvent.click(within(menu).getByRole('menuitemradio', { name: '画笔' }));
  menu = openMenu();
  fireEvent.click(
    within(menu).getByRole('menuitemradio', { name: '透明度 50%' }),
  );
  menu = openMenu();
  fireEvent.click(within(menu).getByRole('menuitemradio', { name: '线宽 4' }));
  menu = openMenu();
  expect(
    within(menu).getByRole('menuitem', { name: '删除所选批注' }),
  ).toBeDisabled();

  expect(calls).toEqual([
    'annotation:pointer',
    'annotation:ink',
    'annotation-opacity:0.5',
    'annotation-stroke-width:4',
  ]);
});

test('moves through and exits the PDF more-tools menu with standard keys', async () => {
  const controller = createController([]);
  const annotation = createAnnotationController([]);

  render(
    <PdfToolbar
      annotationState={annotation.state}
      can={createCanCommands(controller)}
      commands={createCommands(controller, annotation, [])}
      editable
      saveLabel="保存"
      saveState="idle"
      searchInputRef={createRef<HTMLInputElement>()}
      state={controller.state}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '更多 PDF 工具' }));
  const menu = screen.getByRole('menu', { name: '更多 PDF 工具' });
  const enabledItems = [
    ...menu.querySelectorAll<HTMLButtonElement>(
      'button[role^="menuitem"]:not(:disabled)',
    ),
  ];

  fireEvent.keyDown(menu, { key: 'End' });
  expect(enabledItems.at(-1)).toHaveFocus();
  fireEvent.keyDown(menu, { key: 'ArrowDown' });
  expect(enabledItems[0]).toHaveFocus();
  fireEvent.keyDown(menu, { key: 'ArrowUp' });
  expect(enabledItems.at(-1)).toHaveFocus();
  fireEvent.keyDown(menu, { key: 'Home' });
  expect(enabledItems[0]).toHaveFocus();

  fireEvent.keyDown(menu, { key: 'Tab' });
  await waitFor(() => {
    expect(
      screen.queryByRole('menu', { name: '更多 PDF 工具' }),
    ).not.toBeInTheDocument();
    expect(
      screen.getByRole('searchbox', { name: '在 PDF 中搜索' }),
    ).toHaveFocus();
  });
});

test('exposes selected PDF overflow tools as radio menu items', () => {
  const controller = createController([]);
  const annotation = createAnnotationController([]);

  render(
    <PdfToolbar
      annotationState={{ ...annotation.state, activeToolId: 'underline' }}
      can={createCanCommands(controller)}
      commands={createCommands(controller, annotation, [])}
      editable
      saveLabel="保存"
      saveState="idle"
      searchInputRef={createRef<HTMLInputElement>()}
      state={controller.state}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '更多 PDF 工具' }));
  expect(
    screen.getByRole('menuitemradio', { name: '下划线批注' }),
  ).toHaveAttribute('aria-checked', 'true');
  expect(screen.getByRole('menuitemradio', { name: '页宽' })).toHaveAttribute(
    'aria-checked',
    'true',
  );
});

test('marks an open annotation popover as an editor-shortcut boundary', () => {
  const calls: string[] = [];
  const controller = createController(calls);
  const annotation = createAnnotationController(calls);

  render(
    <PdfToolbar
      annotationState={{ ...annotation.state, activeToolId: 'ink' }}
      can={createCanCommands(controller)}
      commands={createCommands(controller, annotation, calls)}
      editable
      saveLabel="保存"
      saveState="idle"
      searchInputRef={createRef<HTMLInputElement>()}
      state={controller.state}
    />,
  );

  const trigger = screen.getByRole('button', { name: '批注样式' });
  fireEvent.click(trigger);

  expect(trigger.closest('.ds-popover')).toHaveAttribute(
    'data-office-shortcuts',
    'ignore',
  );
  expect(screen.getByRole('dialog', { name: '批注样式' })).toHaveAttribute(
    'data-office-shortcuts',
    'ignore',
  );
});

test('opens annotation styles on the current value and keeps arrows in group', async () => {
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

  fireEvent.click(screen.getByRole('button', { name: '批注样式' }));
  const opacity = screen.getByRole('radio', { name: '透明度 100%' });
  await waitFor(() => expect(opacity).toHaveFocus());
  expect(opacity).toBeChecked();

  fireEvent.keyDown(opacity, { key: 'ArrowLeft' });
  expect(screen.getByRole('radio', { name: '透明度 75%' })).toHaveFocus();
  expect(calls).toContain('annotation-opacity:0.75');

  const strokeWidth = screen.getByRole('radio', { name: '线宽 6' });
  strokeWidth.focus();
  fireEvent.keyDown(strokeWidth, { key: 'Home' });
  expect(screen.getByRole('radio', { name: '线宽 1' })).toHaveFocus();
  expect(calls).toContain('annotation-stroke-width:1');
});

test('cancels a page-number draft on Escape without navigating on blur', () => {
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

  const page = screen.getByRole('textbox', { name: '页码' });
  page.focus();
  fireEvent.change(page, { target: { value: '7' } });
  fireEvent.keyDown(page, { key: 'Escape' });

  expect(page).toHaveValue('2');
  expect(page).not.toHaveFocus();
  expect(calls).not.toContain('page:7');
});

test('does not restart or navigate a search while its current query is loading', () => {
  const calls: string[] = [];
  const controller = createController(calls);
  controller.state.search = {
    active: true,
    activeResultIndex: 0,
    error: false,
    loading: true,
    query: '架构',
    total: 3,
  };
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

  const search = screen.getByRole('searchbox', { name: '在 PDF 中搜索' });
  fireEvent.submit(search.closest('form') as HTMLFormElement);
  fireEvent.keyDown(search, { key: 'Enter', shiftKey: true });

  expect(calls).not.toContain('search:架构');
  expect(calls).not.toContain('previous-result');
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
    setAnnotationOpacity: ready,
    setAnnotationStrokeWidth: ready,
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
    setAnnotationOpacity: annotation.setAnnotationOpacity,
    setAnnotationStrokeWidth: annotation.setAnnotationStrokeWidth,
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
      annotationOpacity: 1,
      annotationStrokeWidth: 6,
      available: true,
      availableToolIds: [
        'highlight',
        'underline',
        'strikeout',
        'ink',
        'freeText',
      ],
      hasPendingChanges: false,
      selectedCount: 0,
      supportsOpacity: true,
      supportsStrokeWidth: true,
    },
    deleteSelection: () => calls.push('delete-annotation'),
    setAnnotationColor: (color) => calls.push(`annotation-color:${color}`),
    setAnnotationOpacity: (opacity) =>
      calls.push(`annotation-opacity:${opacity}`),
    setAnnotationStrokeWidth: (strokeWidth) =>
      calls.push(`annotation-stroke-width:${strokeWidth}`),
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
