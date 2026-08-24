import { describe, expect, test } from '@rstest/core';
import { createOfficeEditorRuntime } from '../src/internal/features/work/editors/office-editor-extension';
import type { PdfAnnotationController } from '../src/internal/features/work/editors/pdf-annotation-controller';
import {
  createPdfEditorExtensions,
  type PdfEditorCommandContext,
} from '../src/internal/features/work/editors/pdf-editor-extensions';
import {
  initialPdfViewerControllerState,
  type PdfViewerController,
} from '../src/internal/features/work/editors/pdf-viewer-controller';
import type { PdfPageOrganizationController } from '../src/internal/features/work/editors/use-pdf-page-organization';

describe('PDF editor extensions', () => {
  test('composes navigation, search, zoom, annotation, and save commands', async () => {
    const calls: string[] = [];
    const editor = createOfficeEditorRuntime(
      context(calls),
      createPdfEditorExtensions(),
    );

    editor.commands.goToPage(3);
    editor.commands.search('roadmap');
    editor.commands.nextSearchResult();
    editor.commands.zoomIn();
    editor.commands.selectAnnotationTool('highlight');
    editor.commands.setAnnotationColor('#ff0000');
    editor.commands.setAnnotationOpacity(0.5);
    editor.commands.setAnnotationStrokeWidth(4);
    editor.commands.deleteAnnotationSelection();
    editor.commands.openPageOrganizer();
    await editor.commands.rotatePages([1, 2], 90);
    await editor.commands.insertBlankPage(3);
    await editor.commands.deletePages([3]);
    await editor.commands.reorderPages([4, 0, 1, 2, 3]);
    await editor.commands.mergePages(2, new Blob(['pdf']));
    await editor.commands.extractPages([0, 2]);
    await editor.commands.splitPages([1, 3]);
    await editor.commands.save();

    expect(editor.extensionNames).toEqual([
      'pdfDocument',
      'pdfHistory',
      'pdfNavigation',
      'pdfSearch',
      'pdfZoom',
      'pdfAnnotations',
      'pdfPageOrganization',
      'pdfKeyboardShortcuts',
    ]);
    expect(calls).toEqual([
      'page:3',
      'search:roadmap',
      'search:next',
      'zoom:in',
      'annotation:highlight',
      'annotation:color:#ff0000',
      'annotation:opacity:0.5',
      'annotation:stroke-width:4',
      'annotation:delete',
      'pages:open',
      'pages:rotate:1,2:90',
      'pages:insert:3',
      'pages:delete:3',
      'pages:reorder:4,0,1,2,3',
      'pages:merge:2:3',
      'pages:extract:0,2',
      'pages:split:1,3',
      'save',
    ]);
  });

  test('derives command availability from viewer capabilities', () => {
    const editor = createOfficeEditorRuntime(
      context([]),
      createPdfEditorExtensions(),
    );

    expect(editor.can().previousPage()).toBe(true);
    expect(editor.can().nextPage()).toBe(true);
    expect(editor.can().undo()).toBe(true);
    expect(editor.can().redo()).toBe(false);
    expect(editor.can().selectAnnotationTool('ink')).toBe(true);
    expect(editor.can().setAnnotationColor('#ff0000')).toBe(true);
    expect(editor.can().setAnnotationOpacity(0.5)).toBe(true);
    expect(editor.can().setAnnotationStrokeWidth(4)).toBe(true);
    expect(editor.can().save()).toBe(true);
    expect(editor.can().openPageOrganizer()).toBe(true);
    expect(editor.can().rotatePages([0], 90)).toBe(true);
    expect(editor.can().deletePages([0])).toBe(true);
    expect(editor.can().deletePages([0, 1, 2, 3, 4])).toBe(false);
    expect(editor.can().splitPages([4])).toBe(false);

    editor.updateContext({
      ...context([]),
      viewer: {
        ...context([]).viewer,
        state: initialPdfViewerControllerState,
      },
    });
    expect(editor.can().nextPage()).toBe(false);
    expect(editor.can().zoomIn()).toBe(false);
    expect(editor.can().save()).toBe(false);
    expect(editor.can().openPageOrganizer()).toBe(false);
  });

  test('keeps read-only PDF sessions free of edit, history, and save commands', () => {
    const calls: string[] = [];
    const editorContext = context(calls);
    editorContext.editable = false;
    const editor = createOfficeEditorRuntime(
      editorContext,
      createPdfEditorExtensions(),
    );

    expect(editor.can().undo()).toBe(false);
    expect(editor.can().redo()).toBe(false);
    expect(editor.can().save()).toBe(false);
    expect(editor.can().selectAnnotationTool('highlight')).toBe(false);
    expect(editor.can().deleteAnnotationSelection()).toBe(false);
    expect(editor.can().openPageOrganizer()).toBe(false);
    expect(editor.can().rotatePages([0], 90)).toBe(false);
    expect(editor.can().nextPage()).toBe(true);
    expect(editor.can().zoomIn()).toBe(true);
    expect(editor.can().search('A3S')).toBe(true);

    const undo = new KeyboardEvent('keydown', {
      cancelable: true,
      key: 'z',
      metaKey: true,
    });
    const save = new KeyboardEvent('keydown', {
      cancelable: true,
      key: 's',
      metaKey: true,
    });
    expect(editor.handleKeyDown(undo)).toBe(false);
    expect(editor.handleKeyDown(save)).toBe(false);
    expect(undo.defaultPrevented).toBe(false);
    expect(save.defaultPrevented).toBe(false);
    expect(calls).toEqual([]);
  });

  test('offers only annotation tools reported by the PDF capability', () => {
    const calls: string[] = [];
    const editorContext = context(calls);
    editorContext.annotation.state.availableToolIds = ['highlight'];
    const editor = createOfficeEditorRuntime(
      editorContext,
      createPdfEditorExtensions(),
    );

    expect(editor.can().selectAnnotationTool('highlight')).toBe(true);
    expect(editor.can().selectAnnotationTool('ink')).toBe(false);
    expect(calls).toEqual([]);
  });

  test('disables stale search-result navigation while search is pending or failed', () => {
    const editorContext = context([]);
    editorContext.viewer.state.search = {
      active: true,
      activeResultIndex: 0,
      error: false,
      loading: true,
      query: 'A3S',
      total: 3,
    };
    const editor = createOfficeEditorRuntime(
      editorContext,
      createPdfEditorExtensions(),
    );

    expect(editor.can().previousSearchResult()).toBe(false);
    expect(editor.can().nextSearchResult()).toBe(false);

    editor.updateContext({
      ...editorContext,
      viewer: {
        ...editorContext.viewer,
        state: {
          ...editorContext.viewer.state,
          search: {
            ...editorContext.viewer.state.search,
            error: true,
            loading: false,
          },
        },
      },
    });
    expect(editor.can().previousSearchResult()).toBe(false);
    expect(editor.can().nextSearchResult()).toBe(false);

    editor.updateContext({
      ...editorContext,
      viewer: {
        ...editorContext.viewer,
        state: {
          ...editorContext.viewer.state,
          search: {
            ...editorContext.viewer.state.search,
            error: false,
            loading: false,
          },
        },
      },
    });
    expect(editor.can().previousSearchResult()).toBe(true);
    expect(editor.can().nextSearchResult()).toBe(true);
  });

  test('routes PDF shortcuts through typed commands and capabilities', () => {
    const calls: string[] = [];
    const editor = createOfficeEditorRuntime(
      context(calls),
      createPdfEditorExtensions(),
    );
    const zoom = new KeyboardEvent('keydown', {
      cancelable: true,
      key: '=',
      metaKey: true,
    });

    expect(editor.handleKeyDown(zoom)).toBe(true);
    expect(zoom.defaultPrevented).toBe(true);
    expect(calls).toEqual(['zoom:in']);

    const save = new KeyboardEvent('keydown', {
      cancelable: true,
      key: 's',
      metaKey: true,
    });
    expect(editor.handleKeyDown(save)).toBe(true);
    expect(save.defaultPrevented).toBe(true);
    expect(calls).toEqual(['zoom:in', 'save']);
  });

  test('keeps the active annotation tool when Escape belongs to a popover', () => {
    const calls: string[] = [];
    const editorContext = context(calls);
    editorContext.annotation.state.activeToolId = 'ink';
    const editor = createOfficeEditorRuntime(
      editorContext,
      createPdfEditorExtensions(),
    );
    const boundary = document.createElement('div');
    boundary.dataset.officeShortcuts = 'ignore';
    const control = document.createElement('button');
    boundary.append(control);
    document.body.append(boundary);
    try {
      let handled = true;
      boundary.addEventListener('keydown', (event) => {
        handled = editor.handleKeyDown(event);
      });

      const escapeEvent = new KeyboardEvent('keydown', {
        bubbles: true,
        cancelable: true,
        key: 'Escape',
      });
      control.dispatchEvent(escapeEvent);

      expect(handled).toBe(false);
      expect(escapeEvent.defaultPrevented).toBe(false);
      expect(calls).toEqual([]);
    } finally {
      boundary.remove();
    }
  });

  test('leaves text-field editing shortcuts with the focused control', () => {
    const calls: string[] = [];
    const editorContext = context(calls);
    editorContext.annotation.state.activeToolId = 'ink';
    const editor = createOfficeEditorRuntime(
      editorContext,
      createPdfEditorExtensions(),
    );
    const input = document.createElement('input');
    document.body.append(input);
    try {
      const shortcuts = [
        new KeyboardEvent('keydown', {
          bubbles: true,
          cancelable: true,
          key: 'z',
          metaKey: true,
        }),
        new KeyboardEvent('keydown', {
          bubbles: true,
          cancelable: true,
          key: 'Backspace',
        }),
        new KeyboardEvent('keydown', {
          bubbles: true,
          cancelable: true,
          key: 'Escape',
        }),
      ];
      const handled: boolean[] = [];
      input.addEventListener('keydown', (event) => {
        handled.push(editor.handleKeyDown(event));
      });
      input.focus();
      for (const shortcut of shortcuts) input.dispatchEvent(shortcut);

      expect(handled).toEqual([false, false, false]);
      expect(shortcuts.every((shortcut) => !shortcut.defaultPrevented)).toBe(
        true,
      );
      expect(calls).toEqual([]);
    } finally {
      input.remove();
    }
  });
});

function context(calls: string[]): PdfEditorCommandContext {
  const viewer: PdfViewerController = {
    state: {
      ...initialPdfViewerControllerState,
      canUndo: true,
      currentPage: 2,
      documentOpen: true,
      features: {
        export: true,
        history: true,
        navigation: true,
        search: true,
        zoom: true,
      },
      ready: true,
      totalPages: 5,
    },
    clearSearch: () => calls.push('search:clear'),
    fitPage: () => calls.push('zoom:page'),
    fitWidth: () => calls.push('zoom:width'),
    goToPage: (page) => calls.push(`page:${page}`),
    nextPage: () => calls.push('page:next'),
    nextSearchResult: () => calls.push('search:next'),
    previousPage: () => calls.push('page:previous'),
    previousSearchResult: () => calls.push('search:previous'),
    redo: () => calls.push('history:redo'),
    saveAsCopy: () => Promise.resolve(new Blob()),
    search: (query) => calls.push(`search:${query}`),
    undo: () => calls.push('history:undo'),
    zoomIn: () => calls.push('zoom:in'),
    zoomOut: () => calls.push('zoom:out'),
  };
  const annotation: PdfAnnotationController = {
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
      selectedCount: 1,
      supportsOpacity: true,
      supportsStrokeWidth: true,
    },
    deleteSelection: () => calls.push('annotation:delete'),
    setAnnotationColor: (color) => calls.push(`annotation:color:${color}`),
    setAnnotationOpacity: (opacity) =>
      calls.push(`annotation:opacity:${opacity}`),
    setAnnotationStrokeWidth: (strokeWidth) =>
      calls.push(`annotation:stroke-width:${strokeWidth}`),
    selectTool: (tool) => calls.push(`annotation:${tool ?? 'select'}`),
  };
  const pages: PdfPageOrganizationController = {
    state: {
      available: true,
      busy: false,
      canRedo: false,
      canUndo: false,
      diagnostics: [],
      error: null,
      revision: 0,
    },
    dismissError: () => calls.push('pages:dismiss-error'),
    exportPages: async (operation) => {
      if (operation.kind === 'extract') {
        calls.push(`pages:extract:${operation.pageIndexes.join(',')}`);
      } else {
        calls.push(`pages:split:${operation.splitAfterPageIndexes.join(',')}`);
      }
      return true;
    },
    mutate: async (mutation, source) => {
      if (mutation.kind === 'rotate') {
        calls.push(
          `pages:rotate:${mutation.pageIndexes.join(',')}:${mutation.degrees}`,
        );
      } else if (mutation.kind === 'insert-blank') {
        calls.push(`pages:insert:${mutation.index}`);
      } else if (mutation.kind === 'delete') {
        calls.push(`pages:delete:${mutation.pageIndexes.join(',')}`);
      } else if (mutation.kind === 'reorder') {
        calls.push(`pages:reorder:${mutation.pageOrder.join(',')}`);
      } else {
        calls.push(`pages:merge:${mutation.index}:${source?.size ?? 0}`);
      }
      return true;
    },
    redo: () => calls.push('pages:redo'),
    undo: () => calls.push('pages:undo'),
  };
  return {
    annotation,
    editable: true,
    focusSearch: () => calls.push('search:focus'),
    openPageOrganizer: () => calls.push('pages:open'),
    pages,
    save: {
      enabled: true,
      execute: async () => {
        calls.push('save');
      },
    },
    viewer,
  };
}
