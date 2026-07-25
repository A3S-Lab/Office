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
    editor.commands.deleteAnnotationSelection();
    await editor.commands.save();

    expect(editor.extensionNames).toEqual([
      'pdfDocument',
      'pdfHistory',
      'pdfNavigation',
      'pdfSearch',
      'pdfZoom',
      'pdfAnnotations',
      'pdfKeyboardShortcuts',
    ]);
    expect(calls).toEqual([
      'page:3',
      'search:roadmap',
      'search:next',
      'zoom:in',
      'annotation:highlight',
      'annotation:delete',
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
    expect(editor.can().save()).toBe(true);

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
      available: true,
      hasPendingChanges: false,
      selectedCount: 1,
    },
    deleteSelection: () => calls.push('annotation:delete'),
    selectTool: (tool) => calls.push(`annotation:${tool ?? 'select'}`),
  };
  return {
    annotation,
    editable: true,
    focusSearch: () => calls.push('search:focus'),
    save: {
      enabled: true,
      execute: async () => {
        calls.push('save');
      },
    },
    viewer,
  };
}
