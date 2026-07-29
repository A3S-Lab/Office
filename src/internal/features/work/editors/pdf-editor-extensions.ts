import type {
  PdfAnnotationController,
  PdfAnnotationToolId,
} from './pdf-annotation-controller';
import { isOfficeShortcutBlocked } from './office-shortcuts';
import type { PdfViewerController } from './pdf-viewer-controller';
import {
  createOfficeEditorExtension,
  type OfficeEditorCanCommands,
  type OfficeEditorExtension,
} from './office-editor-extension';

export interface PdfEditorCommands {
  clearSearch: () => void;
  deleteAnnotationSelection: () => void;
  fitPage: () => void;
  fitWidth: () => void;
  goToPage: (page: number) => void;
  nextPage: () => void;
  nextSearchResult: () => void;
  previousPage: () => void;
  previousSearchResult: () => void;
  redo: () => void;
  save: () => Promise<void>;
  search: (query: string) => void;
  setAnnotationColor: (color: string) => void;
  setAnnotationOpacity: (opacity: number) => void;
  setAnnotationStrokeWidth: (strokeWidth: number) => void;
  selectAnnotationTool: (toolId: PdfAnnotationToolId | null) => void;
  undo: () => void;
  zoomIn: () => void;
  zoomOut: () => void;
}

export type PdfEditorCanCommands = OfficeEditorCanCommands<PdfEditorCommands>;

export interface PdfEditorCommandContext {
  annotation: PdfAnnotationController;
  editable: boolean;
  focusSearch: () => void;
  save: {
    enabled: boolean;
    execute: () => Promise<void>;
  };
  viewer: PdfViewerController;
}

export function createPdfEditorExtensions(): readonly OfficeEditorExtension<
  PdfEditorCommandContext,
  PdfEditorCommands
>[] {
  return [
    createOfficeEditorExtension<PdfEditorCommandContext, PdfEditorCommands>({
      name: 'pdfDocument',
      addCommands: () => ({
        save: {
          canExecute: canSave,
          execute: ({ save }) => save.execute(),
        },
      }),
    }),
    createOfficeEditorExtension<PdfEditorCommandContext, PdfEditorCommands>({
      name: 'pdfHistory',
      addCommands: () => ({
        redo: {
          canExecute: ({ editable, viewer }) =>
            editable &&
            documentReady(viewer) &&
            viewer.state.features.history &&
            viewer.state.canRedo,
          execute: ({ viewer }) => viewer.redo(),
        },
        undo: {
          canExecute: ({ editable, viewer }) =>
            editable &&
            documentReady(viewer) &&
            viewer.state.features.history &&
            viewer.state.canUndo,
          execute: ({ viewer }) => viewer.undo(),
        },
      }),
    }),
    createOfficeEditorExtension<PdfEditorCommandContext, PdfEditorCommands>({
      name: 'pdfNavigation',
      addCommands: () => ({
        goToPage: {
          canExecute: ({ viewer }, page) =>
            canNavigate(viewer) &&
            Number.isInteger(page) &&
            page >= 1 &&
            page <= viewer.state.totalPages,
          execute: ({ viewer }, page) => viewer.goToPage(page),
        },
        nextPage: {
          canExecute: ({ viewer }) =>
            canNavigate(viewer) &&
            viewer.state.currentPage < viewer.state.totalPages,
          execute: ({ viewer }) => viewer.nextPage(),
        },
        previousPage: {
          canExecute: ({ viewer }) =>
            canNavigate(viewer) && viewer.state.currentPage > 1,
          execute: ({ viewer }) => viewer.previousPage(),
        },
      }),
    }),
    createOfficeEditorExtension<PdfEditorCommandContext, PdfEditorCommands>({
      name: 'pdfSearch',
      addCommands: () => ({
        clearSearch: {
          canExecute: ({ viewer }) => canSearch(viewer),
          execute: ({ viewer }) => viewer.clearSearch(),
        },
        nextSearchResult: {
          canExecute: ({ viewer }) => canNavigateSearchResults(viewer),
          execute: ({ viewer }) => viewer.nextSearchResult(),
        },
        previousSearchResult: {
          canExecute: ({ viewer }) => canNavigateSearchResults(viewer),
          execute: ({ viewer }) => viewer.previousSearchResult(),
        },
        search: {
          canExecute: ({ viewer }) => canSearch(viewer),
          execute: ({ viewer }, query) => viewer.search(query),
        },
      }),
    }),
    createOfficeEditorExtension<PdfEditorCommandContext, PdfEditorCommands>({
      name: 'pdfZoom',
      addCommands: () => ({
        fitPage: {
          canExecute: ({ viewer }) => canZoom(viewer),
          execute: ({ viewer }) => viewer.fitPage(),
        },
        fitWidth: {
          canExecute: ({ viewer }) => canZoom(viewer),
          execute: ({ viewer }) => viewer.fitWidth(),
        },
        zoomIn: {
          canExecute: ({ viewer }) => canZoom(viewer),
          execute: ({ viewer }) => viewer.zoomIn(),
        },
        zoomOut: {
          canExecute: ({ viewer }) => canZoom(viewer),
          execute: ({ viewer }) => viewer.zoomOut(),
        },
      }),
    }),
    createOfficeEditorExtension<PdfEditorCommandContext, PdfEditorCommands>({
      name: 'pdfAnnotations',
      addCommands: () => ({
        deleteAnnotationSelection: {
          canExecute: ({ annotation, editable, viewer }) =>
            editable &&
            documentReady(viewer) &&
            annotation.state.available &&
            annotation.state.selectedCount > 0,
          execute: ({ annotation }) => annotation.deleteSelection(),
        },
        selectAnnotationTool: {
          canExecute: ({ annotation, editable, viewer }, toolId) =>
            editable &&
            documentReady(viewer) &&
            annotation.state.available &&
            (toolId === null ||
              annotation.state.availableToolIds.includes(toolId)),
          execute: ({ annotation }, toolId) => annotation.selectTool(toolId),
        },
        setAnnotationColor: {
          canExecute: ({ annotation, editable, viewer }, color) =>
            editable &&
            documentReady(viewer) &&
            annotation.state.available &&
            (annotation.state.selectedCount > 0 ||
              Boolean(annotation.state.activeToolId)) &&
            /^#[0-9a-f]{6}$/i.test(color),
          execute: ({ annotation }, color) =>
            annotation.setAnnotationColor(color),
        },
        setAnnotationOpacity: {
          canExecute: ({ annotation, editable, viewer }, opacity) =>
            editable &&
            documentReady(viewer) &&
            annotation.state.available &&
            annotation.state.supportsOpacity &&
            Number.isFinite(opacity) &&
            opacity >= 0 &&
            opacity <= 1,
          execute: ({ annotation }, opacity) =>
            annotation.setAnnotationOpacity(opacity),
        },
        setAnnotationStrokeWidth: {
          canExecute: ({ annotation, editable, viewer }, strokeWidth) =>
            editable &&
            documentReady(viewer) &&
            annotation.state.available &&
            annotation.state.supportsStrokeWidth &&
            Number.isFinite(strokeWidth) &&
            strokeWidth > 0 &&
            strokeWidth <= 24,
          execute: ({ annotation }, strokeWidth) =>
            annotation.setAnnotationStrokeWidth(strokeWidth),
        },
      }),
    }),
    createOfficeEditorExtension<PdfEditorCommandContext, PdfEditorCommands>({
      name: 'pdfKeyboardShortcuts',
      addKeyboardShortcuts: () => ({
        'Mod-s': ({ can, commands }, event) =>
          runPdfShortcut(event, can.save, () => {
            void commands.save();
          }),
        'Mod-z': ({ can, commands }, event) =>
          runPdfTextAwareShortcut(event, can.undo, commands.undo),
        'Mod-Shift-z': ({ can, commands }, event) =>
          runPdfTextAwareShortcut(event, can.redo, commands.redo),
        'Mod-y': ({ can, commands }, event) =>
          runPdfTextAwareShortcut(event, can.redo, commands.redo),
        'Mod-Equal': ({ can, commands }, event) =>
          runPdfShortcut(event, can.zoomIn, commands.zoomIn),
        'Mod-Shift-Plus': ({ can, commands }, event) =>
          runPdfShortcut(event, can.zoomIn, commands.zoomIn),
        'Mod-Minus': ({ can, commands }, event) =>
          runPdfShortcut(event, can.zoomOut, commands.zoomOut),
        'Mod-0': ({ can, commands }, event) =>
          runPdfShortcut(event, can.fitPage, commands.fitPage),
        'Mod-f': ({ can, context }, event) => {
          if (
            event.repeat ||
            isOfficeShortcutBlocked(event.target) ||
            !can.search('')
          ) {
            return false;
          }
          context.focusSearch();
          return true;
        },
        Delete: ({ can, commands }, event) =>
          runPdfTextAwareShortcut(
            event,
            can.deleteAnnotationSelection,
            commands.deleteAnnotationSelection,
          ),
        Backspace: ({ can, commands }, event) =>
          runPdfTextAwareShortcut(
            event,
            can.deleteAnnotationSelection,
            commands.deleteAnnotationSelection,
          ),
        Escape: ({ can, commands, context }, event) => {
          if (
            event.repeat ||
            isOfficeShortcutBlocked(event.target) ||
            isPdfTextEditingTarget(event.target) ||
            !context.annotation.state.activeToolId ||
            !can.selectAnnotationTool(null)
          ) {
            return false;
          }
          commands.selectAnnotationTool(null);
          return true;
        },
      }),
    }),
  ];
}

function runPdfShortcut(
  event: KeyboardEvent,
  canExecute: () => boolean,
  execute: () => void,
): boolean {
  if (event.repeat || isOfficeShortcutBlocked(event.target) || !canExecute()) {
    return false;
  }
  execute();
  return true;
}

function runPdfTextAwareShortcut(
  event: KeyboardEvent,
  canExecute: () => boolean,
  execute: () => void,
): boolean {
  if (isPdfTextEditingTarget(event.target)) return false;
  return runPdfShortcut(event, canExecute, execute);
}

function isPdfTextEditingTarget(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) return false;
  return (
    target instanceof HTMLInputElement ||
    target instanceof HTMLTextAreaElement ||
    target instanceof HTMLSelectElement ||
    target.isContentEditable ||
    Boolean(target.closest('[contenteditable="true"]'))
  );
}

function documentReady(viewer: PdfViewerController): boolean {
  return viewer.state.ready && viewer.state.documentOpen;
}

function canNavigateSearchResults(viewer: PdfViewerController): boolean {
  return (
    canSearch(viewer) &&
    !viewer.state.search.loading &&
    !viewer.state.search.error &&
    viewer.state.search.total > 0
  );
}

function canNavigate(viewer: PdfViewerController): boolean {
  return documentReady(viewer) && viewer.state.features.navigation;
}

function canSearch(viewer: PdfViewerController): boolean {
  return documentReady(viewer) && viewer.state.features.search;
}

function canZoom(viewer: PdfViewerController): boolean {
  return documentReady(viewer) && viewer.state.features.zoom;
}

function canSave({ editable, save, viewer }: PdfEditorCommandContext): boolean {
  return (
    editable &&
    save.enabled &&
    documentReady(viewer) &&
    viewer.state.features.export
  );
}
