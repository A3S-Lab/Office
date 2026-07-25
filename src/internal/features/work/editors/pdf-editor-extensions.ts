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
          canExecute: ({ viewer }) =>
            documentReady(viewer) && viewer.state.canRedo,
          execute: ({ viewer }) => viewer.redo(),
        },
        undo: {
          canExecute: ({ viewer }) =>
            documentReady(viewer) && viewer.state.canUndo,
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
          canExecute: ({ viewer }) =>
            canSearch(viewer) && viewer.state.search.total > 0,
          execute: ({ viewer }) => viewer.nextSearchResult(),
        },
        previousSearchResult: {
          canExecute: ({ viewer }) =>
            canSearch(viewer) && viewer.state.search.total > 0,
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
          canExecute: ({ annotation, editable, viewer }) =>
            editable && documentReady(viewer) && annotation.state.available,
          execute: ({ annotation }, toolId) => annotation.selectTool(toolId),
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
          runPdfShortcut(event, can.undo, commands.undo),
        'Mod-Shift-z': ({ can, commands }, event) =>
          runPdfShortcut(event, can.redo, commands.redo),
        'Mod-y': ({ can, commands }, event) =>
          runPdfShortcut(event, can.redo, commands.redo),
        'Mod-Equal': ({ can, commands }, event) =>
          runPdfShortcut(event, can.zoomIn, commands.zoomIn),
        'Mod-Shift-Plus': ({ can, commands }, event) =>
          runPdfShortcut(event, can.zoomIn, commands.zoomIn),
        'Mod-Minus': ({ can, commands }, event) =>
          runPdfShortcut(event, can.zoomOut, commands.zoomOut),
        'Mod-0': ({ can, commands }, event) =>
          runPdfShortcut(event, can.fitPage, commands.fitPage),
        'Mod-f': ({ can, context }, event) => {
          if (event.repeat || !can.search('')) return false;
          context.focusSearch();
          return true;
        },
        Delete: ({ can, commands }, event) =>
          runPdfShortcut(
            event,
            can.deleteAnnotationSelection,
            commands.deleteAnnotationSelection,
          ),
        Backspace: ({ can, commands }, event) =>
          runPdfShortcut(
            event,
            can.deleteAnnotationSelection,
            commands.deleteAnnotationSelection,
          ),
        Escape: ({ can, commands, context }, event) => {
          if (
            event.repeat ||
            isOfficeShortcutBlocked(event.target) ||
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

function documentReady(viewer: PdfViewerController): boolean {
  return viewer.state.ready && viewer.state.documentOpen;
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

function canSave({ save, viewer }: PdfEditorCommandContext): boolean {
  return save.enabled && documentReady(viewer) && viewer.state.features.export;
}
