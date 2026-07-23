import { lazy, Suspense, type ReactNode } from 'react';
import { WorkEditorLoadingState } from './internal/features/work/components/work-editor-loading-state';
import type { DocumentEditorProps as InternalDocumentEditorProps } from './internal/features/work/editors/document-editor';
import type { MarkdownEditorProps as InternalMarkdownEditorProps } from './internal/features/work/editors/markdown-editor';
import type { PdfViewerProps as InternalPdfViewerProps } from './internal/features/work/editors/pdf-viewer';
import type { PresentationEditorProps as InternalPresentationEditorProps } from './internal/features/work/editors/presentation-editor';
import type { SpreadsheetEditorProps as InternalSpreadsheetEditorProps } from './internal/features/work/editors/spreadsheet-editor';
import type { WorkOfficeFileAction } from './internal/features/work/editors/work-office-chrome';
import {
  OfficeSurface,
  type OfficeSurfaceProps,
  type OfficeTheme,
} from './office-surface';

const PDFIUM_WASM_FILE_NAME = 'pdfium.wasm';
const OFFICE_KERNEL_WASM_FILE_NAME = 'office-kernel.wasm';

const loadDocumentEditor = () =>
  import(
    /* webpackChunkName: "document-editor" */
    './internal/features/work/editors/document-editor'
  );
const loadMarkdownEditor = () =>
  import(
    /* webpackChunkName: "markdown-editor" */
    './internal/features/work/editors/markdown-editor'
  );
const loadSpreadsheetEditor = () =>
  import(
    /* webpackChunkName: "spreadsheet-editor" */
    './internal/features/work/editors/spreadsheet-editor'
  );
const loadPresentationEditor = () =>
  import(
    /* webpackChunkName: "presentation-editor" */
    './internal/features/work/editors/presentation-editor'
  );
const loadPdfViewer = () =>
  import(
    /* webpackChunkName: "pdf-viewer" */
    './internal/features/work/editors/pdf-viewer'
  );

const LazyDocumentEditor = lazy(async () => ({
  default: (await loadDocumentEditor()).DocumentEditor,
}));
const LazyMarkdownEditor = lazy(async () => ({
  default: (await loadMarkdownEditor()).MarkdownEditor,
}));
const LazySpreadsheetEditor = lazy(async () => ({
  default: (await loadSpreadsheetEditor()).SpreadsheetEditor,
}));
const LazyPresentationEditor = lazy(async () => ({
  default: (await loadPresentationEditor()).PresentationEditor,
}));
const LazyPdfViewer = lazy(async () => ({
  default: (await loadPdfViewer()).PdfViewer,
}));

export type OfficeEditorKind =
  | 'document'
  | 'markdown'
  | 'spreadsheet'
  | 'presentation'
  | 'pdf';

const officeEditorLoaders: Record<OfficeEditorKind, () => Promise<unknown>> = {
  document: loadDocumentEditor,
  markdown: loadMarkdownEditor,
  spreadsheet: loadSpreadsheetEditor,
  presentation: loadPresentationEditor,
  pdf: loadPdfViewer,
};

/**
 * Starts loading one editor without mounting it.
 *
 * Call this from an intent signal such as hover or keyboard focus to keep the
 * initial application bundle small without adding latency to editor opening.
 */
export async function preloadOfficeEditor(
  kind: OfficeEditorKind,
): Promise<void> {
  await officeEditorLoaders[kind]();
}

export const defaultPdfiumWasmUrl = siblingAssetUrl(
  import.meta.url,
  PDFIUM_WASM_FILE_NAME,
);
export const defaultOfficeKernelWasmUrl = siblingAssetUrl(
  import.meta.url,
  OFFICE_KERNEL_WASM_FILE_NAME,
);

function siblingAssetUrl(moduleUrl: string, fileName: string): string {
  return `${moduleUrl.slice(0, moduleUrl.lastIndexOf('/') + 1)}${fileName}`;
}

export type { WorkOfficeFileAction as OfficeFileAction };
export type { OfficeSurfaceProps, OfficeTheme };

function OfficeEditorLoader({
  children,
  title,
}: {
  children: ReactNode;
  title: string;
}) {
  return (
    <Suspense fallback={<WorkEditorLoadingState title={title} />}>
      {children}
    </Suspense>
  );
}

export interface DocumentEditorProps
  extends Omit<InternalDocumentEditorProps, 'preview'>,
    OfficeSurfaceProps {
  preview?: boolean;
}

export function DocumentEditor({
  className,
  kernelWasmUrl = defaultOfficeKernelWasmUrl,
  preview = false,
  style,
  theme,
  ...editorProps
}: DocumentEditorProps) {
  return (
    <OfficeSurface className={className} style={style} theme={theme}>
      <OfficeEditorLoader title="正在打开文字编辑器">
        <LazyDocumentEditor
          {...editorProps}
          kernelWasmUrl={kernelWasmUrl}
          preview={preview}
        />
      </OfficeEditorLoader>
    </OfficeSurface>
  );
}

export interface MarkdownEditorProps
  extends Omit<InternalMarkdownEditorProps, 'preview'>,
    OfficeSurfaceProps {
  preview?: boolean;
}

export function MarkdownEditor({
  className,
  preview = false,
  style,
  theme,
  ...editorProps
}: MarkdownEditorProps) {
  return (
    <OfficeSurface className={className} style={style} theme={theme}>
      <OfficeEditorLoader title="正在打开 Markdown 编辑器">
        <LazyMarkdownEditor {...editorProps} preview={preview} />
      </OfficeEditorLoader>
    </OfficeSurface>
  );
}

export interface SpreadsheetEditorProps
  extends Omit<InternalSpreadsheetEditorProps, 'preview'>,
    OfficeSurfaceProps {
  preview?: boolean;
}

export function SpreadsheetEditor({
  className,
  preview = false,
  style,
  theme,
  ...editorProps
}: SpreadsheetEditorProps) {
  return (
    <OfficeSurface className={className} style={style} theme={theme}>
      <OfficeEditorLoader title="正在打开表格编辑器">
        <LazySpreadsheetEditor {...editorProps} preview={preview} />
      </OfficeEditorLoader>
    </OfficeSurface>
  );
}

export interface PresentationEditorProps
  extends Omit<InternalPresentationEditorProps, 'preview'>,
    OfficeSurfaceProps {
  preview?: boolean;
}

export function PresentationEditor({
  className,
  preview = false,
  style,
  theme,
  ...editorProps
}: PresentationEditorProps) {
  return (
    <OfficeSurface className={className} style={style} theme={theme}>
      <OfficeEditorLoader title="正在打开演示编辑器">
        <LazyPresentationEditor {...editorProps} preview={preview} />
      </OfficeEditorLoader>
    </OfficeSurface>
  );
}

export interface PdfViewerProps
  extends InternalPdfViewerProps,
    OfficeSurfaceProps {}

export function PdfViewer({
  className,
  style,
  theme,
  wasmUrl = defaultPdfiumWasmUrl,
  ...viewerProps
}: PdfViewerProps) {
  return (
    <OfficeSurface className={className} style={style} theme={theme}>
      <OfficeEditorLoader title="正在打开 PDF">
        <LazyPdfViewer {...viewerProps} wasmUrl={wasmUrl} />
      </OfficeEditorLoader>
    </OfficeSurface>
  );
}
