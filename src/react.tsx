import {
  DocumentEditor as InternalDocumentEditor,
  type DocumentEditorProps as InternalDocumentEditorProps,
} from './internal/features/work/editors/document-editor';
import {
  PdfViewer as InternalPdfViewer,
  type PdfViewerProps as InternalPdfViewerProps,
} from './internal/features/work/editors/pdf-viewer';
import {
  PresentationEditor as InternalPresentationEditor,
  type PresentationEditorProps as InternalPresentationEditorProps,
} from './internal/features/work/editors/presentation-editor';
import {
  SpreadsheetEditor as InternalSpreadsheetEditor,
  type SpreadsheetEditorProps as InternalSpreadsheetEditorProps,
} from './internal/features/work/editors/spreadsheet-editor';
import type { WorkOfficeFileAction } from './internal/features/work/editors/work-office-chrome';
import {
  OfficeSurface,
  type OfficeSurfaceProps,
  type OfficeTheme,
} from './office-surface';

const PDFIUM_WASM_FILE_NAME = 'pdfium.wasm';

export const defaultPdfiumWasmUrl = siblingAssetUrl(
  import.meta.url,
  PDFIUM_WASM_FILE_NAME,
);

function siblingAssetUrl(moduleUrl: string, fileName: string): string {
  return `${moduleUrl.slice(0, moduleUrl.lastIndexOf('/') + 1)}${fileName}`;
}

export type { WorkOfficeFileAction as OfficeFileAction };
export type { OfficeSurfaceProps, OfficeTheme };

export interface DocumentEditorProps
  extends Omit<InternalDocumentEditorProps, 'preview'>,
    OfficeSurfaceProps {
  preview?: boolean;
}

export function DocumentEditor({
  className,
  preview = false,
  style,
  theme,
  ...editorProps
}: DocumentEditorProps) {
  return (
    <OfficeSurface className={className} style={style} theme={theme}>
      <InternalDocumentEditor {...editorProps} preview={preview} />
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
      <InternalSpreadsheetEditor {...editorProps} preview={preview} />
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
      <InternalPresentationEditor {...editorProps} preview={preview} />
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
      <InternalPdfViewer {...viewerProps} wasmUrl={wasmUrl} />
    </OfficeSurface>
  );
}
