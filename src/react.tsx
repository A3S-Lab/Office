import { lazy, type ReactNode, Suspense } from 'react';
import type { WorkOfficeCollaborationSession } from './internal/collaboration/office-collaboration';
import type { WorkOfficeCollaborationPresence } from './internal/collaboration/office-collaboration-presence';
import { WorkEditorLoadingState } from './internal/features/work/components/work-editor-loading-state';
import type { DocumentEditorProps as InternalDocumentEditorProps } from './internal/features/work/editors/document-editor';
import type { MarkdownEditorProps as InternalMarkdownEditorProps } from './internal/features/work/editors/markdown-editor';
import {
  assertOfficeCollaborationPresencePairing,
  OfficeCollaborationPresenceProvider,
} from './internal/features/work/editors/office-collaboration-presence-context';
import {
  OfficeEditorErrorBoundary,
  type OfficeEditorDiagnostic,
  type OfficeEditorHostError,
  type OfficeEditorKind,
} from './internal/features/work/editors/office-editor-error-boundary';
import { OfficeEditorFocusHandoff } from './internal/features/work/editors/office-editor-focus-handoff';
import { OfficeMessagesProvider } from './internal/features/work/editors/office-messages-context';
import type { PdfViewerProps as InternalPdfViewerProps } from './internal/features/work/editors/pdf-viewer';
import {
  officeMessage,
  resolveOfficeMessages,
} from './internal/i18n/office-locale';
import type {
  OfficeLocale,
  OfficeMessageCatalog,
  OfficeMessagesOverride,
} from './internal/i18n/office-messages';

export type {
  OfficeEditorDiagnostic,
  OfficeEditorHostError,
  OfficeEditorHostErrorPhase,
  OfficeEditorKind,
} from './internal/features/work/editors/office-editor-error-boundary';
export { OfficeEditorErrorBoundary } from './internal/features/work/editors/office-editor-error-boundary';
export type {
  OfficeLocale,
  OfficeMessageCatalog,
  OfficeMessageKey,
  OfficeMessagesOverride,
} from './internal/i18n/office-messages';
export {
  OFFICE_DEFAULT_LOCALE,
  officeMessage,
  resolveOfficeLocale,
  resolveOfficeMessages,
} from './internal/i18n/office-locale';

export {
  PDF_EVIDENCE_COORDINATE_BASIS,
  type PdfEvidenceBounds,
  type PdfEvidenceOverlay,
  type PdfEvidencePage,
  type PdfEvidenceRegion,
  type PdfEvidenceRegionLocation,
} from './internal/features/work/editors/pdf-evidence-contract';
export type { PdfPageOrganizationExport } from './internal/features/work/editors/use-pdf-page-organization';
export type {
  WorkDocumentMailMergeFilterOperator as DocumentMailMergeFilterOperator,
  WorkDocumentMailMergeFilterRule as DocumentMailMergeFilterRule,
  WorkDocumentMailMergeGeneratedDocument as DocumentMailMergeGeneratedDocument,
  WorkDocumentMailMergeRecipientFilter as DocumentMailMergeRecipientFilter,
  WorkDocumentMailMergeRecord as DocumentMailMergeRecord,
  WorkDocumentMailMergeSource as DocumentMailMergeSource,
} from './internal/features/work/work-document-mail-merge';
export {
  activeMailMergeRecord,
  createMailMergeFieldContextResolver,
  emptyMailMergeSource,
  filteredMailMergeRecords,
  generateMailMergeDocuments,
  mailMergeFieldNamesFromEditor,
  mailMergeFieldNamesFromHtml,
  mailMergeFieldNamesFromRecords,
  mailMergeRecordMatchesFilter,
  normalizeMailMergeRecipientFilter,
  normalizeMailMergeRecord,
  normalizeMailMergeSource,
  previewMailMergeSource,
  setMailMergeRecipientFilter,
  stepMailMergeSource,
} from './internal/features/work/work-document-mail-merge';

import type { PresentationEditorProps as InternalPresentationEditorProps } from './internal/features/work/editors/presentation-editor';
import type { SpreadsheetEditorProps as InternalSpreadsheetEditorProps } from './internal/features/work/editors/spreadsheet-editor';
import type { WorkOfficeFileAction } from './internal/features/work/editors/work-office-chrome';
import {
  OFFICE_DOCUMENT_LAYOUT_ARABIC_FONT_FAMILY,
  OFFICE_DOCUMENT_LAYOUT_ARABIC_FONT_ID,
  OFFICE_DOCUMENT_LAYOUT_FONT_FAMILY,
  OFFICE_DOCUMENT_LAYOUT_FONT_ID,
  OFFICE_DOCUMENT_LAYOUT_HEBREW_FONT_FAMILY,
  OFFICE_DOCUMENT_LAYOUT_HEBREW_FONT_ID,
  OFFICE_DOCUMENT_LAYOUT_LATIN_FONT_FAMILY,
  OFFICE_DOCUMENT_LAYOUT_LATIN_FONT_ID,
} from './internal/features/work/work-document-fonts';
import {
  OfficeSurface,
  type OfficeSurfaceProps,
  type OfficeTheme,
} from './office-surface';

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

const officeEditorLoaders: Record<OfficeEditorKind, () => Promise<unknown>> = {
  document: loadDocumentEditor,
  markdown: loadMarkdownEditor,
  spreadsheet: loadSpreadsheetEditor,
  presentation: loadPresentationEditor,
  pdf: loadPdfViewer,
};
const officeRuntimeAssetPreloads = new Map<string, Promise<void>>();

export interface OfficeEditorPreloadOptions {
  /** Overrides the packaged PDFium asset when preloading the PDF runtime. */
  pdfWasmUrl?: string;
  /** Also fetches PDFium WebAssembly. Disabled unless explicitly enabled. */
  preloadRuntimeAssets?: boolean;
}

/**
 * Starts loading one editor without mounting it.
 *
 * Call this from an intent signal such as hover or keyboard focus to keep the
 * initial application bundle small without adding latency to editor opening.
 */
export async function preloadOfficeEditor(
  kind: OfficeEditorKind,
  options: OfficeEditorPreloadOptions = {},
): Promise<void> {
  const editor = officeEditorLoaders[kind]();
  if (kind !== 'pdf' || options.preloadRuntimeAssets !== true) {
    await editor;
    return;
  }
  await Promise.all([
    editor,
    preloadOfficeRuntimeAsset(options.pdfWasmUrl ?? defaultPdfiumWasmUrl),
  ]);
}

async function preloadOfficeRuntimeAsset(url: string): Promise<void> {
  if (typeof document === 'undefined' || typeof fetch !== 'function') return;
  const cached = officeRuntimeAssetPreloads.get(url);
  if (cached) return cached;
  const preload = Promise.resolve()
    .then(() =>
      fetch(url, {
        cache: 'force-cache',
        credentials: 'same-origin',
      }),
    )
    .then(async (response) => {
      if (!response.ok) {
        throw new Error(`Unable to preload the Office runtime asset: ${url}`);
      }
      await response.arrayBuffer();
    })
    .catch(() => {
      officeRuntimeAssetPreloads.delete(url);
    });
  officeRuntimeAssetPreloads.set(url, preload);
  await preload;
}

export const defaultPdfiumWasmUrl = new URL(
  /* webpackIgnore: true */ './pdfium.wasm',
  import.meta.url,
).href;
export const defaultOfficeKernelWasmUrl = new URL(
  /* webpackIgnore: true */ './office-kernel.wasm',
  import.meta.url,
).href;
export const defaultDocumentLayoutFontUrl = new URL(
  /* webpackIgnore: true */ './noto-sans-hans-regular.otf',
  import.meta.url,
).href;
export const defaultDocumentLatinLayoutFontUrl = new URL(
  /* webpackIgnore: true */ './noto-sans-regular.ttf',
  import.meta.url,
).href;
export const defaultDocumentArabicLayoutFontUrl = new URL(
  /* webpackIgnore: true */ './noto-naskh-arabic-regular.ttf',
  import.meta.url,
).href;
export const defaultDocumentHebrewLayoutFontUrl = new URL(
  /* webpackIgnore: true */ './noto-sans-hebrew-regular.ttf',
  import.meta.url,
).href;
export interface DocumentLayoutFont {
  id: string;
  family: string;
  url: string;
  weight?: number;
  style?: 'normal' | 'italic';
}

export const defaultDocumentLayoutFonts: readonly DocumentLayoutFont[] = [
  {
    id: OFFICE_DOCUMENT_LAYOUT_LATIN_FONT_ID,
    family: OFFICE_DOCUMENT_LAYOUT_LATIN_FONT_FAMILY,
    url: defaultDocumentLatinLayoutFontUrl,
    weight: 400,
    style: 'normal',
  },
  {
    id: OFFICE_DOCUMENT_LAYOUT_FONT_ID,
    family: OFFICE_DOCUMENT_LAYOUT_FONT_FAMILY,
    url: defaultDocumentLayoutFontUrl,
    weight: 400,
    style: 'normal',
  },
  {
    id: OFFICE_DOCUMENT_LAYOUT_ARABIC_FONT_ID,
    family: OFFICE_DOCUMENT_LAYOUT_ARABIC_FONT_FAMILY,
    url: defaultDocumentArabicLayoutFontUrl,
    weight: 400,
    style: 'normal',
  },
  {
    id: OFFICE_DOCUMENT_LAYOUT_HEBREW_FONT_ID,
    family: OFFICE_DOCUMENT_LAYOUT_HEBREW_FONT_FAMILY,
    url: defaultDocumentHebrewLayoutFontUrl,
    weight: 400,
    style: 'normal',
  },
];

export type {
  OfficeSurfaceProps,
  OfficeTheme,
  WorkOfficeFileAction as OfficeFileAction,
};

function OfficeEditorLoader({
  children,
  collaboration,
  kind,
  messages,
  onDiagnostic,
  onError,
  presence,
  title,
}: {
  children: ReactNode;
  collaboration?: WorkOfficeCollaborationSession;
  kind: OfficeEditorKind;
  messages: OfficeMessageCatalog;
  onDiagnostic?: (diagnostic: OfficeEditorDiagnostic) => void;
  onError?: (error: OfficeEditorHostError) => void;
  presence?: WorkOfficeCollaborationPresence;
  title: string;
}) {
  assertOfficeCollaborationPresencePairing({
    expectedKind: kind,
    presence,
    session: collaboration,
  });
  return (
    <OfficeCollaborationPresenceProvider presence={presence}>
      <OfficeMessagesProvider catalog={messages}>
        <OfficeEditorFocusHandoff>
          <OfficeEditorErrorBoundary
            editor={kind}
            messages={messages}
            onDiagnostic={onDiagnostic}
            onError={onError}
            title={title}
          >
            <Suspense fallback={<WorkEditorLoadingState title={title} />}>
              {children}
            </Suspense>
          </OfficeEditorErrorBoundary>
        </OfficeEditorFocusHandoff>
      </OfficeMessagesProvider>
    </OfficeCollaborationPresenceProvider>
  );
}

interface OfficeCollaborationSurfaceProps {
  /** Host-owned Awareness projection paired with the editor collaboration session. */
  presence?: WorkOfficeCollaborationPresence;
}

interface OfficeEditorObservabilityProps {
  /** Fires once per render/chunk failure caught by the host error boundary. */
  onError?: (error: OfficeEditorHostError) => void;
  /** Broader diagnostic sink for host observability (errors, later marks). */
  onDiagnostic?: (diagnostic: OfficeEditorDiagnostic) => void;
}

interface OfficeEditorLocaleProps {
  /** BCP 47 UI locale. Defaults to `zh-CN`; `en-US` is the first alternate. */
  locale?: OfficeLocale | string;
  /** Partial message override merged on top of the resolved locale catalog. */
  messages?: OfficeMessagesOverride;
}

const editorLoadingMessageKey = {
  document: 'editor.loading.document',
  markdown: 'editor.loading.markdown',
  spreadsheet: 'editor.loading.spreadsheet',
  presentation: 'editor.loading.presentation',
  pdf: 'editor.loading.pdf',
} as const satisfies Record<
  OfficeEditorKind,
  import('./internal/i18n/office-messages').OfficeMessageKey
>;

export interface DocumentEditorProps
  extends Omit<
      InternalDocumentEditorProps,
      'defaultRibbonCollapsed' | 'layoutFonts' | 'preview'
    >,
    OfficeSurfaceProps,
    OfficeCollaborationSurfaceProps,
    OfficeEditorObservabilityProps,
    OfficeEditorLocaleProps {
  preview?: boolean;
  defaultRibbonCollapsed?: boolean;
  layoutFonts?: readonly DocumentLayoutFont[];
}

export function DocumentEditor({
  className,
  defaultRibbonCollapsed = false,
  kernelWasmUrl = defaultOfficeKernelWasmUrl,
  layoutFonts = defaultDocumentLayoutFonts,
  locale,
  messages: messagesOverride,
  onDiagnostic,
  onError,
  presence,
  preview = false,
  style,
  theme,
  ...editorProps
}: DocumentEditorProps) {
  const messages = resolveOfficeMessages({
    locale,
    messages: messagesOverride,
  });
  return (
    <OfficeSurface className={className} style={style} theme={theme}>
      <OfficeEditorLoader
        collaboration={editorProps.collaboration}
        kind="document"
        messages={messages}
        onDiagnostic={onDiagnostic}
        onError={onError}
        presence={presence}
        title={officeMessage(messages, editorLoadingMessageKey.document)}
      >
        <LazyDocumentEditor
          {...editorProps}
          kernelWasmUrl={kernelWasmUrl}
          layoutFonts={layoutFonts}
          preview={preview}
          defaultRibbonCollapsed={defaultRibbonCollapsed}
        />
      </OfficeEditorLoader>
    </OfficeSurface>
  );
}

export interface MarkdownEditorProps
  extends Omit<InternalMarkdownEditorProps, 'preview'>,
    OfficeSurfaceProps,
    OfficeCollaborationSurfaceProps,
    OfficeEditorObservabilityProps,
    OfficeEditorLocaleProps {
  preview?: boolean;
}

export function MarkdownEditor({
  className,
  locale,
  messages: messagesOverride,
  onDiagnostic,
  onError,
  presence,
  preview = false,
  style,
  theme,
  ...editorProps
}: MarkdownEditorProps) {
  const messages = resolveOfficeMessages({
    locale,
    messages: messagesOverride,
  });
  return (
    <OfficeSurface className={className} style={style} theme={theme}>
      <OfficeEditorLoader
        collaboration={editorProps.collaboration}
        kind="markdown"
        messages={messages}
        onDiagnostic={onDiagnostic}
        onError={onError}
        presence={presence}
        title={officeMessage(messages, editorLoadingMessageKey.markdown)}
      >
        <LazyMarkdownEditor {...editorProps} preview={preview} />
      </OfficeEditorLoader>
    </OfficeSurface>
  );
}

export interface SpreadsheetEditorProps
  extends Omit<InternalSpreadsheetEditorProps, 'preview'>,
    OfficeSurfaceProps,
    OfficeCollaborationSurfaceProps,
    OfficeEditorObservabilityProps,
    OfficeEditorLocaleProps {
  preview?: boolean;
}

export function SpreadsheetEditor({
  className,
  kernelWasmUrl = defaultOfficeKernelWasmUrl,
  locale,
  messages: messagesOverride,
  onDiagnostic,
  onError,
  presence,
  preview = false,
  style,
  theme,
  ...editorProps
}: SpreadsheetEditorProps) {
  const messages = resolveOfficeMessages({
    locale,
    messages: messagesOverride,
  });
  return (
    <OfficeSurface className={className} style={style} theme={theme}>
      <OfficeEditorLoader
        collaboration={editorProps.collaboration}
        kind="spreadsheet"
        messages={messages}
        onDiagnostic={onDiagnostic}
        onError={onError}
        presence={presence}
        title={officeMessage(messages, editorLoadingMessageKey.spreadsheet)}
      >
        <LazySpreadsheetEditor
          {...editorProps}
          kernelWasmUrl={kernelWasmUrl}
          preview={preview}
        />
      </OfficeEditorLoader>
    </OfficeSurface>
  );
}

export interface PresentationEditorProps
  extends Omit<InternalPresentationEditorProps, 'preview'>,
    OfficeSurfaceProps,
    OfficeCollaborationSurfaceProps,
    OfficeEditorObservabilityProps,
    OfficeEditorLocaleProps {
  preview?: boolean;
}

export function PresentationEditor({
  className,
  kernelWasmUrl = defaultOfficeKernelWasmUrl,
  locale,
  messages: messagesOverride,
  onDiagnostic,
  onError,
  presence,
  preview = false,
  style,
  theme,
  ...editorProps
}: PresentationEditorProps) {
  const messages = resolveOfficeMessages({
    locale,
    messages: messagesOverride,
  });
  return (
    <OfficeSurface className={className} style={style} theme={theme}>
      <OfficeEditorLoader
        collaboration={editorProps.collaboration}
        kind="presentation"
        messages={messages}
        onDiagnostic={onDiagnostic}
        onError={onError}
        presence={presence}
        title={officeMessage(messages, editorLoadingMessageKey.presentation)}
      >
        <LazyPresentationEditor
          {...editorProps}
          kernelWasmUrl={kernelWasmUrl}
          preview={preview}
        />
      </OfficeEditorLoader>
    </OfficeSurface>
  );
}

export interface PdfViewerProps
  extends InternalPdfViewerProps,
    OfficeSurfaceProps,
    OfficeCollaborationSurfaceProps,
    OfficeEditorObservabilityProps,
    OfficeEditorLocaleProps {}

export function PdfViewer({
  className,
  locale,
  messages: messagesOverride,
  onDiagnostic,
  onError,
  presence,
  style,
  theme,
  wasmUrl = defaultPdfiumWasmUrl,
  ...viewerProps
}: PdfViewerProps) {
  const messages = resolveOfficeMessages({
    locale,
    messages: messagesOverride,
  });
  return (
    <OfficeSurface className={className} style={style} theme={theme}>
      <OfficeEditorLoader
        collaboration={viewerProps.collaboration}
        kind="pdf"
        messages={messages}
        onDiagnostic={onDiagnostic}
        onError={onError}
        presence={presence}
        title={officeMessage(messages, editorLoadingMessageKey.pdf)}
      >
        <LazyPdfViewer {...viewerProps} wasmUrl={wasmUrl} />
      </OfficeEditorLoader>
    </OfficeSurface>
  );
}
