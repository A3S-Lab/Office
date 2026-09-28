import type { Extensions } from '@tiptap/core';
import { createElement, type ReactNode } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import {
  defineComponent,
  h,
  onBeforeUnmount,
  onMounted,
  onUpdated,
  type PropType,
  ref,
} from 'vue';
import type {
  DocumentContent,
  DocumentReviewConflictEvent,
  EditorAgentRequest,
  GetDocumentSelectionMenuItems,
  GetMarkdownSelectionMenuItems,
  MarkdownContent,
  OfficeCollaborationPresence,
  OfficeCollaborationSession,
  PdfCollaborationContent,
  PdfPageOrganizationExport,
  PresentationContent,
  SpreadsheetContent,
  SpreadsheetSortCustomListStore,
} from './core';
import type { OfficeTheme } from './office-surface';
import {
  type DocumentLayoutFont,
  type OfficeEditorDiagnostic,
  type OfficeEditorHostError,
  type OfficeFileAction,
  type OfficeLocale,
  type OfficeMessagesOverride,
  type PdfEvidenceOverlay,
  type PdfEvidenceRegion,
  DocumentEditor as ReactDocumentEditor,
  MarkdownEditor as ReactMarkdownEditor,
  PdfViewer as ReactPdfViewer,
  PresentationEditor as ReactPresentationEditor,
  SpreadsheetEditor as ReactSpreadsheetEditor,
} from './react';

function createReactRenderer(renderNode: () => ReactNode) {
  const host = ref<HTMLDivElement | null>(null);
  let root: Root | null = null;

  const render = () => {
    if (!host.value) return;
    root ??= createRoot(host.value);
    root.render(renderNode());
  };

  onMounted(render);
  onUpdated(render);
  onBeforeUnmount(() => {
    root?.unmount();
    root = null;
  });

  return () =>
    h('div', {
      class: 'a3s-office-vue-host',
      ref: host,
      style: { height: '100%', minHeight: 0, minWidth: 0, width: '100%' },
    });
}

const themeProp = {
  default: 'system',
  type: String as PropType<OfficeTheme>,
} as const;

const fileActionsProp = {
  default: undefined,
  type: Array as PropType<readonly OfficeFileAction[]>,
};

const localeProp = {
  default: undefined,
  type: String as PropType<OfficeLocale | string>,
};

const messagesProp = {
  default: undefined,
  type: Object as PropType<OfficeMessagesOverride>,
};

const officeLocaleProps = {
  locale: localeProp,
  messages: messagesProp,
} as const;

const officeObservabilityEmits = {
  diagnostic: (_diagnostic: OfficeEditorDiagnostic) => true,
  error: (_error: OfficeEditorHostError) => true,
} as const;

function officeObservabilityBindings(
  props: {
    locale?: OfficeLocale | string;
    messages?: OfficeMessagesOverride;
  },
  emit: (event: string, payload?: unknown) => void,
) {
  return {
    locale: props.locale,
    messages: props.messages,
    onDiagnostic: (diagnostic: OfficeEditorDiagnostic) =>
      emit('diagnostic', diagnostic),
    onError: (error: OfficeEditorHostError) => emit('error', error),
  };
}

export const DocumentEditor = defineComponent({
  name: 'A3SDocumentEditor',
  props: {
    artifactId: String,
    collaboration: Object as PropType<OfficeCollaborationSession>,
    presence: Object as PropType<OfficeCollaborationPresence>,
    content: {
      required: true,
      type: Object as PropType<DocumentContent>,
    },
    fileActions: fileActionsProp,
    extensions: {
      default: undefined,
      type: Array as PropType<Extensions>,
    },
    getSelectionMenuItems: Function as PropType<GetDocumentSelectionMenuItems>,
    kernelWasmUrl: String,
    layoutFonts: {
      default: undefined,
      type: Array as PropType<readonly DocumentLayoutFont[]>,
    },
    preview: {
      default: false,
      type: Boolean,
    },
    saveStatus: String,
    theme: themeProp,
    ...officeLocaleProps,
  },
  emits: {
    agentRequest: (_request: EditorAgentRequest) => true,
    change: (_content: DocumentContent) => true,
    reviewConflict: (_event: DocumentReviewConflictEvent) => true,
    'update:content': (_content: DocumentContent) => true,
    ...officeObservabilityEmits,
  },
  setup(props, { emit }) {
    return createReactRenderer(() =>
      createElement(ReactDocumentEditor, {
        artifactId: props.artifactId,
        collaboration: props.collaboration,
        presence: props.presence,
        content: props.content,
        extensions: props.extensions,
        fileActions: props.fileActions,
        getSelectionMenuItems: props.getSelectionMenuItems,
        kernelWasmUrl: props.kernelWasmUrl,
        layoutFonts: props.layoutFonts,
        onAgentRequest: (request) => emit('agentRequest', request),
        onReviewConflict: (event) => emit('reviewConflict', event),
        onChange: (content) => {
          emit('update:content', content);
          emit('change', content);
        },
        preview: props.preview,
        saveStatus: props.saveStatus,
        theme: props.theme,
        ...officeObservabilityBindings(
          props,
          emit as (event: string, payload?: unknown) => void,
        ),
      }),
    );
  },
});

export const MarkdownEditor = defineComponent({
  name: 'A3SMarkdownEditor',
  props: {
    collaboration: Object as PropType<OfficeCollaborationSession>,
    presence: Object as PropType<OfficeCollaborationPresence>,
    content: {
      required: true,
      type: Object as PropType<MarkdownContent>,
    },
    fileActions: fileActionsProp,
    extensions: {
      default: undefined,
      type: Array as PropType<Extensions>,
    },
    getSelectionMenuItems: Function as PropType<GetMarkdownSelectionMenuItems>,
    preview: {
      default: false,
      type: Boolean,
    },
    saveStatus: String,
    theme: themeProp,
    ...officeLocaleProps,
  },
  emits: {
    change: (_content: MarkdownContent) => true,
    'update:content': (_content: MarkdownContent) => true,
    ...officeObservabilityEmits,
  },
  setup(props, { emit }) {
    return createReactRenderer(() =>
      createElement(ReactMarkdownEditor, {
        collaboration: props.collaboration,
        presence: props.presence,
        content: props.content,
        extensions: props.extensions,
        fileActions: props.fileActions,
        getSelectionMenuItems: props.getSelectionMenuItems,
        onChange: (content) => {
          emit('update:content', content);
          emit('change', content);
        },
        preview: props.preview,
        saveStatus: props.saveStatus,
        theme: props.theme,
        ...officeObservabilityBindings(
          props,
          emit as (event: string, payload?: unknown) => void,
        ),
      }),
    );
  },
});

export const SpreadsheetEditor = defineComponent({
  name: 'A3SSpreadsheetEditor',
  props: {
    collaboration: Object as PropType<OfficeCollaborationSession>,
    presence: Object as PropType<OfficeCollaborationPresence>,
    content: {
      required: true,
      type: Object as PropType<SpreadsheetContent>,
    },
    fileActions: fileActionsProp,
    kernelWasmUrl: String,
    preview: {
      default: false,
      type: Boolean,
    },
    saveStatus: String,
    sortCustomListStore: Object as PropType<SpreadsheetSortCustomListStore>,
    theme: themeProp,
    ...officeLocaleProps,
  },
  emits: {
    agentRequest: (_request: EditorAgentRequest) => true,
    change: (_content: SpreadsheetContent) => true,
    'update:content': (_content: SpreadsheetContent) => true,
    ...officeObservabilityEmits,
  },
  setup(props, { emit }) {
    return createReactRenderer(() =>
      createElement(ReactSpreadsheetEditor, {
        collaboration: props.collaboration,
        presence: props.presence,
        content: props.content,
        fileActions: props.fileActions,
        kernelWasmUrl: props.kernelWasmUrl,
        onAgentRequest: (request) => emit('agentRequest', request),
        onChange: (content) => {
          emit('update:content', content);
          emit('change', content);
        },
        preview: props.preview,
        saveStatus: props.saveStatus,
        sortCustomListStore: props.sortCustomListStore,
        theme: props.theme,
        ...officeObservabilityBindings(
          props,
          emit as (event: string, payload?: unknown) => void,
        ),
      }),
    );
  },
});

export const PresentationEditor = defineComponent({
  name: 'A3SPresentationEditor',
  props: {
    collaboration: Object as PropType<OfficeCollaborationSession>,
    presence: Object as PropType<OfficeCollaborationPresence>,
    content: {
      required: true,
      type: Object as PropType<PresentationContent>,
    },
    fileActions: fileActionsProp,
    kernelWasmUrl: String,
    preview: {
      default: false,
      type: Boolean,
    },
    saveStatus: String,
    theme: themeProp,
    ...officeLocaleProps,
  },
  emits: {
    agentRequest: (_request: EditorAgentRequest) => true,
    change: (_content: PresentationContent) => true,
    startSlideshow: () => true,
    'update:content': (_content: PresentationContent) => true,
    ...officeObservabilityEmits,
  },
  setup(props, { emit }) {
    return createReactRenderer(() =>
      createElement(ReactPresentationEditor, {
        collaboration: props.collaboration,
        presence: props.presence,
        content: props.content,
        fileActions: props.fileActions,
        kernelWasmUrl: props.kernelWasmUrl,
        onAgentRequest: (request) => emit('agentRequest', request),
        onChange: (content) => {
          emit('update:content', content);
          emit('change', content);
        },
        onStartSlideshow: () => emit('startSlideshow'),
        preview: props.preview,
        saveStatus: props.saveStatus,
        theme: props.theme,
        ...officeObservabilityBindings(
          props,
          emit as (event: string, payload?: unknown) => void,
        ),
      }),
    );
  },
});

export const PdfViewer = defineComponent({
  name: 'A3SPdfViewer',
  props: {
    collaboration: Object as PropType<OfficeCollaborationSession>,
    evidenceOverlay: Object as PropType<PdfEvidenceOverlay>,
    presence: Object as PropType<OfficeCollaborationPresence>,
    fileName: String,
    loadSource: {
      required: true,
      type: Function as PropType<() => Promise<Blob>>,
    },
    onSave: Function as PropType<(pdf: Blob) => Promise<boolean>>,
    onPageExport: Function as PropType<
      (
        files: readonly PdfPageOrganizationExport[],
      ) => boolean | Promise<boolean>
    >,
    saveLabel: String,
    selectedEvidenceRegionId: String,
    sourceKey: String,
    theme: themeProp,
    wasmUrl: String,
    worker: {
      default: undefined,
      type: Boolean as PropType<boolean | undefined>,
    },
    ...officeLocaleProps,
  },
  emits: {
    collaborationChange: (_content: PdfCollaborationContent) => true,
    evidenceRegionSelect: (_region: PdfEvidenceRegion) => true,
    pageChange: (_pageNumber: number) => true,
    ...officeObservabilityEmits,
  },
  setup(props, { emit }) {
    return createReactRenderer(() =>
      createElement(ReactPdfViewer, {
        collaboration: props.collaboration,
        evidenceOverlay: props.evidenceOverlay,
        presence: props.presence,
        fileName: props.fileName,
        loadSource: props.loadSource,
        onCollaborationChange: (content) =>
          emit('collaborationChange', content),
        onEvidenceRegionSelect: (region) =>
          emit('evidenceRegionSelect', region),
        onPageChange: (pageNumber) => emit('pageChange', pageNumber),
        onPageExport: props.onPageExport,
        onSave: props.onSave,
        saveLabel: props.saveLabel,
        selectedEvidenceRegionId: props.selectedEvidenceRegionId,
        sourceKey: props.sourceKey,
        theme: props.theme,
        wasmUrl: props.wasmUrl,
        worker: props.worker,
        ...officeObservabilityBindings(
          props,
          emit as (event: string, payload?: unknown) => void,
        ),
      }),
    );
  },
});

export {
  DocumentEditor as A3SDocumentEditor,
  MarkdownEditor as A3SMarkdownEditor,
  PdfViewer as A3SPdfViewer,
  PresentationEditor as A3SPresentationEditor,
  SpreadsheetEditor as A3SSpreadsheetEditor,
};
