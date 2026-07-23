import { createElement, type ReactNode } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import {
  DocumentEditor as ReactDocumentEditor,
  MarkdownEditor as ReactMarkdownEditor,
  type OfficeFileAction,
  PdfViewer as ReactPdfViewer,
  PresentationEditor as ReactPresentationEditor,
  SpreadsheetEditor as ReactSpreadsheetEditor,
} from './react';
import type {
  DocumentContent,
  EditorAgentRequest,
  MarkdownContent,
  PresentationContent,
  SpreadsheetContent,
} from './core';
import {
  defineComponent,
  h,
  onBeforeUnmount,
  onMounted,
  onUpdated,
  ref,
  type PropType,
} from 'vue';
import type { OfficeTheme } from './office-surface';

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

export const DocumentEditor = defineComponent({
  name: 'A3SDocumentEditor',
  props: {
    content: {
      required: true,
      type: Object as PropType<DocumentContent>,
    },
    fileActions: fileActionsProp,
    kernelWasmUrl: String,
    preview: {
      default: false,
      type: Boolean,
    },
    saveStatus: String,
    theme: themeProp,
  },
  emits: {
    agentRequest: (_request: EditorAgentRequest) => true,
    change: (_content: DocumentContent) => true,
    'update:content': (_content: DocumentContent) => true,
  },
  setup(props, { emit }) {
    return createReactRenderer(() =>
      createElement(ReactDocumentEditor, {
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
        theme: props.theme,
      }),
    );
  },
});

export const MarkdownEditor = defineComponent({
  name: 'A3SMarkdownEditor',
  props: {
    content: {
      required: true,
      type: Object as PropType<MarkdownContent>,
    },
    fileActions: fileActionsProp,
    preview: {
      default: false,
      type: Boolean,
    },
    saveStatus: String,
    theme: themeProp,
  },
  emits: {
    change: (_content: MarkdownContent) => true,
    'update:content': (_content: MarkdownContent) => true,
  },
  setup(props, { emit }) {
    return createReactRenderer(() =>
      createElement(ReactMarkdownEditor, {
        content: props.content,
        fileActions: props.fileActions,
        onChange: (content) => {
          emit('update:content', content);
          emit('change', content);
        },
        preview: props.preview,
        saveStatus: props.saveStatus,
        theme: props.theme,
      }),
    );
  },
});

export const SpreadsheetEditor = defineComponent({
  name: 'A3SSpreadsheetEditor',
  props: {
    content: {
      required: true,
      type: Object as PropType<SpreadsheetContent>,
    },
    fileActions: fileActionsProp,
    preview: {
      default: false,
      type: Boolean,
    },
    saveStatus: String,
    theme: themeProp,
  },
  emits: {
    agentRequest: (_request: EditorAgentRequest) => true,
    change: (_content: SpreadsheetContent) => true,
    'update:content': (_content: SpreadsheetContent) => true,
  },
  setup(props, { emit }) {
    return createReactRenderer(() =>
      createElement(ReactSpreadsheetEditor, {
        content: props.content,
        fileActions: props.fileActions,
        onAgentRequest: (request) => emit('agentRequest', request),
        onChange: (content) => {
          emit('update:content', content);
          emit('change', content);
        },
        preview: props.preview,
        saveStatus: props.saveStatus,
        theme: props.theme,
      }),
    );
  },
});

export const PresentationEditor = defineComponent({
  name: 'A3SPresentationEditor',
  props: {
    content: {
      required: true,
      type: Object as PropType<PresentationContent>,
    },
    fileActions: fileActionsProp,
    preview: {
      default: false,
      type: Boolean,
    },
    saveStatus: String,
    theme: themeProp,
  },
  emits: {
    agentRequest: (_request: EditorAgentRequest) => true,
    change: (_content: PresentationContent) => true,
    startSlideshow: () => true,
    'update:content': (_content: PresentationContent) => true,
  },
  setup(props, { emit }) {
    return createReactRenderer(() =>
      createElement(ReactPresentationEditor, {
        content: props.content,
        fileActions: props.fileActions,
        onAgentRequest: (request) => emit('agentRequest', request),
        onChange: (content) => {
          emit('update:content', content);
          emit('change', content);
        },
        onStartSlideshow: () => emit('startSlideshow'),
        preview: props.preview,
        saveStatus: props.saveStatus,
        theme: props.theme,
      }),
    );
  },
});

export const PdfViewer = defineComponent({
  name: 'A3SPdfViewer',
  props: {
    fileName: String,
    loadSource: {
      required: true,
      type: Function as PropType<() => Promise<Blob>>,
    },
    onSave: Function as PropType<(pdf: Blob) => Promise<boolean>>,
    saveLabel: String,
    sourceKey: String,
    theme: themeProp,
    wasmUrl: String,
  },
  setup(props) {
    return createReactRenderer(() =>
      createElement(ReactPdfViewer, {
        fileName: props.fileName,
        loadSource: props.loadSource,
        onSave: props.onSave,
        saveLabel: props.saveLabel,
        sourceKey: props.sourceKey,
        theme: props.theme,
        wasmUrl: props.wasmUrl,
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
