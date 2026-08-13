import type { Extensions } from '@tiptap/core';
import { createElement, type ReactNode } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import type {
  DocumentContent,
  DocumentReviewConflictEvent,
  EditorAgentRequest,
  GetDocumentSelectionMenuItems,
  GetMarkdownSelectionMenuItems,
  MarkdownContent,
  OfficeCollaborationSession,
  PresentationContent,
  SpreadsheetContent,
} from './core';
import type { OfficeTheme } from './office-surface';
import {
  DocumentEditor,
  type DocumentLayoutFont,
  MarkdownEditor,
  type OfficeFileAction,
  PdfViewer,
  PresentationEditor,
  SpreadsheetEditor,
} from './react';

const HTMLElementBase =
  typeof HTMLElement === 'undefined'
    ? (class {} as unknown as typeof HTMLElement)
    : HTMLElement;

abstract class A3SOfficeElement extends HTMLElementBase {
  #root: Root | null = null;

  connectedCallback() {
    this.renderReactTree();
  }

  disconnectedCallback() {
    this.#root?.unmount();
    this.#root = null;
  }

  attributeChangedCallback() {
    this.renderReactTree();
  }

  protected requestRender() {
    this.renderReactTree();
  }

  protected abstract editorNode(): ReactNode;

  private renderReactTree() {
    if (!this.isConnected) return;
    this.#root ??= createRoot(this);
    this.#root.render(this.editorNode());
  }
}

function themeFrom(element: Element): OfficeTheme {
  const value = element.getAttribute('theme');
  return value === 'dark' || value === 'light' ? value : 'system';
}

function missingContent(kind: string, theme: OfficeTheme): ReactNode {
  return createElement(
    'div',
    {
      className: 'a3s-office a3s-office-empty',
      'data-a3s-office': '',
      'data-theme': theme,
      role: 'status',
    },
    `Set the element's content property to a ${kind} content model.`,
  );
}

function dispatchDetail<T>(target: EventTarget, type: string, detail: T): void {
  target.dispatchEvent(
    new CustomEvent<T>(type, {
      bubbles: true,
      composed: true,
      detail,
    }),
  );
}

abstract class A3SContentEditorElement<
  TContent extends
    | DocumentContent
    | MarkdownContent
    | SpreadsheetContent
    | PresentationContent,
> extends A3SOfficeElement {
  #content: TContent | undefined;
  #fileActions: readonly OfficeFileAction[] | undefined;

  get content(): TContent | undefined {
    return this.#content;
  }

  set content(value: TContent | undefined) {
    this.#content = value;
    this.requestRender();
  }

  get fileActions(): readonly OfficeFileAction[] | undefined {
    return this.#fileActions;
  }

  set fileActions(value: readonly OfficeFileAction[] | undefined) {
    this.#fileActions = value;
    this.requestRender();
  }

  get preview(): boolean {
    return this.hasAttribute('preview');
  }

  set preview(value: boolean) {
    this.toggleAttribute('preview', value);
  }

  get saveStatus(): string | undefined {
    return this.getAttribute('save-status') ?? undefined;
  }

  set saveStatus(value: string | undefined) {
    if (value === undefined) this.removeAttribute('save-status');
    else this.setAttribute('save-status', value);
  }

  get theme(): OfficeTheme {
    return themeFrom(this);
  }

  set theme(value: OfficeTheme) {
    this.setAttribute('theme', value);
  }

  protected changeContent(value: TContent): void {
    this.#content = value;
    dispatchDetail(this, 'change', value);
    this.requestRender();
  }

  protected requestAgent(request: EditorAgentRequest): void {
    dispatchDetail(this, 'agent-request', request);
  }
}

export class A3SDocumentEditorElement extends A3SContentEditorElement<DocumentContent> {
  #collaboration: OfficeCollaborationSession | undefined;
  #extensions: Extensions | undefined;
  #getSelectionMenuItems: GetDocumentSelectionMenuItems | undefined;
  #layoutFonts: readonly DocumentLayoutFont[] | undefined;

  static get observedAttributes() {
    return [
      'artifact-id',
      'kernel-wasm-url',
      'preview',
      'save-status',
      'theme',
    ];
  }

  get artifactId(): string | undefined {
    return this.getAttribute('artifact-id') ?? undefined;
  }

  get collaboration(): OfficeCollaborationSession | undefined {
    return this.#collaboration;
  }

  set collaboration(value: OfficeCollaborationSession | undefined) {
    this.#collaboration = value;
    this.requestRender();
  }

  set artifactId(value: string | undefined) {
    if (value === undefined) this.removeAttribute('artifact-id');
    else this.setAttribute('artifact-id', value);
  }

  get kernelWasmUrl(): string | undefined {
    return this.getAttribute('kernel-wasm-url') ?? undefined;
  }

  set kernelWasmUrl(value: string | undefined) {
    if (value === undefined) this.removeAttribute('kernel-wasm-url');
    else this.setAttribute('kernel-wasm-url', value);
  }

  get extensions(): Extensions | undefined {
    return this.#extensions;
  }

  set extensions(value: Extensions | undefined) {
    this.#extensions = value;
    this.requestRender();
  }

  get getSelectionMenuItems(): GetDocumentSelectionMenuItems | undefined {
    return this.#getSelectionMenuItems;
  }

  set getSelectionMenuItems(value: GetDocumentSelectionMenuItems | undefined,) {
    this.#getSelectionMenuItems = value;
    this.requestRender();
  }

  get layoutFonts(): readonly DocumentLayoutFont[] | undefined {
    return this.#layoutFonts;
  }

  set layoutFonts(value: readonly DocumentLayoutFont[] | undefined) {
    this.#layoutFonts = value;
    this.requestRender();
  }

  protected editorNode(): ReactNode {
    if (!this.content) return missingContent('document', this.theme);
    return createElement(DocumentEditor, {
      artifactId: this.artifactId,
      collaboration: this.collaboration,
      content: this.content,
      extensions: this.extensions,
      fileActions: this.fileActions,
      getSelectionMenuItems: this.getSelectionMenuItems,
      kernelWasmUrl: this.kernelWasmUrl,
      layoutFonts: this.layoutFonts,
      onAgentRequest: (request) => this.requestAgent(request),
      onChange: (content) => this.changeContent(content),
      onReviewConflict: (event: DocumentReviewConflictEvent) =>
        dispatchDetail(this, 'review-conflict', event),
      preview: this.preview,
      saveStatus: this.saveStatus,
      theme: this.theme,
    });
  }
}

export class A3SMarkdownEditorElement extends A3SContentEditorElement<MarkdownContent> {
  #collaboration: OfficeCollaborationSession | undefined;
  #extensions: Extensions | undefined;
  #getSelectionMenuItems: GetMarkdownSelectionMenuItems | undefined;

  static get observedAttributes() {
    return ['preview', 'save-status', 'theme'];
  }

  get collaboration(): OfficeCollaborationSession | undefined {
    return this.#collaboration;
  }

  set collaboration(value: OfficeCollaborationSession | undefined) {
    this.#collaboration = value;
    this.requestRender();
  }

  get extensions(): Extensions | undefined {
    return this.#extensions;
  }

  set extensions(value: Extensions | undefined) {
    this.#extensions = value;
    this.requestRender();
  }

  get getSelectionMenuItems(): GetMarkdownSelectionMenuItems | undefined {
    return this.#getSelectionMenuItems;
  }

  set getSelectionMenuItems(value: GetMarkdownSelectionMenuItems | undefined) {
    this.#getSelectionMenuItems = value;
    this.requestRender();
  }

  protected editorNode(): ReactNode {
    if (!this.content) return missingContent('Markdown', this.theme);
    return createElement(MarkdownEditor, {
      collaboration: this.collaboration,
      content: this.content,
      extensions: this.extensions,
      fileActions: this.fileActions,
      getSelectionMenuItems: this.getSelectionMenuItems,
      onChange: (content) => this.changeContent(content),
      preview: this.preview,
      saveStatus: this.saveStatus,
      theme: this.theme,
    });
  }
}

export class A3SSpreadsheetEditorElement extends A3SContentEditorElement<SpreadsheetContent> {
  #collaboration: OfficeCollaborationSession | undefined;

  static get observedAttributes() {
    return ['kernel-wasm-url', 'preview', 'save-status', 'theme'];
  }

  get kernelWasmUrl(): string | undefined {
    return this.getAttribute('kernel-wasm-url') ?? undefined;
  }

  get collaboration(): OfficeCollaborationSession | undefined {
    return this.#collaboration;
  }

  set collaboration(value: OfficeCollaborationSession | undefined) {
    this.#collaboration = value;
    this.requestRender();
  }

  set kernelWasmUrl(value: string | undefined) {
    if (value === undefined) this.removeAttribute('kernel-wasm-url');
    else this.setAttribute('kernel-wasm-url', value);
  }

  protected editorNode(): ReactNode {
    if (!this.content) return missingContent('spreadsheet', this.theme);
    return createElement(SpreadsheetEditor, {
      collaboration: this.collaboration,
      content: this.content,
      fileActions: this.fileActions,
      kernelWasmUrl: this.kernelWasmUrl,
      onAgentRequest: (request) => this.requestAgent(request),
      onChange: (content) => this.changeContent(content),
      preview: this.preview,
      saveStatus: this.saveStatus,
      theme: this.theme,
    });
  }
}

export class A3SPresentationEditorElement extends A3SContentEditorElement<PresentationContent> {
  #collaboration: OfficeCollaborationSession | undefined;

  static get observedAttributes() {
    return ['kernel-wasm-url', 'preview', 'save-status', 'theme'];
  }

  get kernelWasmUrl(): string | undefined {
    return this.getAttribute('kernel-wasm-url') ?? undefined;
  }

  get collaboration(): OfficeCollaborationSession | undefined {
    return this.#collaboration;
  }

  set collaboration(value: OfficeCollaborationSession | undefined) {
    this.#collaboration = value;
    this.requestRender();
  }

  set kernelWasmUrl(value: string | undefined) {
    if (value === undefined) this.removeAttribute('kernel-wasm-url');
    else this.setAttribute('kernel-wasm-url', value);
  }

  protected editorNode(): ReactNode {
    if (!this.content) return missingContent('presentation', this.theme);
    return createElement(PresentationEditor, {
      collaboration: this.collaboration,
      content: this.content,
      fileActions: this.fileActions,
      kernelWasmUrl: this.kernelWasmUrl,
      onAgentRequest: (request) => this.requestAgent(request),
      onChange: (content) => this.changeContent(content),
      onStartSlideshow: () =>
        dispatchDetail(this, 'start-slideshow', this.content),
      preview: this.preview,
      saveStatus: this.saveStatus,
      theme: this.theme,
    });
  }
}

export class A3SPdfViewerElement extends A3SOfficeElement {
  #loadSource: (() => Promise<Blob>) | undefined;
  #onSave: ((pdf: Blob) => Promise<boolean>) | undefined;

  static get observedAttributes() {
    return ['file-name', 'save-label', 'source-key', 'theme', 'wasm-url'];
  }

  get loadSource(): (() => Promise<Blob>) | undefined {
    return this.#loadSource;
  }

  set loadSource(value: (() => Promise<Blob>) | undefined) {
    this.#loadSource = value;
    this.requestRender();
  }

  get onSave(): ((pdf: Blob) => Promise<boolean>) | undefined {
    return this.#onSave;
  }

  set onSave(value: ((pdf: Blob) => Promise<boolean>) | undefined) {
    this.#onSave = value;
    this.requestRender();
  }

  get theme(): OfficeTheme {
    return themeFrom(this);
  }

  set theme(value: OfficeTheme) {
    this.setAttribute('theme', value);
  }

  protected editorNode(): ReactNode {
    if (!this.loadSource)
      return missingContent('PDF source loader', this.theme);
    return createElement(PdfViewer, {
      fileName: this.getAttribute('file-name') ?? undefined,
      loadSource: this.loadSource,
      onSave: this.onSave,
      saveLabel: this.getAttribute('save-label') ?? undefined,
      sourceKey: this.getAttribute('source-key') ?? undefined,
      theme: this.theme,
      wasmUrl: this.getAttribute('wasm-url') ?? undefined,
    });
  }
}

export const A3S_OFFICE_ELEMENT_NAMES = {
  document: 'a3s-document-editor',
  markdown: 'a3s-markdown-editor',
  pdf: 'a3s-pdf-viewer',
  presentation: 'a3s-presentation-editor',
  spreadsheet: 'a3s-spreadsheet-editor',
} as const;

export function defineA3SOfficeElements(
  registry?: CustomElementRegistry,
): void {
  const target =
    registry ??
    (typeof customElements === 'undefined' ? undefined : customElements);
  if (!target) return;

  const elements = [
    [A3S_OFFICE_ELEMENT_NAMES.document, A3SDocumentEditorElement],
    [A3S_OFFICE_ELEMENT_NAMES.markdown, A3SMarkdownEditorElement],
    [A3S_OFFICE_ELEMENT_NAMES.spreadsheet, A3SSpreadsheetEditorElement],
    [A3S_OFFICE_ELEMENT_NAMES.presentation, A3SPresentationEditorElement],
    [A3S_OFFICE_ELEMENT_NAMES.pdf, A3SPdfViewerElement],
  ] as const;

  for (const [name, elementClass] of elements) {
    if (!target.get(name)) target.define(name, elementClass);
  }
}

declare global {
  interface HTMLElementTagNameMap {
    'a3s-document-editor': A3SDocumentEditorElement;
    'a3s-markdown-editor': A3SMarkdownEditorElement;
    'a3s-pdf-viewer': A3SPdfViewerElement;
    'a3s-presentation-editor': A3SPresentationEditorElement;
    'a3s-spreadsheet-editor': A3SSpreadsheetEditorElement;
  }
}
