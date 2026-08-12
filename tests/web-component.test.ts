import { expect, test } from '@rstest/core';
import { fireEvent, waitFor } from '@testing-library/dom';
import { Extension } from '@tiptap/core';
import type { ReactElement } from 'react';
import {
  createOfficeCollaborationSession,
  createPdfCollaborationContent,
  type DocumentContent,
  type DocumentReviewConflictEvent,
  initializeOfficeDocumentCollaboration,
  initializeOfficeMarkdownCollaboration,
  initializeOfficePresentationCollaboration,
  initializeOfficeSpreadsheetCollaboration,
  type OfficeCollaborationPresence,
  readOfficeMarkdownCollaboration,
} from '../src/core';
import type { PdfEvidenceRegion } from '../src/react';
import {
  A3S_OFFICE_ELEMENT_NAMES,
  A3SDocumentEditorElement,
  A3SMarkdownEditorElement,
  A3SPdfViewerElement,
  type A3SPresentationEditorElement,
  type A3SSpreadsheetEditorElement,
  defineA3SOfficeElements,
} from '../src/web-component';
import { PDF_COLLABORATION_SOURCE } from './fixtures/pdf-collaboration';
import { presentationCollaborationFixture } from './fixtures/presentation-collaboration';
import { spreadsheetCollaborationFixture } from './fixtures/spreadsheet-collaboration';

test('registers every custom element idempotently', async () => {
  defineA3SOfficeElements();
  defineA3SOfficeElements();

  expect(customElements.get(A3S_OFFICE_ELEMENT_NAMES.document)).toBe(
    A3SDocumentEditorElement,
  );
  expect(customElements.get(A3S_OFFICE_ELEMENT_NAMES.markdown)).toBe(
    A3SMarkdownEditorElement,
  );

  const element = document.createElement(A3S_OFFICE_ELEMENT_NAMES.document);
  document.body.append(element);
  await new Promise((resolve) => setTimeout(resolve, 0));

  expect(element.textContent).toContain("Set the element's content property");

  const documentEditor = element as A3SDocumentEditorElement;
  documentEditor.artifactId = 'document-1';
  expect(documentEditor.getAttribute('artifact-id')).toBe('document-1');
  documentEditor.kernelWasmUrl = '/assets/office-kernel.wasm';
  expect(documentEditor.getAttribute('kernel-wasm-url')).toBe(
    '/assets/office-kernel.wasm',
  );
  documentEditor.layoutFonts = [
    {
      family: 'Host Font',
      id: 'host-font',
      url: '/assets/host-font.woff2',
    },
  ];
  expect(documentEditor.layoutFonts?.[0]?.id).toBe('host-font');
  const hostExtension = Extension.create({ name: 'hostExtension' });
  documentEditor.extensions = [hostExtension];
  expect(documentEditor.extensions).toEqual([hostExtension]);
  const getSelectionMenuItems = () => [];
  documentEditor.getSelectionMenuItems = getSelectionMenuItems;
  expect(documentEditor.getSelectionMenuItems).toBe(getSelectionMenuItems);
  const presence = {} as OfficeCollaborationPresence;
  documentEditor.presence = presence;
  expect(documentEditor.presence).toBe(presence);

  const markdownEditor = document.createElement(
    A3S_OFFICE_ELEMENT_NAMES.markdown,
  ) as A3SMarkdownEditorElement;
  markdownEditor.extensions = [hostExtension];
  expect(markdownEditor.extensions).toEqual([hostExtension]);
  markdownEditor.getSelectionMenuItems = getSelectionMenuItems;
  expect(markdownEditor.getSelectionMenuItems).toBe(getSelectionMenuItems);

  const spreadsheetEditor = document.createElement(
    A3S_OFFICE_ELEMENT_NAMES.spreadsheet,
  ) as A3SSpreadsheetEditorElement;
  spreadsheetEditor.kernelWasmUrl = '/assets/spreadsheet-kernel.wasm';
  expect(spreadsheetEditor.getAttribute('kernel-wasm-url')).toBe(
    '/assets/spreadsheet-kernel.wasm',
  );
  const presentationEditor = document.createElement(
    A3S_OFFICE_ELEMENT_NAMES.presentation,
  );
  presentationEditor.kernelWasmUrl = '/assets/presentation-kernel.wasm';
  expect(presentationEditor.getAttribute('kernel-wasm-url')).toBe(
    '/assets/presentation-kernel.wasm',
  );

  const pdfViewer = document.createElement(
    A3S_OFFICE_ELEMENT_NAMES.pdf,
  ) as A3SPdfViewerElement;
  const pdfSession = createOfficeCollaborationSession({
    artifactId: 'element-shared-pdf',
    kind: 'pdf',
  });
  const loadPdfSource = () => Promise.resolve(new Blob());
  const onCollaborationChange = () => undefined;
  pdfViewer.collaboration = pdfSession;
  pdfViewer.loadSource = loadPdfSource;
  pdfViewer.onCollaborationChange = onCollaborationChange;
  expect(pdfViewer.collaboration).toBe(pdfSession);
  expect(pdfViewer.loadSource).toBe(loadPdfSource);
  expect(pdfViewer.onCollaborationChange).toBe(onCollaborationChange);

  const evidenceOverlay = {
    coordinateBasis: 1_000_000 as const,
    loadPage: async () => null,
    renderProfileSha256: 'b'.repeat(64),
    sourceSha256: 'a'.repeat(64),
  };
  pdfViewer.evidenceOverlay = evidenceOverlay;
  pdfViewer.selectedEvidenceRegionId = 'region-1';
  pdfViewer.worker = false;
  expect(pdfViewer.evidenceOverlay).toBe(evidenceOverlay);
  expect(pdfViewer.selectedEvidenceRegionId).toBe('region-1');
  expect(pdfViewer.getAttribute('worker')).toBe('false');
  expect(pdfViewer.worker).toBe(false);

  const parent = document.createElement('div');
  parent.append(pdfViewer);
  pdfViewer.loadSource = async () =>
    new Blob(['%PDF-1.7'], { type: 'application/pdf' });
  let selectedRegion: PdfEvidenceRegion | undefined;
  let changedPage: number | undefined;
  parent.addEventListener('evidence-select', (event) => {
    selectedRegion = (event as CustomEvent<PdfEvidenceRegion>).detail;
  });
  parent.addEventListener('page-change', (event) => {
    changedPage = (event as CustomEvent<number>).detail;
  });
  const viewerNode = (
    pdfViewer as unknown as {
      editorNode(): ReactElement<{
        onEvidenceRegionSelect?: (region: PdfEvidenceRegion) => void;
        onPageChange?: (pageNumber: number) => void;
      }>;
    }
  ).editorNode();
  const region: PdfEvidenceRegion = {
    bounds: { bottom: 20, left: 10, right: 30, top: 5 },
    id: 'region-1',
    sourceRegionIds: ['source-region-1'],
    targetIds: ['node-1'],
  };
  viewerNode.props.onEvidenceRegionSelect?.(region);
  viewerNode.props.onPageChange?.(3);
  expect(selectedRegion).toBe(region);
  expect(changedPage).toBe(3);

  element.remove();
  parent.remove();
});

test('bridges PDF collaboration snapshots to its callback and custom event', () => {
  const name = 'a3s-test-pdf-collaboration-events';
  if (!customElements.get(name)) {
    customElements.define(name, InspectablePdfViewerElement);
  }
  const element = document.createElement(name) as InspectablePdfViewerElement;
  const detail = createPdfCollaborationContent(PDF_COLLABORATION_SOURCE);
  const presence = {} as OfficeCollaborationPresence;
  const callbacks: unknown[] = [];
  const events: CustomEvent[] = [];
  element.loadSource = () => Promise.resolve(new Blob());
  element.presence = presence;
  element.onCollaborationChange = (content) => callbacks.push(content);
  element.addEventListener('collaboration-change', (event) => {
    events.push(event as CustomEvent);
  });

  const node = element.readEditorNode();
  if (!node || typeof node !== 'object' || !('props' in node)) {
    throw new Error('Expected the PDF custom element to render a React node.');
  }
  const props = (
    node as {
      props: {
        onCollaborationChange: (content: typeof detail) => void;
        presence?: OfficeCollaborationPresence;
      };
    }
  ).props;
  expect(props.presence).toBe(presence);
  props.onCollaborationChange(detail);

  expect(callbacks).toEqual([detail]);
  expect(events).toHaveLength(1);
  expect(events[0]?.detail).toBe(detail);
  expect(events[0]?.bubbles).toBe(true);
  expect(events[0]?.composed).toBe(true);
});

test('dispatches controlled document review conflicts from the custom element', async () => {
  defineA3SOfficeElements();
  const element = document.createElement(
    A3S_OFFICE_ELEMENT_NAMES.document,
  ) as A3SDocumentEditorElement;
  element.artifactId = 'document-1';
  element.content = reviewDocument('Alpha');
  const events: DocumentReviewConflictEvent[] = [];
  element.addEventListener('review-conflict', (event) => {
    events.push((event as CustomEvent<DocumentReviewConflictEvent>).detail);
  });
  document.body.append(element);
  await waitFor(() => {
    expect(element.querySelector('[role="textbox"]')).not.toBeNull();
  });

  element.content = reviewDocument('Omega');

  await waitFor(() => expect(events).toHaveLength(1));
  expect(events[0]?.conflicts[0]).toMatchObject({
    id: 'comment-1',
    kind: 'comment',
    reason: 'text-changed',
  });

  element.remove();
});

test('passes a synchronized Document session through the custom element property', async () => {
  defineA3SOfficeElements();
  const content: DocumentContent = {
    type: 'document',
    html: '<p>Shared through an element</p>',
    pageSize: 'a4',
  };
  const session = createOfficeCollaborationSession({
    artifactId: 'element-shared-document',
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(session, content);
  const element = document.createElement(
    A3S_OFFICE_ELEMENT_NAMES.document,
  ) as A3SDocumentEditorElement;
  element.collaboration = session;
  element.content = content;
  document.body.append(element);

  await waitFor(() => {
    expect(element.querySelector('[aria-label="文档正文"]')).not.toBeNull();
  });
  expect(element.collaboration).toBe(session);

  element.remove();
});

test('passes a synchronized Markdown session through the custom element property', async () => {
  defineA3SOfficeElements();
  const session = createOfficeCollaborationSession({
    artifactId: 'element-shared-markdown',
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(session, {
    type: 'markdown',
    markdown: '# Shared through an element',
  });
  const element = document.createElement(
    A3S_OFFICE_ELEMENT_NAMES.markdown,
  ) as A3SMarkdownEditorElement;
  element.collaboration = session;
  element.content = { type: 'markdown', markdown: '# Stale host value' };
  document.body.append(element);

  const source = await waitFor(() => {
    const textarea = element.querySelector<HTMLTextAreaElement>('textarea');
    expect(textarea).not.toBeNull();
    if (!textarea) throw new Error('Expected the Markdown source editor.');
    return textarea;
  });
  expect(element.collaboration).toBe(session);
  expect(source.value).toBe('# Shared through an element');
  fireEvent.change(source, {
    target: { value: '# Element collaboration edit' },
  });

  await waitFor(() =>
    expect(readOfficeMarkdownCollaboration(session).markdown).toBe(
      '# Element collaboration edit',
    ),
  );
  element.remove();
});

test('passes a synchronized Presentation session through the custom element property', async () => {
  defineA3SOfficeElements();
  const content = presentationCollaborationFixture();
  const session = createOfficeCollaborationSession({
    artifactId: 'element-shared-presentation',
    kind: 'presentation',
  });
  initializeOfficePresentationCollaboration(session, content);
  const element = document.createElement(
    A3S_OFFICE_ELEMENT_NAMES.presentation,
  ) as A3SPresentationEditorElement;
  element.collaboration = session;
  element.content = content;
  document.body.append(element);

  await waitFor(() => {
    expect(element.textContent).toContain('Shared presentation');
  });
  expect(element.collaboration).toBe(session);

  element.remove();
});

test('passes a synchronized Spreadsheet session through the custom element property', async () => {
  defineA3SOfficeElements();
  const content = spreadsheetCollaborationFixture();
  const session = createOfficeCollaborationSession({
    artifactId: 'element-shared-spreadsheet',
    kind: 'spreadsheet',
  });
  initializeOfficeSpreadsheetCollaboration(session, content);
  const element = document.createElement(
    A3S_OFFICE_ELEMENT_NAMES.spreadsheet,
  ) as A3SSpreadsheetEditorElement;
  element.collaboration = session;
  element.content = content;
  element.preview = true;
  document.body.append(element);

  await waitFor(() => {
    expect(element.textContent).toContain('Inputs');
  });
  expect(element.collaboration).toBe(session);

  element.remove();
});

function reviewDocument(anchorText: string): DocumentContent {
  return {
    type: 'document',
    html: `<p><span data-comment-id="comment-1" data-document-comment="true">${anchorText}</span></p>`,
    pageSize: 'a4',
    comments: [
      {
        id: 'comment-1',
        author: 'Reviewer',
        date: '',
        text: 'Review Alpha.',
        resolved: false,
      },
    ],
  };
}

class InspectablePdfViewerElement extends A3SPdfViewerElement {
  readEditorNode() {
    return this.editorNode();
  }
}
