import { expect, test } from '@rstest/core';
import { fireEvent, waitFor } from '@testing-library/dom';
import { Extension } from '@tiptap/core';
import {
  createOfficeCollaborationSession,
  type DocumentContent,
  type DocumentReviewConflictEvent,
  initializeOfficeDocumentCollaboration,
  initializeOfficeMarkdownCollaboration,
  initializeOfficePresentationCollaboration,
  initializeOfficeSpreadsheetCollaboration,
  readOfficeMarkdownCollaboration,
} from '../src/core';
import {
  A3S_OFFICE_ELEMENT_NAMES,
  A3SDocumentEditorElement,
  A3SMarkdownEditorElement,
  type A3SPresentationEditorElement,
  type A3SSpreadsheetEditorElement,
  defineA3SOfficeElements,
} from '../src/web-component';
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

  element.remove();
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
