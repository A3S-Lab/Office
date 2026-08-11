import { Extension } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import { waitFor } from '@testing-library/dom';
import type { DocumentContent, DocumentReviewConflictEvent } from '../src/core';
import {
  A3S_OFFICE_ELEMENT_NAMES,
  A3SDocumentEditorElement,
  A3SMarkdownEditorElement,
  type A3SSpreadsheetEditorElement,
  defineA3SOfficeElements,
} from '../src/web-component';

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
