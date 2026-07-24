import { expect, test } from '@rstest/core';
import {
  A3S_OFFICE_ELEMENT_NAMES,
  A3SDocumentEditorElement,
  A3SMarkdownEditorElement,
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

  element.remove();
});
