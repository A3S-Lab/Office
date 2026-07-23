import { expect, test } from '@rstest/core';
import {
  A3S_OFFICE_ELEMENT_NAMES,
  A3SDocumentEditorElement,
  defineA3SOfficeElements,
} from '../src/web-component';

test('registers every custom element idempotently', async () => {
  defineA3SOfficeElements();
  defineA3SOfficeElements();

  expect(customElements.get(A3S_OFFICE_ELEMENT_NAMES.document)).toBe(
    A3SDocumentEditorElement,
  );

  const element = document.createElement(A3S_OFFICE_ELEMENT_NAMES.document);
  document.body.append(element);
  await new Promise((resolve) => setTimeout(resolve, 0));

  expect(element.textContent).toContain("Set the element's content property");

  element.remove();
});
