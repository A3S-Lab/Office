import { mergeAttributes, Node } from '@tiptap/core';

export const DOCUMENT_EQUATION_OPAQUE_SELECTOR =
  'span[data-document-equation-opaque]';
export const MAX_DOCUMENT_EQUATION_OPAQUE_OMML_LENGTH = 262_144;

export interface WorkDocumentEquationOpaque {
  display: 'inline' | 'block';
  omml: string;
  text: string;
}

export const DocumentEquationOpaque = Node.create({
  name: 'documentEquationOpaque',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: true,

  addAttributes() {
    return {
      display: {
        default: 'inline' as const,
        parseHTML: (element) =>
          element instanceof HTMLElement &&
          element.dataset.equationDisplay === 'block'
            ? 'block'
            : 'inline',
        renderHTML: (attributes) => ({
          'data-equation-display':
            attributes.display === 'block' ? 'block' : 'inline',
        }),
      },
      omml: {
        default: '',
        parseHTML: (element) =>
          element instanceof HTMLElement
            ? decodeDocumentEquationOpaqueOmml(
                element.dataset.equationOmml ?? '',
              ) ?? ''
            : '',
        renderHTML: (attributes) => {
          const encoded = encodeDocumentEquationOpaqueOmml(
            typeof attributes.omml === 'string' ? attributes.omml : '',
          );
          return encoded ? { 'data-equation-omml': encoded } : {};
        },
      },
      text: {
        default: 'Equation',
        parseHTML: (element) =>
          element instanceof HTMLElement
            ? element.textContent?.trim() ||
              element.getAttribute('aria-label')?.trim() ||
              'Equation'
            : 'Equation',
        renderHTML: () => ({}),
      },
    };
  },

  parseHTML() {
    return [
      {
        tag: DOCUMENT_EQUATION_OPAQUE_SELECTOR,
        priority: 130,
        getAttrs: (node) => {
          if (!(node instanceof HTMLElement)) return false;
          const opaque = documentEquationOpaqueFromElement(node);
          return opaque
            ? {
                display: opaque.display,
                omml: opaque.omml,
                text: opaque.text,
              }
            : false;
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const display = node.attrs.display === 'block' ? 'block' : 'inline';
    const text =
      typeof node.attrs.text === 'string' && node.attrs.text.trim()
        ? node.attrs.text.trim()
        : 'Equation';
    const encoded = encodeDocumentEquationOpaqueOmml(
      typeof node.attrs.omml === 'string' ? node.attrs.omml : '',
    );
    return [
      'span',
      mergeAttributes(HTMLAttributes, {
        'data-document-equation-opaque': 'true',
        'data-equation-display': display,
        ...(encoded ? { 'data-equation-omml': encoded } : {}),
        'aria-label': text,
        class: documentEquationOpaqueClassName(display),
        contenteditable: 'false',
        role: 'math',
      }),
      text,
    ];
  },

  renderText({ node }) {
    return typeof node.attrs.text === 'string' && node.attrs.text.trim()
      ? node.attrs.text.trim()
      : 'Equation';
  },
});

export function createDocumentEquationOpaqueElement(
  document: Document,
  source: WorkDocumentEquationOpaque,
): HTMLElement {
  const opaque = normalizeDocumentEquationOpaque(source);
  if (!opaque) throw new Error('The opaque equation markup is invalid.');
  const element = document.createElement('span');
  const encoded = encodeDocumentEquationOpaqueOmml(opaque.omml);
  element.dataset.documentEquationOpaque = 'true';
  element.dataset.equationDisplay = opaque.display;
  element.dataset.equationOmml = encoded;
  element.className = documentEquationOpaqueClassName(opaque.display);
  element.contentEditable = 'false';
  element.setAttribute('role', 'math');
  element.setAttribute('aria-label', opaque.text);
  element.textContent = opaque.text;
  return element;
}

export function documentEquationOpaqueFromElement(
  element: Element,
): WorkDocumentEquationOpaque | null {
  if (!(element instanceof HTMLElement)) return null;
  if (element.dataset.documentEquationOpaque !== 'true') return null;
  const omml = decodeDocumentEquationOpaqueOmml(
    element.dataset.equationOmml ?? '',
  );
  if (!omml) return null;
  return normalizeDocumentEquationOpaque({
    display:
      element.dataset.equationDisplay === 'block' ? 'block' : 'inline',
    omml,
    text:
      element.textContent?.trim() ||
      element.getAttribute('aria-label')?.trim() ||
      'Equation',
  });
}

export function normalizeDocumentEquationOpaque(
  source: unknown,
): WorkDocumentEquationOpaque | null {
  if (!source || typeof source !== 'object') return null;
  const record = source as Record<string, unknown>;
  const display =
    record.display === 'block' || record.display === 'inline'
      ? record.display
      : null;
  const omml = typeof record.omml === 'string' ? record.omml.trim() : '';
  const text =
    typeof record.text === 'string' && record.text.trim()
      ? record.text.trim()
      : 'Equation';
  if (!display || !omml || omml.length > MAX_DOCUMENT_EQUATION_OPAQUE_OMML_LENGTH) {
    return null;
  }
  return { display, omml, text };
}

export function encodeDocumentEquationOpaqueOmml(omml: string): string {
  if (!omml || omml.length > MAX_DOCUMENT_EQUATION_OPAQUE_OMML_LENGTH) return '';
  const bytes = new TextEncoder().encode(omml);
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}

export function decodeDocumentEquationOpaqueOmml(
  encoded: string,
): string | null {
  if (!encoded) return null;
  try {
    const binary = atob(encoded);
    const bytes = Uint8Array.from(binary, (character) =>
      character.charCodeAt(0),
    );
    const omml = new TextDecoder().decode(bytes).trim();
    if (!omml || omml.length > MAX_DOCUMENT_EQUATION_OPAQUE_OMML_LENGTH) {
      return null;
    }
    return omml;
  } catch {
    return null;
  }
}

function documentEquationOpaqueClassName(
  display: WorkDocumentEquationOpaque['display'],
): string {
  return display === 'block'
    ? 'work-document-equation opaque block'
    : 'work-document-equation opaque inline';
}
