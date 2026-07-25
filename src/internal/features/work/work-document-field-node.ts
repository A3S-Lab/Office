import {
  type CommandProps,
  type Editor,
  mergeAttributes,
  Node,
} from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { documentPageDescriptors } from './work-document-pages';
import { syncDocumentContentFromHtml } from './work-document-section';
import {
  documentFieldDisplay,
  documentFieldInstruction,
  documentFieldKind,
  documentFieldLabel,
  docxDocumentFieldKind,
  type WorkDocumentFieldContext,
  type WorkDocumentFieldKind,
} from './work-document-fields';
import { createWorkId } from './work-templates';
import type { WorkDocumentContent } from './work-types';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentField: {
      insertDocumentField: (kind: WorkDocumentFieldKind) => ReturnType;
      refreshDocumentFields: (content: WorkDocumentContent) => ReturnType;
    };
  }
}

export const DocumentField = Node.create({
  name: 'documentField',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: true,

  addCommands() {
    return {
      insertDocumentField: (kind) => (props) =>
        insertDocumentFieldCommand(props, kind),
      refreshDocumentFields: (content) => (props) =>
        refreshDocumentFieldsCommand(props, content),
    };
  },

  addAttributes() {
    return {
      id: hiddenAttribute(''),
      kind: hiddenAttribute('page'),
      instruction: hiddenAttribute('PAGE'),
      display: hiddenAttribute('1'),
    };
  },

  parseHTML() {
    return [
      {
        tag: 'span[data-document-field]',
        getAttrs: (node) => {
          if (!(node instanceof HTMLElement)) return false;
          const instruction = node.dataset.fieldInstruction?.trim() ?? '';
          const kind =
            documentFieldKind(node.dataset.fieldKind) ??
            docxDocumentFieldKind(instruction) ??
            'page';
          return {
            id: node.dataset.fieldId ?? '',
            kind,
            instruction: instruction || documentFieldInstruction(kind),
            display:
              node.dataset.fieldDisplay?.trim() ||
              node.textContent?.trim() ||
              documentFieldLabel(kind),
          };
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const instruction =
      typeof node.attrs.instruction === 'string'
        ? node.attrs.instruction.trim()
        : '';
    const kind =
      documentFieldKind(node.attrs.kind) ??
      docxDocumentFieldKind(instruction) ??
      'page';
    const display =
      typeof node.attrs.display === 'string' && node.attrs.display.trim()
        ? node.attrs.display.trim()
        : documentFieldLabel(kind);
    return [
      'span',
      mergeAttributes(HTMLAttributes, {
        'data-document-field': 'true',
        'data-field-id': typeof node.attrs.id === 'string' ? node.attrs.id : '',
        'data-field-kind': kind,
        'data-field-instruction': instruction || documentFieldInstruction(kind),
        'data-field-display': display,
        class: 'work-document-field',
        title: documentFieldLabel(kind),
      }),
      display,
    ];
  },

  renderText({ node }) {
    const kind = documentFieldKind(node.attrs.kind) ?? 'page';
    return typeof node.attrs.display === 'string' && node.attrs.display.trim()
      ? node.attrs.display.trim()
      : documentFieldLabel(kind);
  },
});

function insertDocumentFieldCommand(
  { dispatch, editor, state, tr }: CommandProps,
  kind: WorkDocumentFieldKind,
): boolean {
  const fieldType = editor.schema.nodes.documentField;
  if (!fieldType) return false;
  if (!dispatch) return true;
  const instruction = documentFieldInstruction(kind);
  tr.replaceSelectionWith(
    fieldType.create({
      id: createWorkId('field'),
      kind,
      instruction,
      display: documentFieldDisplay(kind, fallbackContext(state), instruction),
    }),
    false,
  );
  tr.scrollIntoView();
  return true;
}

function refreshDocumentFieldsCommand(
  { editor, state, tr }: CommandProps,
  content: WorkDocumentContent,
): boolean {
  const canonical = syncDocumentContentFromHtml(content, editor.getHTML());
  const displays = documentPageDescriptors(canonical).flatMap((page) =>
    page.segments.flatMap((segment) => {
      const document = new DOMParser().parseFromString(
        segment.html,
        'text/html',
      );
      return Array.from(
        document.body.querySelectorAll<HTMLElement>('[data-document-field]'),
      ).map(
        (element) =>
          element.dataset.fieldDisplay?.trim() ||
          element.textContent?.trim() ||
          '',
      );
    }),
  );
  let fieldIndex = 0;
  state.doc.descendants((node, position) => {
    if (node.type.name !== 'documentField') return;
    const display = displays[fieldIndex];
    fieldIndex += 1;
    if (!display || node.attrs.display === display) return;
    tr.setNodeMarkup(position, undefined, { ...node.attrs, display });
  });
  return tr.docChanged;
}

function fallbackContext(state: Editor['state']): WorkDocumentFieldContext {
  let totalPages = 0;
  let sectionNumber = 1;
  let sectionPages = 1;
  let pageNumber = 1;
  let pagesBeforeSelection = 0;
  let foundSelection = false;
  state.doc.forEach((section, offset, index) => {
    if (section.type.name !== 'documentSection') return;
    const pages = countPageBreaks(section) + 1;
    totalPages += pages;
    if (!foundSelection && state.selection.from <= offset + section.nodeSize) {
      sectionNumber = index + 1;
      sectionPages = pages;
      pageNumber =
        pagesBeforeSelection +
        1 +
        countPageBreaksBefore(section, state.selection.from - offset - 1);
      foundSelection = true;
    }
    if (!foundSelection) pagesBeforeSelection += pages;
  });
  return {
    pageNumber,
    totalPages: Math.max(1, totalPages),
    sectionNumber,
    sectionPages,
  };
}

function countPageBreaks(node: ProseMirrorNode): number {
  let count = 0;
  node.descendants((child) => {
    if (child.type.name === 'pageBreak') count += 1;
  });
  return count;
}

function countPageBreaksBefore(
  node: ProseMirrorNode,
  position: number,
): number {
  let count = 0;
  node.descendants((child, offset) => {
    if (offset >= position) return false;
    if (child.type.name === 'pageBreak') count += 1;
    return true;
  });
  return count;
}

function hiddenAttribute(defaultValue: unknown) {
  return {
    default: defaultValue,
    rendered: false,
  };
}
