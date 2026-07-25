import {
  type CommandProps,
  type Editor,
  mergeAttributes,
  Node,
} from '@tiptap/core';
import { TextSelection } from '@tiptap/pm/state';
import { activeDocumentSectionFromState } from './work-document-section-editor';
import {
  documentNoteKey,
  documentNoteKind,
  type WorkDocumentNoteKind,
} from './work-document-notes';
import { createWorkId } from './work-templates';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentNote: {
      insertDocumentNote: (kind: WorkDocumentNoteKind) => ReturnType;
    };
  }
}

export const DocumentNoteReference = Node.create({
  name: 'documentNoteReference',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: true,

  addCommands() {
    return {
      insertDocumentNote: (kind) => (props) =>
        insertDocumentNoteCommand(props, kind),
    };
  },

  addAttributes() {
    return {
      id: hiddenAttribute(''),
      kind: hiddenAttribute('footnote'),
      number: hiddenAttribute(1),
    };
  },

  parseHTML() {
    return [
      {
        tag: 'sup[data-document-note-reference]',
        getAttrs: (node) => {
          if (!(node instanceof HTMLElement)) return false;
          return {
            id: node.dataset.noteId ?? '',
            kind: documentNoteKind(node.dataset.noteKind) ?? 'footnote',
            number: positiveInteger(node.dataset.noteNumber),
          };
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const kind = documentNoteKind(node.attrs.kind) ?? 'footnote';
    const number = positiveInteger(node.attrs.number);
    return [
      'sup',
      mergeAttributes(HTMLAttributes, {
        'data-document-note-reference': 'true',
        'data-note-kind': kind,
        'data-note-id': typeof node.attrs.id === 'string' ? node.attrs.id : '',
        'data-note-number': String(number),
        class: 'work-document-note-reference',
      }),
      String(number),
    ];
  },

  renderText({ node }) {
    return String(positiveInteger(node.attrs.number));
  },
});

export const DocumentNote = Node.create({
  name: 'documentNote',
  group: 'block',
  content: 'block+',
  defining: true,
  isolating: true,

  addAttributes() {
    return {
      id: hiddenAttribute(''),
      kind: hiddenAttribute('footnote'),
      number: hiddenAttribute(1),
    };
  },

  parseHTML() {
    return [
      {
        tag: 'aside[data-document-note]',
        getAttrs: (node) => {
          if (!(node instanceof HTMLElement)) return false;
          return {
            id: node.dataset.noteId ?? '',
            kind: documentNoteKind(node.dataset.noteKind) ?? 'footnote',
            number: positiveInteger(node.dataset.noteNumber),
          };
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const kind = documentNoteKind(node.attrs.kind) ?? 'footnote';
    const number = positiveInteger(node.attrs.number);
    return [
      'aside',
      mergeAttributes(HTMLAttributes, {
        'data-document-note': 'true',
        'data-note-kind': kind,
        'data-note-id': typeof node.attrs.id === 'string' ? node.attrs.id : '',
        'data-note-number': String(number),
        class: 'work-document-note',
      }),
      0,
    ];
  },
});

function insertDocumentNoteCommand(
  { dispatch, editor, state, tr }: CommandProps,
  kind: WorkDocumentNoteKind,
): boolean {
  const section = activeDocumentSectionFromState(state);
  const referenceType = editor.schema.nodes.documentNoteReference;
  const noteType = editor.schema.nodes.documentNote;
  const paragraphType = editor.schema.nodes.paragraph;
  if (
    !section ||
    !referenceType ||
    !noteType ||
    !paragraphType ||
    selectionInsideNode(state, 'documentNote')
  ) {
    return false;
  }
  if (!dispatch) return true;

  const number = nextNoteNumber(state, kind);
  const id = createWorkId(kind);
  const attributes = { id, kind, number };
  tr.replaceSelectionWith(referenceType.create(attributes), false);
  const updatedSection = tr.doc.nodeAt(section.position);
  if (!updatedSection || updatedSection.type.name !== 'documentSection') {
    return false;
  }

  const insertPosition = section.position + updatedSection.nodeSize - 1;
  tr.insert(
    insertPosition,
    noteType.create(attributes, paragraphType.create()),
  );
  tr.setSelection(TextSelection.near(tr.doc.resolve(insertPosition + 2)));
  tr.scrollIntoView();
  return true;
}

function nextNoteNumber(
  state: Editor['state'],
  kind: WorkDocumentNoteKind,
): number {
  const ids = new Set<string>();
  state.doc.descendants((node) => {
    if (
      node.type.name !== 'documentNoteReference' ||
      node.attrs.kind !== kind
    ) {
      return;
    }
    const id = typeof node.attrs.id === 'string' ? node.attrs.id : '';
    if (id) ids.add(documentNoteKey(kind, id));
  });
  return ids.size + 1;
}

function selectionInsideNode(
  state: Editor['state'],
  nodeName: string,
): boolean {
  for (let depth = state.selection.$from.depth; depth > 0; depth -= 1) {
    if (state.selection.$from.node(depth).type.name === nodeName) return true;
  }
  return false;
}

function hiddenAttribute(defaultValue: unknown) {
  return {
    default: defaultValue,
    rendered: false,
  };
}

function positiveInteger(value: unknown): number {
  const number = Number(value);
  return Number.isSafeInteger(number) && number > 0 ? number : 1;
}
