import {
  type CommandProps,
  type Editor,
  mergeAttributes,
  Node,
} from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { TextSelection } from '@tiptap/pm/state';
import { activeDocumentSectionFromState } from './work-document-section-editor';
import {
  documentCaptionDisplay,
  documentCaptionKind,
  documentCaptionLabel,
  type WorkDocumentCaptionKind,
  type WorkDocumentCaptionTarget,
} from './work-document-captions';
import { createWorkId } from './work-templates';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentCaption: {
      insertDocumentCaption: (
        kind: WorkDocumentCaptionKind,
        title: string,
      ) => ReturnType;
      insertDocumentCrossReference: (
        target: WorkDocumentCaptionTarget,
      ) => ReturnType;
    };
  }
}

export const DocumentCaption = Node.create({
  name: 'documentCaption',
  group: 'block',
  content: 'inline*',
  defining: true,

  addCommands() {
    return {
      insertDocumentCaption: (kind, title) => (props) =>
        insertDocumentCaptionCommand(props, kind, title),
      insertDocumentCrossReference: (target) => (props) =>
        insertDocumentCrossReferenceCommand(props, target),
    };
  },

  addAttributes() {
    return {
      id: hiddenAttribute(''),
      kind: hiddenAttribute('figure'),
      number: hiddenAttribute(1),
    };
  },

  parseHTML() {
    return [
      {
        tag: 'figcaption[data-document-caption]',
        getAttrs: (node) => {
          if (!(node instanceof HTMLElement)) return false;
          return {
            id: node.dataset.captionId ?? '',
            kind: documentCaptionKind(node.dataset.captionKind) ?? 'figure',
            number: positiveInteger(node.dataset.captionNumber),
          };
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const kind = documentCaptionKind(node.attrs.kind) ?? 'figure';
    const number = positiveInteger(node.attrs.number);
    return [
      'figcaption',
      mergeAttributes(HTMLAttributes, {
        'data-document-caption': 'true',
        'data-caption-id':
          typeof node.attrs.id === 'string' ? node.attrs.id : '',
        'data-caption-kind': kind,
        'data-caption-number': String(number),
        'data-caption-label': documentCaptionLabel(kind),
        class: 'work-document-caption',
      }),
      0,
    ];
  },

  renderText({ node }) {
    const kind = documentCaptionKind(node.attrs.kind) ?? 'figure';
    return `${documentCaptionDisplay(kind, positiveInteger(node.attrs.number))} ${node.textContent}`.trim();
  },
});

export const DocumentCrossReference = Node.create({
  name: 'documentCrossReference',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: true,

  addAttributes() {
    return {
      targetId: hiddenAttribute(''),
      kind: hiddenAttribute('figure'),
      number: hiddenAttribute(1),
      orphaned: hiddenAttribute(false),
    };
  },

  parseHTML() {
    return [
      {
        tag: 'span[data-document-cross-reference]',
        getAttrs: (node) => {
          if (!(node instanceof HTMLElement)) return false;
          return {
            targetId: node.dataset.referenceTargetId ?? '',
            kind: documentCaptionKind(node.dataset.captionKind) ?? 'figure',
            number: positiveInteger(node.dataset.captionNumber),
            orphaned: node.dataset.referenceOrphaned === 'true',
          };
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const kind = documentCaptionKind(node.attrs.kind) ?? 'figure';
    const number = positiveInteger(node.attrs.number);
    const orphaned = Boolean(node.attrs.orphaned);
    return [
      'span',
      mergeAttributes(HTMLAttributes, {
        'data-document-cross-reference': 'true',
        'data-reference-target-id':
          typeof node.attrs.targetId === 'string' ? node.attrs.targetId : '',
        'data-caption-kind': kind,
        'data-caption-number': String(number),
        'data-caption-label': documentCaptionLabel(kind),
        'data-reference-orphaned': orphaned ? 'true' : undefined,
        class: 'work-document-cross-reference',
      }),
      orphaned ? '引用缺失' : documentCaptionDisplay(kind, number),
    ];
  },

  renderText({ node }) {
    if (node.attrs.orphaned) return '引用缺失';
    const kind = documentCaptionKind(node.attrs.kind) ?? 'figure';
    return documentCaptionDisplay(kind, positiveInteger(node.attrs.number));
  },
});

export function editorDocumentCaptionTargets(
  editor: Editor,
): WorkDocumentCaptionTarget[] {
  return documentCaptionTargets(editor.state.doc);
}

function insertDocumentCaptionCommand(
  { dispatch, editor, state, tr }: CommandProps,
  kind: WorkDocumentCaptionKind,
  title: string,
): boolean {
  const section = activeDocumentSectionFromState(state);
  const captionType = editor.schema.nodes.documentCaption;
  const paragraphType = editor.schema.nodes.paragraph;
  if (
    !section ||
    !captionType ||
    !paragraphType ||
    selectionInsideNode(state, 'documentCaption')
  ) {
    return false;
  }
  const child = activeSectionChild(section, state.selection.from);
  if (!child) return false;
  if (!dispatch) return true;

  const id = createWorkId(`caption-${kind}`);
  const number = nextCaptionNumber(state.doc, kind);
  const normalizedTitle = title.trim();
  const caption = captionType.create(
    { id, kind, number },
    normalizedTitle ? editor.schema.text(normalizedTitle) : undefined,
  );
  const insertPosition = section.position + 1 + child.offset + child.nodeSize;
  tr.insert(insertPosition, caption);
  let selectionPosition = insertPosition + caption.nodeSize;
  if (child.index === section.node.childCount - 1) {
    tr.insert(selectionPosition, paragraphType.create());
    selectionPosition += 1;
  }
  tr.setSelection(TextSelection.near(tr.doc.resolve(selectionPosition)));
  tr.scrollIntoView();
  return true;
}

function insertDocumentCrossReferenceCommand(
  { dispatch, editor, tr }: CommandProps,
  target: WorkDocumentCaptionTarget,
): boolean {
  const referenceType = editor.schema.nodes.documentCrossReference;
  if (!referenceType || !target.id.trim()) return false;
  if (!dispatch) return true;
  tr.replaceSelectionWith(
    referenceType.create({
      targetId: target.id,
      kind: target.kind,
      number: target.number,
      orphaned: false,
    }),
    false,
  );
  tr.scrollIntoView();
  return true;
}

function documentCaptionTargets(
  document: ProseMirrorNode,
): WorkDocumentCaptionTarget[] {
  const targets: WorkDocumentCaptionTarget[] = [];
  document.descendants((node) => {
    if (node.type.name !== 'documentCaption') return;
    const id = typeof node.attrs.id === 'string' ? node.attrs.id.trim() : '';
    const kind = documentCaptionKind(node.attrs.kind);
    if (!id || !kind) return;
    const number = positiveInteger(node.attrs.number);
    targets.push({
      id,
      kind,
      number,
      label: kind === 'figure' ? '图' : '表',
      title: node.textContent.trim(),
      display: documentCaptionDisplay(kind, number),
    });
  });
  return targets;
}

function nextCaptionNumber(
  document: ProseMirrorNode,
  kind: WorkDocumentCaptionKind,
): number {
  return (
    documentCaptionTargets(document).filter((caption) => caption.kind === kind)
      .length + 1
  );
}

function activeSectionChild(
  section: NonNullable<ReturnType<typeof activeDocumentSectionFromState>>,
  selectionPosition: number,
): { index: number; offset: number; nodeSize: number } | null {
  const relativePosition = Math.max(
    0,
    selectionPosition - section.position - 1,
  );
  let active: { index: number; offset: number; nodeSize: number } | null = null;
  section.node.forEach((node, offset, index) => {
    if (relativePosition >= offset) {
      active = { index, offset, nodeSize: node.nodeSize };
    }
  });
  return active;
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
