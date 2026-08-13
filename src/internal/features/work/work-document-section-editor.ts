import type { CommandProps, Editor } from '@tiptap/core';
import { Fragment, type Node as ProseMirrorNode } from '@tiptap/pm/model';
import { TextSelection } from '@tiptap/pm/state';
import { createWorkId } from './work-templates';
import {
  documentSectionLayoutFromNodeAttributes,
  documentSectionNodeAttributes,
  type DocumentSectionNodeAttributes,
} from './work-document-section';
import {
  documentPageMarginGlobalsChanged,
  documentPageMarginsForLayout,
  reconcileDocumentPageMarginUpdate,
  synchronizeDocumentPageMarginGlobals,
} from './work-document-page-margins';
import type {
  WorkDocumentSectionBreakType,
  WorkDocumentSectionLayout,
} from './work-types';

export interface ActiveDocumentSection {
  id: string;
  index: number;
  count: number;
  position: number;
  node: ProseMirrorNode;
  layout: WorkDocumentSectionLayout;
}

type DocumentSectionCommandContext = Pick<CommandProps, 'state' | 'tr'>;

export function activeDocumentSection(
  editor: Editor,
): ActiveDocumentSection | null {
  return activeDocumentSectionFromState(editor.state);
}

export function activeDocumentSectionFromState(
  state: Editor['state'],
): ActiveDocumentSection | null {
  const { doc, selection } = state;
  let position: number | null = null;
  let node: ProseMirrorNode | null = null;
  for (let depth = selection.$from.depth; depth > 0; depth -= 1) {
    const candidate = selection.$from.node(depth);
    if (candidate.type.name !== 'documentSection') continue;
    position = selection.$from.before(depth);
    node = candidate;
    break;
  }

  const sections = directDocumentSections(doc);
  if (!node || position === null) {
    const containing = sections.find(
      (section) =>
        selection.from >= section.position &&
        selection.from <= section.position + section.node.nodeSize,
    );
    if (!containing) return null;
    position = containing.position;
    node = containing.node;
  }
  const index = sections.findIndex((section) => section.position === position);
  if (index < 0) return null;
  const id =
    typeof node.attrs.id === 'string' && node.attrs.id
      ? node.attrs.id
      : `document-section-${index + 1}`;
  return {
    id,
    index,
    count: sections.length,
    position,
    node,
    layout: documentSectionLayoutFromNodeAttributes(
      node.attrs as Partial<DocumentSectionNodeAttributes>,
    ),
  };
}

export function updateActiveDocumentSection(
  layout: WorkDocumentSectionLayout,
  commandContext: DocumentSectionCommandContext,
): boolean {
  const state = commandContext.state;
  const section = activeDocumentSectionFromState(state);
  if (!section) return false;
  return updateDocumentSection(section.id, layout, commandContext);
}

export function documentSectionById(
  editor: Editor,
  sectionId: string,
): ActiveDocumentSection | null {
  return documentSectionByIdInDocument(editor.state.doc, sectionId);
}

export function updateDocumentSection(
  sectionId: string,
  layout: WorkDocumentSectionLayout,
  commandContext: DocumentSectionCommandContext,
): boolean {
  const state = commandContext.state;
  const section = documentSectionByIdInDocument(state.doc, sectionId);
  if (!section) return false;
  const next = reconcileDocumentPageMarginUpdate(section.layout, layout);
  if (documentPageMarginGlobalsChanged(section.layout, next)) {
    const globals = documentPageMarginsForLayout(next);
    for (const [index, candidate] of directDocumentSections(
      state.doc,
    ).entries()) {
      const id = documentSectionId(candidate.node, index);
      const current = documentSectionLayoutFromNodeAttributes(
        candidate.node.attrs as Partial<DocumentSectionNodeAttributes>,
      );
      const synchronized = synchronizeDocumentPageMarginGlobals(
        id === sectionId ? next : current,
        globals,
      );
      commandContext.tr.setNodeMarkup(
        candidate.position,
        undefined,
        documentSectionNodeAttributes(synchronized, id),
      );
    }
  } else {
    commandContext.tr.setNodeMarkup(
      section.position,
      undefined,
      documentSectionNodeAttributes(next, section.id),
    );
  }
  return true;
}

export function insertDocumentSection(
  editor: Editor,
  breakAfter: WorkDocumentSectionBreakType,
  commandContext: DocumentSectionCommandContext,
): boolean {
  const state = commandContext.state;
  const section = activeDocumentSectionFromState(state);
  const sectionType = editor.schema.nodes.documentSection;
  const paragraphType = editor.schema.nodes.paragraph;
  if (!section || !sectionType || !paragraphType) return false;

  const children: ProseMirrorNode[] = [];
  section.node.forEach((child) => {
    children.push(child);
  });
  const activeChildIndex = sectionChildIndex(section, state.selection.from);
  const currentChildren = children.slice(0, activeChildIndex + 1);
  const followingChildren = children.slice(activeChildIndex + 1);
  if (!currentChildren.length) currentChildren.push(paragraphType.create());
  if (!followingChildren.length) followingChildren.push(paragraphType.create());

  const currentLayout = { ...section.layout, breakAfter };
  const nextLayout: WorkDocumentSectionLayout = {
    ...section.layout,
    margins: { ...section.layout.margins },
    columns: { ...section.layout.columns },
    breakAfter: 'nextPage',
    pageNumberStart: undefined,
  };
  const currentNode = sectionType.create(
    documentSectionNodeAttributes(currentLayout, section.id),
    Fragment.fromArray(currentChildren),
  );
  const nextNode = sectionType.create(
    documentSectionNodeAttributes(nextLayout, createWorkId('section')),
    Fragment.fromArray(followingChildren),
  );
  const transaction = commandContext.tr.replaceWith(
    section.position,
    section.position + section.node.nodeSize,
    Fragment.fromArray([currentNode, nextNode]),
  );
  const nextPosition = section.position + currentNode.nodeSize + 1;
  transaction.setSelection(
    TextSelection.near(transaction.doc.resolve(nextPosition)),
  );
  transaction.scrollIntoView();
  return true;
}

export function mergeDocumentSectionWithPrevious(
  _editor: Editor,
  commandContext: DocumentSectionCommandContext,
): boolean {
  const state = commandContext.state;
  const section = activeDocumentSectionFromState(state);
  if (!section || section.index === 0) return false;
  const sections = directDocumentSections(state.doc);
  const previous = sections[section.index - 1];
  const merged = previous.node.type.create(
    previous.node.attrs,
    previous.node.content.append(section.node.content),
    previous.node.marks,
  );
  const transaction = commandContext.tr.replaceWith(
    previous.position,
    section.position + section.node.nodeSize,
    merged,
  );
  const selectionPosition = Math.min(
    previous.position + previous.node.nodeSize - 1,
    transaction.doc.content.size,
  );
  transaction.setSelection(
    TextSelection.near(transaction.doc.resolve(selectionPosition)),
  );
  transaction.scrollIntoView();
  return true;
}

function documentSectionByIdInDocument(
  document: ProseMirrorNode,
  sectionId: string,
): ActiveDocumentSection | null {
  const sections = directDocumentSections(document);
  const index = sections.findIndex(({ node }, candidateIndex) => {
    const id =
      typeof node.attrs.id === 'string' && node.attrs.id
        ? node.attrs.id
        : `document-section-${candidateIndex + 1}`;
    return id === sectionId;
  });
  const section = sections[index];
  if (!section) return null;
  return {
    id: sectionId,
    index,
    count: sections.length,
    position: section.position,
    node: section.node,
    layout: documentSectionLayoutFromNodeAttributes(
      section.node.attrs as Partial<DocumentSectionNodeAttributes>,
    ),
  };
}

function directDocumentSections(
  document: ProseMirrorNode,
): Array<{ position: number; node: ProseMirrorNode }> {
  const sections: Array<{ position: number; node: ProseMirrorNode }> = [];
  document.forEach((node, position) => {
    if (node.type.name === 'documentSection') sections.push({ position, node });
  });
  return sections;
}

function documentSectionId(node: ProseMirrorNode, index: number): string {
  return typeof node.attrs.id === 'string' && node.attrs.id
    ? node.attrs.id
    : `document-section-${index + 1}`;
}

function sectionChildIndex(
  section: ActiveDocumentSection,
  selectionPosition: number,
): number {
  const relativePosition = Math.max(
    0,
    selectionPosition - section.position - 1,
  );
  let activeIndex = 0;
  section.node.forEach((_child, offset, index) => {
    if (relativePosition >= offset) activeIndex = index;
  });
  return activeIndex;
}
