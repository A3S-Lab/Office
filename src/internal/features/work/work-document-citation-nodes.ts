import {
  type CommandProps,
  type Editor,
  mergeAttributes,
  Node,
} from '@tiptap/core';
import {
  DOMParser as ProseMirrorDOMParser,
  type Node as ProseMirrorNode,
} from '@tiptap/pm/model';
import { NodeSelection } from '@tiptap/pm/state';
import {
  documentCitationInstruction,
  renameDocumentCitationTagInInstruction,
  documentCitationTags,
  documentCitationTagsFromInstruction,
  renderDocumentBibliographyHtml,
  resolveDocumentCitation,
} from './work-document-citations';
import { createWorkId } from './work-templates';
import type {
  WorkDocumentBibliography,
  WorkDocumentCitationSource,
  WorkDocumentContent,
} from './work-types';

export interface DocumentCitationTagRename {
  next: string;
  previous: string;
}

interface DocumentCitationOptions {
  getContent: () => WorkDocumentContent | null;
  onContentChange: (content: WorkDocumentContent) => void;
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentCitation: {
      insertDocumentCitation: (
        source: WorkDocumentCitationSource,
        bibliography: WorkDocumentBibliography,
      ) => ReturnType;
      insertDocumentBibliography: (
        bibliography: WorkDocumentBibliography,
      ) => ReturnType;
      refreshDocumentCitations: (content: WorkDocumentContent) => ReturnType;
      renameDocumentCitationTag: (
        previousTag: string,
        nextTag: string,
      ) => ReturnType;
      setDocumentBibliography: (
        bibliography: WorkDocumentBibliography,
        renamedTag?: DocumentCitationTagRename,
      ) => ReturnType;
    };
  }
}

export const DocumentCitation = Node.create<DocumentCitationOptions>({
  name: 'documentCitation',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: true,

  addOptions() {
    return {
      getContent: () => null,
      onContentChange: () => undefined,
    };
  },

  addCommands() {
    return {
      insertDocumentCitation: (source, bibliography) => (props) =>
        insertDocumentCitationCommand(props, source, bibliography),
      insertDocumentBibliography: (bibliography) => (props) =>
        insertDocumentBibliographyCommand(props, bibliography),
      refreshDocumentCitations: (content) => (props) =>
        refreshDocumentCitationsCommand(props, content),
      renameDocumentCitationTag: (previousTag, nextTag) => (props) =>
        renameDocumentCitationTagCommand(props, previousTag, nextTag),
      setDocumentBibliography:
        (bibliography, renamedTag) =>
        ({ dispatch, editor }) => {
          const content = this.options.getContent();
          if (!content) return false;
          if (!dispatch) return true;
          const next = { ...content, bibliography };
          this.options.onContentChange(next);
          const chain = editor.chain();
          if (renamedTag) {
            chain.renameDocumentCitationTag(
              renamedTag.previous,
              renamedTag.next,
            );
          }
          chain.refreshDocumentCitations(next).run();
          return true;
        },
    };
  },

  addAttributes() {
    return {
      id: hiddenAttribute(''),
      tags: hiddenAttribute(''),
      instruction: hiddenAttribute(''),
      display: hiddenAttribute(''),
      orphaned: hiddenAttribute(false),
    };
  },

  parseHTML() {
    return [
      {
        tag: 'span[data-document-citation]',
        getAttrs: (node) => {
          if (!(node instanceof HTMLElement)) return false;
          const instruction = node.dataset.citationInstruction?.trim() ?? '';
          const tags = documentCitationTags(node.dataset.citationTags);
          const resolvedTags = tags.length
            ? tags
            : documentCitationTagsFromInstruction(instruction);
          return {
            id: node.dataset.citationId ?? '',
            tags: resolvedTags.join(' '),
            instruction:
              instruction || documentCitationInstruction(resolvedTags),
            display:
              node.dataset.citationDisplay?.trim() ||
              node.textContent?.trim() ||
              '缺失引文',
            orphaned: node.dataset.citationOrphaned === 'true',
          };
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const tags =
      typeof node.attrs.tags === 'string'
        ? documentCitationTags(node.attrs.tags)
        : [];
    const instruction =
      typeof node.attrs.instruction === 'string'
        ? node.attrs.instruction.trim()
        : '';
    const display =
      typeof node.attrs.display === 'string' && node.attrs.display.trim()
        ? node.attrs.display.trim()
        : '缺失引文';
    const orphaned = Boolean(node.attrs.orphaned);
    return [
      'span',
      mergeAttributes(HTMLAttributes, {
        'data-document-citation': 'true',
        'data-citation-id':
          typeof node.attrs.id === 'string' ? node.attrs.id : '',
        'data-citation-tags': tags.join(' '),
        'data-citation-instruction':
          instruction || documentCitationInstruction(tags),
        'data-citation-display': display,
        'data-citation-orphaned': orphaned ? 'true' : undefined,
        class: 'work-document-citation',
        title: tags.length ? `引文：${tags.join('、')}` : '缺失引文',
      }),
      display,
    ];
  },

  renderText({ node }) {
    return typeof node.attrs.display === 'string' && node.attrs.display.trim()
      ? node.attrs.display.trim()
      : '缺失引文';
  },
});

export const DocumentBibliography = Node.create({
  name: 'documentBibliography',
  group: 'block',
  content: 'block+',
  atom: true,
  selectable: true,
  defining: true,
  isolating: true,

  addAttributes() {
    return {
      id: hiddenAttribute('document-bibliography-1'),
      style: hiddenAttribute('apa'),
    };
  },

  parseHTML() {
    return [
      {
        tag: 'section[data-document-bibliography]',
        getAttrs: (node) => {
          if (!(node instanceof HTMLElement)) return false;
          return {
            id: node.dataset.bibliographyId || 'document-bibliography-1',
            style: node.dataset.bibliographyStyle || 'apa',
          };
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    return [
      'section',
      mergeAttributes(HTMLAttributes, {
        'data-document-bibliography': 'true',
        'data-bibliography-id':
          typeof node.attrs.id === 'string'
            ? node.attrs.id
            : 'document-bibliography-1',
        'data-bibliography-style':
          typeof node.attrs.style === 'string' ? node.attrs.style : 'apa',
        class: 'work-document-bibliography',
      }),
      0,
    ];
  },
});

export function documentCitationCount(editor: Editor): number {
  let count = 0;
  editor.state.doc.descendants((node) => {
    if (node.type.name === 'documentCitation') count += 1;
  });
  return count;
}

function insertDocumentCitationCommand(
  { dispatch, editor, state, tr }: CommandProps,
  source: WorkDocumentCitationSource,
  bibliography: WorkDocumentBibliography,
): boolean {
  const citationType = editor.schema.nodes.documentCitation;
  if (
    !citationType ||
    !source.tag.trim() ||
    selectionInsideNode(state, 'documentBibliography')
  ) {
    return false;
  }
  if (!dispatch) return true;
  const instruction = documentCitationInstruction([source.tag]);
  const resolved = resolveDocumentCitation(
    [source.tag],
    bibliography,
    instruction,
  );
  tr.replaceSelectionWith(
    citationType.create({
      id: createWorkId('citation'),
      tags: source.tag,
      instruction,
      display: resolved.text,
      orphaned: resolved.orphaned,
    }),
    false,
  );
  tr.scrollIntoView();
  return true;
}

function insertDocumentBibliographyCommand(
  { dispatch, editor, state, tr }: CommandProps,
  bibliography: WorkDocumentBibliography,
): boolean {
  if (!editor.schema.nodes.documentBibliography) return false;
  const existing = documentBibliographyNodes(state.doc)[0];
  if (!dispatch) return true;

  if (existing) {
    refreshDocumentCitationsInTransaction(editor, tr, bibliography);
    tr.setSelection(NodeSelection.create(tr.doc, existing.position));
    tr.scrollIntoView();
    return true;
  }

  const replacement = bibliographyNode(
    editor,
    bibliography,
    createWorkId('bibliography'),
  );
  if (!replacement) return false;
  tr.replaceSelectionWith(replacement);
  const paragraphType = editor.schema.nodes.paragraph;
  if (paragraphType && tr.selection.$to.parent.type === replacement.type) {
    tr.insert(tr.selection.to, paragraphType.create());
  }
  tr.scrollIntoView();
  return true;
}

function refreshDocumentCitationsCommand(
  { editor, tr }: CommandProps,
  content: WorkDocumentContent,
): boolean {
  return refreshDocumentCitationsInTransaction(
    editor,
    tr,
    content.bibliography,
  );
}

function renameDocumentCitationTagCommand(
  { tr }: CommandProps,
  previousTag: string,
  nextTag: string,
): boolean {
  if (!previousTag || previousTag === nextTag) return false;
  const document = tr.doc;
  document.descendants((node, position) => {
    if (node.type.name !== 'documentCitation') return;
    const tags = documentCitationTags(
      typeof node.attrs.tags === 'string' ? node.attrs.tags : '',
    );
    if (!tags.includes(previousTag)) return;
    const renamed = tags.map((tag) => (tag === previousTag ? nextTag : tag));
    const instruction =
      typeof node.attrs.instruction === 'string'
        ? renameDocumentCitationTagInInstruction(
            node.attrs.instruction,
            previousTag,
            nextTag,
          )
        : documentCitationInstruction(renamed);
    tr.setNodeMarkup(position, undefined, {
      ...node.attrs,
      tags: renamed.join(' '),
      instruction,
    });
  });
  return tr.docChanged;
}

function refreshDocumentCitationsInTransaction(
  editor: Editor,
  tr: CommandProps['tr'],
  bibliography: WorkDocumentBibliography | undefined,
): boolean {
  const document = tr.doc;
  const bibliographyNodes: Array<{
    node: ProseMirrorNode;
    position: number;
  }> = [];
  document.descendants((node, position) => {
    if (node.type.name === 'documentCitation') {
      const tags = documentCitationTags(
        typeof node.attrs.tags === 'string' ? node.attrs.tags : '',
      );
      const instruction =
        typeof node.attrs.instruction === 'string'
          ? node.attrs.instruction
          : '';
      const cached =
        typeof node.attrs.display === 'string' ? node.attrs.display : '';
      const resolved = resolveDocumentCitation(
        tags,
        bibliography,
        instruction,
        cached,
      );
      if (
        node.attrs.display !== resolved.text ||
        Boolean(node.attrs.orphaned) !== resolved.orphaned
      ) {
        tr.setNodeMarkup(position, undefined, {
          ...node.attrs,
          display: resolved.text,
          orphaned: resolved.orphaned,
        });
      }
      return;
    }
    if (node.type.name === 'documentBibliography') {
      bibliographyNodes.push({ node, position });
    }
  });
  if (bibliography) {
    for (const current of bibliographyNodes.reverse()) {
      const id =
        typeof current.node.attrs.id === 'string'
          ? current.node.attrs.id
          : 'document-bibliography-1';
      const replacement = bibliographyNode(editor, bibliography, id);
      if (replacement && !current.node.eq(replacement)) {
        tr.replaceWith(
          current.position,
          current.position + current.node.nodeSize,
          replacement,
        );
      }
    }
  }
  return tr.docChanged;
}

function bibliographyNode(
  editor: Editor,
  bibliography: WorkDocumentBibliography,
  id: string,
): ProseMirrorNode | null {
  const document = new DOMParser().parseFromString(
    renderDocumentBibliographyHtml(bibliography, id),
    'text/html',
  );
  const slice = ProseMirrorDOMParser.fromSchema(editor.schema).parseSlice(
    document.body,
  );
  return slice.content.firstChild?.type.name === 'documentBibliography'
    ? slice.content.firstChild
    : null;
}

function documentBibliographyNodes(
  document: ProseMirrorNode,
): Array<{ node: ProseMirrorNode; position: number }> {
  const result: Array<{ node: ProseMirrorNode; position: number }> = [];
  document.descendants((node, position) => {
    if (node.type.name === 'documentBibliography') {
      result.push({ node, position });
    }
  });
  return result;
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
