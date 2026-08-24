import { mergeAttributes, Node, type CommandProps } from '@tiptap/core';
import { NodeSelection, Plugin, TextSelection } from '@tiptap/pm/state';
import {
  buildDocumentIndexEntries,
  DEFAULT_DOCUMENT_INDEX_OPTIONS,
  documentIndexEntryFromElement,
  documentIndexValueFromElement,
  MAX_DOCUMENT_INDEX_ENTRIES,
  normalizeDocumentIndexEntry,
  normalizeDocumentIndexEntryDraft,
  normalizeDocumentIndexOptions,
  normalizeDocumentIndexValue,
  type WorkDocumentIndexBuildOptions,
  type WorkDocumentIndexEntry,
  type WorkDocumentIndexEntryDraft,
  type WorkDocumentIndexGeneratedEntry,
  type WorkDocumentIndexOptions,
} from './work-document-index';
import { createWorkId } from './work-templates';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentIndexEntry: {
      markDocumentIndexEntry: (
        value: WorkDocumentIndexEntryDraft,
      ) => ReturnType;
      updateDocumentIndexEntry: (
        value: WorkDocumentIndexEntryDraft,
      ) => ReturnType;
    };
    documentIndex: {
      insertDocumentIndex: (
        options?: WorkDocumentIndexOptions,
        buildOptions?: WorkDocumentIndexBuildOptions,
      ) => ReturnType;
      updateDocumentIndex: (
        options: WorkDocumentIndexOptions,
        buildOptions?: WorkDocumentIndexBuildOptions,
      ) => ReturnType;
      refreshDocumentIndexes: (
        buildOptions?: WorkDocumentIndexBuildOptions,
      ) => ReturnType;
    };
  }
}

export const DocumentIndexEntry = Node.create({
  name: 'documentIndexEntry',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  draggable: false,

  addAttributes() {
    return {
      id: hiddenAttribute(''),
      mainEntry: hiddenAttribute(''),
      subEntry: hiddenAttribute(''),
      crossReference: hiddenAttribute(''),
      pageBold: hiddenAttribute(false),
      pageItalic: hiddenAttribute(false),
    };
  },

  addCommands() {
    return {
      markDocumentIndexEntry: (value) => (props) =>
        markDocumentIndexEntryCommand(props, value),
      updateDocumentIndexEntry: (value) => (props) =>
        updateDocumentIndexEntryCommand(props, value),
    };
  },

  parseHTML() {
    return [
      {
        tag: 'span[data-document-index-entry]',
        getAttrs: (node) => {
          if (!(node instanceof HTMLElement)) return false;
          const value = documentIndexEntryFromElement(node);
          return value ? indexEntryNodeAttributes(value) : false;
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const value = normalizeDocumentIndexEntry(node.attrs);
    if (!value) return ['span', { hidden: 'true' }];
    const detail = indexEntryLabel(value);
    return [
      'span',
      mergeAttributes(HTMLAttributes, {
        'data-document-index-entry': 'true',
        'data-index-entry-id': value.id,
        'data-index-main-entry': value.mainEntry,
        'data-index-sub-entry': value.subEntry,
        'data-index-cross-reference': value.crossReference,
        'data-index-page-bold': String(value.pageBold),
        'data-index-page-italic': String(value.pageItalic),
        class: 'work-document-index-entry',
        contenteditable: 'false',
        'aria-label': `索引项：${detail}`,
      }),
      ['span', { 'aria-hidden': 'true' }, '索引项'],
      ['strong', detail],
    ];
  },

  renderText() {
    return '';
  },
});

export const DocumentIndex = Node.create({
  name: 'documentIndex',
  group: 'block',
  atom: true,
  selectable: true,
  draggable: false,

  addAttributes() {
    return {
      id: hiddenAttribute(''),
      columns: hiddenAttribute(1),
      format: hiddenAttribute('indented'),
      rightAlignPageNumbers: hiddenAttribute(true),
      leader: hiddenAttribute('dot'),
      entries: hiddenAttribute([]),
      truncated: hiddenAttribute(false),
    };
  },

  addCommands() {
    return {
      insertDocumentIndex:
        (options = DEFAULT_DOCUMENT_INDEX_OPTIONS, buildOptions) =>
        (props) =>
          insertDocumentIndexCommand(props, options, buildOptions),
      updateDocumentIndex: (options, buildOptions) => (props) =>
        updateDocumentIndexCommand(props, options, buildOptions),
      refreshDocumentIndexes: (buildOptions) => (props) =>
        refreshDocumentIndexesCommand(props, buildOptions),
    };
  },

  parseHTML() {
    return [
      {
        tag: 'div[data-document-index]',
        getAttrs: (node) => {
          if (!(node instanceof HTMLElement)) return false;
          return indexNodeAttributes(documentIndexValueFromElement(node));
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const value = normalizeDocumentIndexValue({
      id: stringAttribute(node.attrs.id) || 'document-index',
      options: indexOptionsFromAttributes(node.attrs),
      entries: Array.isArray(node.attrs.entries) ? node.attrs.entries : [],
      truncated: Boolean(node.attrs.truncated),
    });
    const rows = value.entries.length
      ? value.entries.map((entry) => indexEntrySpec(entry))
      : [['li', { class: 'work-document-index-empty' }, '没有已标记的索引项']];
    return [
      'div',
      mergeAttributes(HTMLAttributes, {
        'data-document-index': 'true',
        'data-index-id': value.id,
        'data-index-columns': String(value.options.columns),
        'data-index-format': value.options.format,
        'data-index-right-align-page-numbers': String(
          value.options.rightAlignPageNumbers,
        ),
        'data-index-leader': value.options.leader,
        'data-index-entries': JSON.stringify(value.entries),
        'data-index-truncated': String(value.truncated),
        class: 'work-document-index',
        contenteditable: 'false',
        'aria-label': '索引',
      }),
      [
        'div',
        { class: 'work-document-index-header' },
        ['strong', '索引'],
        [
          'span',
          value.truncated
            ? `仅显示前 ${MAX_DOCUMENT_INDEX_ENTRIES} 项`
            : `${value.entries.length} 项`,
        ],
      ],
      ['ol', { class: 'work-document-index-list' }, ...rows],
    ];
  },

  renderText({ node }) {
    const entries = Array.isArray(node.attrs.entries)
      ? (node.attrs.entries as WorkDocumentIndexGeneratedEntry[])
      : [];
    return entries
      .map((entry) =>
        [entry.mainEntry, entry.subEntry, entry.crossReference]
          .filter(Boolean)
          .join(' '),
      )
      .join('\n');
  },

  addProseMirrorPlugins() {
    return [
      new Plugin({
        props: {
          handleDOMEvents: {
            click: (view, event) => {
              if (!(event.target instanceof Element)) return false;
              const link = event.target.closest<HTMLElement>(
                '[data-document-index] [data-index-target]',
              );
              const targetId = link?.dataset.indexTarget?.trim();
              if (!link || !targetId) return false;
              let targetPosition: number | null = null;
              view.state.doc.descendants((node, position) => {
                if (targetPosition !== null) return false;
                if (
                  node.type.name === 'documentIndexEntry' &&
                  node.attrs.id === targetId
                ) {
                  targetPosition = position;
                  return false;
                }
                return true;
              });
              if (targetPosition === null) return false;
              event.preventDefault();
              view.dispatch(
                view.state.tr
                  .setSelection(
                    NodeSelection.create(view.state.doc, targetPosition),
                  )
                  .scrollIntoView(),
              );
              view.focus();
              return true;
            },
          },
        },
      }),
    ];
  },
});

export function selectedDocumentIndexEntry(editor: {
  state: CommandProps['state'];
}): WorkDocumentIndexEntry | null {
  const selection = editor.state.selection;
  if (
    !(selection instanceof NodeSelection) ||
    selection.node.type.name !== 'documentIndexEntry'
  ) {
    return null;
  }
  return normalizeDocumentIndexEntry(selection.node.attrs);
}

export function selectedDocumentIndexOptions(editor: {
  state: CommandProps['state'];
}): WorkDocumentIndexOptions | null {
  const selection = editor.state.selection;
  if (
    !(selection instanceof NodeSelection) ||
    selection.node.type.name !== 'documentIndex'
  ) {
    return null;
  }
  return indexOptionsFromAttributes(selection.node.attrs);
}

export function selectedDocumentIndexDraft(editor: {
  state: CommandProps['state'];
}): WorkDocumentIndexEntryDraft | null {
  const selected = selectedDocumentIndexEntry(editor);
  if (selected) {
    return {
      mainEntry: selected.mainEntry,
      subEntry: selected.subEntry,
      crossReference: selected.crossReference,
      pageBold: selected.pageBold,
      pageItalic: selected.pageItalic,
    };
  }
  const { from, to } = editor.state.selection;
  const mainEntry = editor.state.doc
    .textBetween(from, to, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return normalizeDocumentIndexEntryDraft({ mainEntry });
}

export function documentHasIndexEntries(editor: {
  state: CommandProps['state'];
}): boolean {
  return documentHasNode(editor, 'documentIndexEntry');
}

export function documentHasIndex(editor: {
  state: CommandProps['state'];
}): boolean {
  return documentHasNode(editor, 'documentIndex');
}

function markDocumentIndexEntryCommand(
  { dispatch, editor, state, tr }: CommandProps,
  source: WorkDocumentIndexEntryDraft,
): boolean {
  const value = normalizeDocumentIndexEntryDraft(source);
  const nodeType = editor.schema.nodes.documentIndexEntry;
  if (!value || !nodeType) return false;
  const selection = state.selection;
  if (
    selection instanceof NodeSelection &&
    selection.node.type.name === 'documentIndexEntry'
  ) {
    return updateDocumentIndexEntryCommand(
      { dispatch, editor, state, tr } as CommandProps,
      value,
    );
  }
  if (!dispatch) return true;
  const position = selection.to;
  const node = nodeType.create({ id: createWorkId('index-entry'), ...value });
  tr.insert(position, node);
  tr.setSelection(TextSelection.near(tr.doc.resolve(position + node.nodeSize)));
  tr.scrollIntoView();
  return tr.docChanged;
}

function updateDocumentIndexEntryCommand(
  { dispatch, state, tr }: CommandProps,
  source: WorkDocumentIndexEntryDraft,
): boolean {
  const selection = state.selection;
  const value = normalizeDocumentIndexEntryDraft(source);
  if (
    !value ||
    !(selection instanceof NodeSelection) ||
    selection.node.type.name !== 'documentIndexEntry'
  ) {
    return false;
  }
  if (!dispatch) return true;
  tr.setNodeMarkup(selection.from, undefined, {
    ...selection.node.attrs,
    ...value,
  });
  tr.scrollIntoView();
  return tr.docChanged;
}

function insertDocumentIndexCommand(
  { dispatch, editor, tr }: CommandProps,
  source: WorkDocumentIndexOptions,
  buildOptions: WorkDocumentIndexBuildOptions = {},
): boolean {
  const nodeType = editor.schema.nodes.documentIndex;
  if (!nodeType) return false;
  if (!dispatch) return true;
  const options = normalizeDocumentIndexOptions(source);
  const built = buildDocumentIndexEntries(tr.doc, buildOptions);
  tr.replaceSelectionWith(
    nodeType.create({ id: createWorkId('index'), ...options, ...built }),
    false,
  );
  tr.scrollIntoView();
  return true;
}

function updateDocumentIndexCommand(
  { dispatch, state, tr }: CommandProps,
  source: WorkDocumentIndexOptions,
  buildOptions: WorkDocumentIndexBuildOptions = {},
): boolean {
  const selection = state.selection;
  if (
    !(selection instanceof NodeSelection) ||
    selection.node.type.name !== 'documentIndex'
  ) {
    return false;
  }
  if (!dispatch) return true;
  const options = normalizeDocumentIndexOptions(source);
  const built = buildDocumentIndexEntries(tr.doc, buildOptions);
  tr.setNodeMarkup(selection.from, undefined, {
    ...selection.node.attrs,
    ...options,
    ...built,
  });
  tr.scrollIntoView();
  return tr.docChanged;
}

function refreshDocumentIndexesCommand(
  { dispatch, state, tr }: CommandProps,
  buildOptions: WorkDocumentIndexBuildOptions = {},
): boolean {
  const updates: Array<{ position: number; attrs: Record<string, unknown> }> =
    [];
  const built = buildDocumentIndexEntries(state.doc, buildOptions);
  state.doc.descendants((node, position) => {
    if (node.type.name !== 'documentIndex') return;
    if (
      JSON.stringify(node.attrs.entries) === JSON.stringify(built.entries) &&
      Boolean(node.attrs.truncated) === built.truncated
    ) {
      return;
    }
    updates.push({
      position,
      attrs: { ...node.attrs, ...built },
    });
  });
  if (!updates.length) return false;
  if (!dispatch) return true;
  for (const update of updates) {
    tr.setNodeMarkup(update.position, undefined, update.attrs);
  }
  tr.scrollIntoView();
  return tr.docChanged;
}

function indexEntryNodeAttributes(
  value: WorkDocumentIndexEntry,
): Record<string, unknown> {
  return { ...value };
}

function indexNodeAttributes(
  value: ReturnType<typeof documentIndexValueFromElement>,
): Record<string, unknown> {
  return {
    id: value.id,
    ...value.options,
    entries: value.entries,
    truncated: value.truncated,
  };
}

function indexOptionsFromAttributes(
  attrs: Record<string, unknown>,
): WorkDocumentIndexOptions {
  return normalizeDocumentIndexOptions({
    columns: Number(attrs.columns),
    format: attrs.format === 'run-in' ? 'run-in' : 'indented',
    rightAlignPageNumbers: Boolean(attrs.rightAlignPageNumbers),
    leader:
      typeof attrs.leader === 'string'
        ? (attrs.leader as WorkDocumentIndexOptions['leader'])
        : 'dot',
  });
}

function indexEntrySpec(entry: WorkDocumentIndexGeneratedEntry): unknown[] {
  const term: unknown[] = [
    ['span', { class: 'work-document-index-main' }, entry.mainEntry],
  ];
  if (entry.subEntry) {
    term.push(['span', { class: 'work-document-index-sub' }, entry.subEntry]);
  }
  const pages: unknown[] = entry.crossReference
    ? [
        [
          'span',
          { class: 'work-document-index-cross-reference' },
          `参见 ${entry.crossReference}`,
        ],
      ]
    : entry.pages.flatMap((page, index) => {
        const targetId = page.targetIds[0] ?? '';
        const result: unknown[] = [];
        if (index > 0) result.push(['span', { 'aria-hidden': 'true' }, ', ']);
        result.push([
          'a',
          {
            href: `#${targetId}`,
            'data-index-target': targetId,
            tabindex: '-1',
            class:
              [page.pageBold ? 'bold' : '', page.pageItalic ? 'italic' : '']
                .filter(Boolean)
                .join(' ') || undefined,
          },
          String(page.pageNumber),
        ]);
        return result;
      });
  return [
    'li',
    {
      'data-index-main-entry': entry.mainEntry,
      'data-index-sub-entry': entry.subEntry,
    },
    ...term,
    ['i', { 'aria-hidden': 'true' }],
    ['span', { class: 'work-document-index-pages' }, ...pages],
  ];
}

function indexEntryLabel(value: WorkDocumentIndexEntryDraft): string {
  const term = value.subEntry
    ? `${value.mainEntry} › ${value.subEntry}`
    : value.mainEntry;
  return value.crossReference ? `${term} · 参见 ${value.crossReference}` : term;
}

function documentHasNode(
  editor: { state: CommandProps['state'] },
  typeName: string,
): boolean {
  let found = false;
  editor.state.doc.descendants((node) => {
    if (node.type.name === typeName) found = true;
    return !found;
  });
  return found;
}

function hiddenAttribute(defaultValue: unknown) {
  return { default: defaultValue, rendered: false };
}

function stringAttribute(value: unknown): string {
  return typeof value === 'string' ? value.trim() : '';
}
