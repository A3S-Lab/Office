import { mergeAttributes, Node, type CommandProps } from '@tiptap/core';
import {
  NodeSelection,
  Plugin,
  TextSelection,
  type Transaction,
} from '@tiptap/pm/state';
import {
  collectWorkDocumentOutline,
  workDocumentOutlineLevel,
} from './work-document-outline';
import {
  createDocumentParagraphIdentity,
  createDocumentParagraphIdentityRegistry,
  normalizeDocumentParagraphId,
  uniqueDocumentParagraphIdentity,
  type WorkDocumentParagraphIdentity,
} from './work-document-paragraph-identity';
import {
  buildDocumentTableOfContentsEntries,
  DEFAULT_DOCUMENT_TABLE_OF_CONTENTS_OPTIONS,
  documentTableOfContentsValueFromElement,
  normalizeDocumentTableOfContentsOptions,
  normalizeDocumentTableOfContentsValue,
  type WorkDocumentTableOfContentsBuildOptions,
  type WorkDocumentTableOfContentsEntry,
  type WorkDocumentTableOfContentsOptions,
} from './work-document-table-of-contents';
import { createWorkId } from './work-templates';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentTableOfContents: {
      insertDocumentTableOfContents: (
        options?: WorkDocumentTableOfContentsOptions,
        buildOptions?: WorkDocumentTableOfContentsBuildOptions,
      ) => ReturnType;
      updateDocumentTableOfContents: (
        options: WorkDocumentTableOfContentsOptions,
        buildOptions?: WorkDocumentTableOfContentsBuildOptions,
      ) => ReturnType;
      refreshDocumentTablesOfContents: (
        buildOptions?: WorkDocumentTableOfContentsBuildOptions,
      ) => ReturnType;
    };
  }
}

export const DocumentTableOfContents = Node.create({
  name: 'documentTableOfContents',
  group: 'block',
  atom: true,
  selectable: true,
  draggable: false,

  addCommands() {
    return {
      insertDocumentTableOfContents:
        (options = DEFAULT_DOCUMENT_TABLE_OF_CONTENTS_OPTIONS, buildOptions) =>
        (props) =>
          insertDocumentTableOfContentsCommand(props, options, buildOptions),
      updateDocumentTableOfContents: (options, buildOptions) => (props) =>
        updateDocumentTableOfContentsCommand(props, options, buildOptions),
      refreshDocumentTablesOfContents: (buildOptions) => (props) =>
        refreshDocumentTablesOfContentsCommand(props, buildOptions),
    };
  },

  addAttributes() {
    return {
      id: hiddenAttribute(''),
      minLevel: hiddenAttribute(1),
      maxLevel: hiddenAttribute(3),
      hyperlinks: hiddenAttribute(true),
      showPageNumbers: hiddenAttribute(true),
      rightAlignPageNumbers: hiddenAttribute(true),
      leader: hiddenAttribute('dot'),
      entries: hiddenAttribute([]),
      truncated: hiddenAttribute(false),
    };
  },

  parseHTML() {
    return [
      {
        tag: 'div[data-document-table-of-contents]',
        getAttrs: (node) => {
          if (!(node instanceof HTMLElement)) return false;
          return tableOfContentsNodeAttributes(
            documentTableOfContentsValueFromElement(node),
          );
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const value = normalizeDocumentTableOfContentsValue({
      id: stringAttribute(node.attrs.id) || 'document-table-of-contents',
      options: tableOfContentsOptionsFromAttributes(node.attrs),
      entries: Array.isArray(node.attrs.entries) ? node.attrs.entries : [],
      truncated: Boolean(node.attrs.truncated),
    });
    const options = value.options;
    const rows = value.entries.length
      ? value.entries.map((entry) => tableOfContentsEntrySpec(entry, options))
      : [
          [
            'li',
            { class: 'work-document-table-of-contents-empty' },
            '没有符合级别范围的标题',
          ],
        ];
    return [
      'div',
      mergeAttributes(HTMLAttributes, {
        'data-document-table-of-contents': 'true',
        'data-toc-id': value.id,
        'data-toc-min-level': String(options.minLevel),
        'data-toc-max-level': String(options.maxLevel),
        'data-toc-hyperlinks': String(options.hyperlinks),
        'data-toc-show-page-numbers': String(options.showPageNumbers),
        'data-toc-right-align-page-numbers': String(
          options.rightAlignPageNumbers,
        ),
        'data-toc-leader': options.leader,
        'data-toc-entries': JSON.stringify(value.entries),
        'data-toc-truncated': String(value.truncated),
        class: 'work-document-table-of-contents',
        contenteditable: 'false',
        'aria-label': '目录',
      }),
      [
        'div',
        { class: 'work-document-table-of-contents-header' },
        ['strong', '目录'],
        [
          'span',
          value.truncated ? '仅显示前 512 项' : `${value.entries.length} 项`,
        ],
      ],
      ['ol', { class: 'work-document-table-of-contents-list' }, ...rows],
    ];
  },

  renderText({ node }) {
    const entries = Array.isArray(node.attrs.entries)
      ? (node.attrs.entries as WorkDocumentTableOfContentsEntry[])
      : [];
    return entries.map((entry) => entry.title).join('\n');
  },

  addProseMirrorPlugins() {
    return [
      new Plugin({
        props: {
          handleDOMEvents: {
            click: (view, event) => {
              if (!(event.target instanceof Element)) return false;
              const table = event.target.closest<HTMLElement>(
                '[data-document-table-of-contents][data-toc-hyperlinks="true"]',
              );
              if (!table) return false;
              const link =
                event.target.closest<HTMLElement>('[data-toc-target]');
              if (!link || !table.contains(link)) return false;
              const targetId = link.dataset.tocTarget?.trim();
              if (!targetId) return false;
              const target = collectWorkDocumentOutline(view.state.doc).find(
                (item) => item.id === targetId,
              );
              if (!target) return false;
              event.preventDefault();
              view.dispatch(
                view.state.tr
                  .setSelection(
                    TextSelection.near(view.state.doc.resolve(target.from)),
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

export function selectedDocumentTableOfContentsOptions(editor: {
  state: CommandProps['state'];
}): WorkDocumentTableOfContentsOptions | null {
  const selection = editor.state.selection;
  if (
    !(selection instanceof NodeSelection) ||
    selection.node.type.name !== 'documentTableOfContents'
  ) {
    return null;
  }
  return tableOfContentsOptionsFromAttributes(selection.node.attrs);
}

export function documentHasTableOfContents(editor: {
  state: CommandProps['state'];
}): boolean {
  let found = false;
  editor.state.doc.descendants((node) => {
    if (node.type.name === 'documentTableOfContents') found = true;
    return !found;
  });
  return found;
}

function insertDocumentTableOfContentsCommand(
  { dispatch, editor, tr }: CommandProps,
  source: WorkDocumentTableOfContentsOptions,
  buildOptions: WorkDocumentTableOfContentsBuildOptions = {},
): boolean {
  const nodeType = editor.schema.nodes.documentTableOfContents;
  if (!nodeType) return false;
  if (!dispatch) return true;
  const options = normalizeDocumentTableOfContentsOptions(source);
  ensureTableOfContentsHeadingIdentities(tr, [options]);
  const built = buildDocumentTableOfContentsEntries(
    tr.doc,
    options,
    buildOptions,
  );
  tr.replaceSelectionWith(
    nodeType.create({
      id: createWorkId('toc'),
      ...options,
      ...built,
    }),
    false,
  );
  tr.scrollIntoView();
  return true;
}

function updateDocumentTableOfContentsCommand(
  { dispatch, state, tr }: CommandProps,
  source: WorkDocumentTableOfContentsOptions,
  buildOptions: WorkDocumentTableOfContentsBuildOptions = {},
): boolean {
  const selection = state.selection;
  if (
    !(selection instanceof NodeSelection) ||
    selection.node.type.name !== 'documentTableOfContents'
  ) {
    return false;
  }
  if (!dispatch) return true;
  const options = normalizeDocumentTableOfContentsOptions(source);
  ensureTableOfContentsHeadingIdentities(tr, [options]);
  const built = buildDocumentTableOfContentsEntries(
    tr.doc,
    options,
    buildOptions,
  );
  tr.setNodeMarkup(selection.from, undefined, {
    ...selection.node.attrs,
    ...options,
    ...built,
  });
  tr.scrollIntoView();
  return tr.docChanged;
}

function refreshDocumentTablesOfContentsCommand(
  { dispatch, state, tr }: CommandProps,
  buildOptions: WorkDocumentTableOfContentsBuildOptions = {},
): boolean {
  const tableOptions: WorkDocumentTableOfContentsOptions[] = [];
  state.doc.descendants((node) => {
    if (node.type.name === 'documentTableOfContents') {
      tableOptions.push(tableOfContentsOptionsFromAttributes(node.attrs));
    }
  });
  if (!tableOptions.length) return false;
  ensureTableOfContentsHeadingIdentities(tr, tableOptions);
  const updates: Array<{
    position: number;
    attrs: Record<string, unknown>;
  }> = [];
  tr.doc.descendants((node, position) => {
    if (node.type.name !== 'documentTableOfContents') return;
    const options = tableOfContentsOptionsFromAttributes(node.attrs);
    const built = buildDocumentTableOfContentsEntries(
      tr.doc,
      options,
      buildOptions,
    );
    if (
      JSON.stringify(node.attrs.entries) === JSON.stringify(built.entries) &&
      Boolean(node.attrs.truncated) === built.truncated
    ) {
      return;
    }
    updates.push({
      position,
      attrs: { ...node.attrs, ...options, ...built },
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

function ensureTableOfContentsHeadingIdentities(
  transaction: Transaction,
  tableOptions: readonly WorkDocumentTableOfContentsOptions[],
): void {
  const paragraphIdCounts = new Map<string, number>();
  transaction.doc.descendants((node) => {
    const paragraphId = normalizeDocumentParagraphId(node.attrs.paragraphId);
    if (!paragraphId) return;
    paragraphIdCounts.set(
      paragraphId,
      (paragraphIdCounts.get(paragraphId) ?? 0) + 1,
    );
  });
  const registry = createDocumentParagraphIdentityRegistry();
  for (const paragraphId of paragraphIdCounts.keys()) {
    registry.paragraphIds.add(paragraphId);
  }
  const updates: Array<{
    attrs: Record<string, unknown>;
    identity: WorkDocumentParagraphIdentity;
    position: number;
  }> = [];
  transaction.doc.descendants((node, position) => {
    const level = workDocumentOutlineLevel(node);
    if (
      level === null ||
      !tableOptions.some(
        (options) => level >= options.minLevel && level <= options.maxLevel,
      )
    ) {
      return;
    }
    const paragraphId = normalizeDocumentParagraphId(node.attrs.paragraphId);
    const textId = normalizeDocumentParagraphId(node.attrs.textId);
    let identity: WorkDocumentParagraphIdentity;
    if (paragraphId && paragraphIdCounts.get(paragraphId) === 1) {
      identity = {
        paragraphId,
        textId: textId ?? createDocumentParagraphIdentity().textId,
      };
    } else {
      identity = uniqueDocumentParagraphIdentity({}, registry);
    }
    if (paragraphId === identity.paragraphId && textId === identity.textId) {
      return;
    }
    updates.push({ attrs: node.attrs, identity, position });
  });
  for (const update of updates) {
    transaction.setNodeMarkup(update.position, undefined, {
      ...update.attrs,
      ...update.identity,
    });
  }
}

function tableOfContentsOptionsFromAttributes(
  attrs: Record<string, unknown>,
): WorkDocumentTableOfContentsOptions {
  return normalizeDocumentTableOfContentsOptions({
    minLevel: Number(attrs.minLevel),
    maxLevel: Number(attrs.maxLevel),
    hyperlinks: Boolean(attrs.hyperlinks),
    showPageNumbers: Boolean(attrs.showPageNumbers),
    rightAlignPageNumbers: Boolean(attrs.rightAlignPageNumbers),
    leader:
      typeof attrs.leader === 'string'
        ? (attrs.leader as WorkDocumentTableOfContentsOptions['leader'])
        : 'dot',
  });
}

function tableOfContentsNodeAttributes(
  value: ReturnType<typeof documentTableOfContentsValueFromElement>,
): Record<string, unknown> {
  return {
    id: value.id,
    ...value.options,
    entries: value.entries,
    truncated: value.truncated,
  };
}

function tableOfContentsEntrySpec(
  entry: WorkDocumentTableOfContentsEntry,
  options: WorkDocumentTableOfContentsOptions,
): unknown[] {
  const title = options.hyperlinks
    ? [
        'a',
        {
          href: `#${entry.targetId}`,
          'data-toc-target': entry.targetId,
          tabindex: '-1',
        },
        entry.title,
      ]
    : ['span', entry.title];
  return [
    'li',
    {
      'data-toc-target': entry.targetId,
      'data-toc-level': String(entry.level),
      style: `--work-toc-level:${entry.level}`,
    },
    title,
    ['i', { 'aria-hidden': 'true' }],
    ...(options.showPageNumbers
      ? [
          [
            'span',
            { class: 'work-document-table-of-contents-page' },
            String(entry.pageNumber),
          ],
        ]
      : []),
  ];
}

function hiddenAttribute(defaultValue: unknown) {
  return { default: defaultValue, rendered: false };
}

function stringAttribute(value: unknown): string {
  return typeof value === 'string' ? value.trim() : '';
}
