import {
  type CommandProps,
  type Editor,
  mergeAttributes,
  Node,
} from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import { Fragment, type Node as ProseMirrorNode } from '@tiptap/pm/model';
import {
  NodeSelection,
  Plugin,
  TextSelection,
  type Transaction,
} from '@tiptap/pm/state';
import { createWorkId } from './work-templates';

/**
 * The deliberately small content-control surface exposed by Writer.
 *
 * Data binding, repeating sections, dates, drop-downs, pictures, and other
 * active form semantics are intentionally not represented by this model. A
 * control is an inline, editable container whose metadata can round-trip to a
 * passive Word `w:sdt`.
 */
export type WorkDocumentContentControlType = 'text' | 'richText';
export type WorkDocumentContentControlLock =
  | 'unlocked'
  | 'contentLocked'
  | 'sdtLocked'
  | 'sdtContentLocked';
export type WorkDocumentContentControlAppearance =
  | 'boundingBox'
  | 'hidden'
  | 'tags';

export interface WorkDocumentContentControlProperties {
  id: string;
  nativeId: number | null;
  type: WorkDocumentContentControlType;
  alias: string;
  tag: string;
  lock: WorkDocumentContentControlLock;
  multiLine: boolean;
  appearance: WorkDocumentContentControlAppearance;
  color: string | null;
}

export interface DocumentContentControlInsertOptions
  extends Partial<WorkDocumentContentControlProperties> {
  /** Initial plain text for an empty selection. */
  text?: string;
}

export interface DocumentContentControlCommandOptions {
  restoreFocus?: boolean;
  /** Explicitly permits deleting a control whose `w:sdt` is locked. */
  allowLocked?: boolean;
}

export interface SelectedDocumentContentControl {
  node: ProseMirrorNode;
  position: number;
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentContentControl: {
      insertDocumentContentControl: (
        options?: DocumentContentControlInsertOptions,
      ) => ReturnType;
      setDocumentContentControlProperties: (
        value: Partial<WorkDocumentContentControlProperties>,
        options?: DocumentContentControlCommandOptions,
      ) => ReturnType;
      deleteDocumentContentControl: (
        options?: DocumentContentControlCommandOptions,
      ) => ReturnType;
    };
  }
}

export const DOCUMENT_CONTENT_CONTROL_DEFAULTS: WorkDocumentContentControlProperties =
  {
    id: '',
    nativeId: null,
    type: 'text',
    alias: '',
    tag: '',
    lock: 'unlocked',
    multiLine: false,
    appearance: 'boundingBox',
    color: null,
  };

/** Transaction meta used by the typed commands to cross the lock guard. */
export const DOCUMENT_CONTENT_CONTROL_MUTATION_META =
  'documentContentControlMutation';

const CONTENT_CONTROL_ID_MAX_LENGTH = 160;
const CONTENT_CONTROL_STRING_MAX_LENGTH = 255;
const CONTENT_CONTROL_COLOR_PATTERN = /^#[0-9a-f]{6}$/i;
const CONTENT_CONTROL_MARKER_ATTRIBUTES = [
  'id',
  'nativeId',
  'type',
  'alias',
  'tag',
  'lock',
  'multiLine',
  'appearance',
  'color',
] as const;

export const DocumentContentControl = Node.create({
  name: 'documentContentControl',
  inline: true,
  group: 'inline',
  content: 'inline*',
  defining: true,
  isolating: true,
  selectable: true,

  addAttributes() {
    return {
      id: hiddenAttribute(''),
      nativeId: hiddenAttribute(null),
      type: hiddenAttribute(DOCUMENT_CONTENT_CONTROL_DEFAULTS.type),
      alias: hiddenAttribute(''),
      tag: hiddenAttribute(''),
      lock: hiddenAttribute(DOCUMENT_CONTENT_CONTROL_DEFAULTS.lock),
      multiLine: hiddenAttribute(false),
      appearance: hiddenAttribute(DOCUMENT_CONTENT_CONTROL_DEFAULTS.appearance),
      color: hiddenAttribute(null),
    };
  },

  parseHTML() {
    return [
      {
        tag: 'span[data-document-content-control]',
        getAttrs: (element) => {
          if (!(element instanceof HTMLElement)) return false;
          return contentControlAttributesFromElement(element);
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const properties = normalizeDocumentContentControlProperties(node.attrs);
    const label = properties.alias || properties.tag || '内容控件';
    return [
      'span',
      mergeAttributes(HTMLAttributes, contentControlDomAttributes(properties), {
        class: 'work-document-content-control',
        role: 'textbox',
        'aria-label': label,
        contenteditable: contentControlLocksContent(properties.lock)
          ? 'false'
          : undefined,
      }),
      0,
    ];
  },

  renderText({ node }) {
    return node.textContent;
  },

  addCommands() {
    return {
      insertDocumentContentControl:
        (options = {}) =>
        (props: CommandProps) =>
          insertDocumentContentControlCommand(props, options),
      setDocumentContentControlProperties:
        (value, options = {}) =>
        (props: CommandProps) =>
          setDocumentContentControlPropertiesCommand(props, value, options),
      deleteDocumentContentControl:
        (options = {}) =>
        (props: CommandProps) =>
          deleteDocumentContentControlCommand(props, options),
    };
  },

  addProseMirrorPlugins() {
    return [
      new Plugin({
        filterTransaction: (transaction, state) =>
          filterLockedContentControlTransaction(transaction, state),
        appendTransaction: (transactions, _oldState, state) => {
          if (!transactions.some((transaction) => transaction.docChanged)) {
            return null;
          }
          const seen = new Set<string>();
          const updates: Array<{
            position: number;
            attrs: Record<string, unknown>;
          }> = [];
          state.doc.descendants((node, position) => {
            if (node.type.name !== this.name) return;
            const properties = normalizeDocumentContentControlProperties(
              node.attrs,
            );
            let id = properties.id;
            if (!id || seen.has(id)) id = createWorkId('content-control');
            seen.add(id);
            const attrs = {
              ...properties,
              id,
            } satisfies WorkDocumentContentControlProperties;
            if (!sameContentControlAttributes(node.attrs, attrs)) {
              updates.push({ position, attrs });
            }
          });
          if (!updates.length) return null;
          const transaction = state.tr;
          for (const update of updates) {
            const node = state.doc.nodeAt(update.position);
            if (!node || node.type.name !== this.name) continue;
            transaction.setNodeMarkup(update.position, undefined, update.attrs);
          }
          transaction.setMeta('addToHistory', false);
          transaction.setMeta(DOCUMENT_CONTENT_CONTROL_MUTATION_META, true);
          return transaction;
        },
      }),
    ];
  },
});

export function documentContentControlProperties(
  editor: Editor,
): WorkDocumentContentControlProperties {
  return normalizeDocumentContentControlProperties(
    editor.getAttributes('documentContentControl') as Record<string, unknown>,
  );
}

export function setDocumentContentControlProperties(
  editor: Editor,
  value: Partial<WorkDocumentContentControlProperties>,
  options: DocumentContentControlCommandOptions = {},
): boolean {
  return editor.commands.setDocumentContentControlProperties(value, options);
}

export function selectedDocumentContentControl(
  stateOrEditor: Editor | Editor['state'],
): SelectedDocumentContentControl | null {
  const state = 'state' in stateOrEditor ? stateOrEditor.state : stateOrEditor;
  const { $from } = state.selection;
  for (let depth = $from.depth; depth > 0; depth -= 1) {
    const node = $from.node(depth);
    if (node.type.name === 'documentContentControl') {
      return { node, position: $from.before(depth) };
    }
  }
  if (
    state.selection instanceof NodeSelection &&
    state.selection.node.type.name === 'documentContentControl'
  ) {
    return { node: state.selection.node, position: state.selection.from };
  }
  return null;
}

export function normalizeDocumentContentControlProperties(
  value: Partial<Record<keyof WorkDocumentContentControlProperties, unknown>>,
): WorkDocumentContentControlProperties {
  return {
    id: normalizeContentControlId(value.id),
    nativeId: nullableContentControlInteger(value.nativeId),
    type: value.type === 'richText' ? 'richText' : 'text',
    alias: normalizeContentControlString(value.alias),
    tag: normalizeContentControlString(value.tag),
    lock: contentControlLock(value.lock),
    multiLine: booleanValue(value.multiLine),
    appearance: contentControlAppearance(value.appearance),
    color: normalizeContentControlColor(value.color),
  };
}

export function contentControlLocksContent(
  value: WorkDocumentContentControlLock | unknown,
): boolean {
  return value === 'contentLocked' || value === 'sdtContentLocked';
}

export function contentControlLocksShell(
  value: WorkDocumentContentControlLock | unknown,
): boolean {
  return value === 'sdtLocked' || value === 'sdtContentLocked';
}

export function contentControlCss(
  value: Partial<Record<keyof WorkDocumentContentControlProperties, unknown>>,
): string {
  const properties = normalizeDocumentContentControlProperties(value);
  return [
    `--work-document-content-control-color:${properties.color ?? 'var(--a3s-accent)'}`,
    `--work-document-content-control-appearance:${properties.appearance}`,
  ].join(';');
}

export function contentControlDomAttributes(
  value: Partial<Record<keyof WorkDocumentContentControlProperties, unknown>>,
): Record<string, string | undefined> {
  const properties = normalizeDocumentContentControlProperties(value);
  return {
    'data-document-content-control': 'true',
    'data-content-control-id': properties.id || undefined,
    'data-content-control-native-id':
      properties.nativeId === null ? undefined : String(properties.nativeId),
    'data-content-control-type': properties.type,
    'data-content-control-alias': properties.alias || undefined,
    'data-content-control-tag': properties.tag || undefined,
    'data-content-control-label':
      properties.alias || properties.tag || undefined,
    'data-content-control-lock': properties.lock,
    'data-content-control-multiline': properties.multiLine ? 'true' : 'false',
    'data-content-control-appearance': properties.appearance,
    'data-content-control-color': properties.color ?? undefined,
    style: contentControlCss(properties),
  };
}

export function documentContentControlPropertiesFromElement(
  element: Element,
): WorkDocumentContentControlProperties {
  return normalizeDocumentContentControlProperties({
    id: element.getAttribute('data-content-control-id'),
    nativeId: element.getAttribute('data-content-control-native-id'),
    type: element.getAttribute('data-content-control-type'),
    alias: element.getAttribute('data-content-control-alias'),
    tag: element.getAttribute('data-content-control-tag'),
    lock: element.getAttribute('data-content-control-lock'),
    multiLine: element.getAttribute('data-content-control-multiline'),
    appearance: element.getAttribute('data-content-control-appearance'),
    color: element.getAttribute('data-content-control-color'),
  });
}

function insertDocumentContentControlCommand(
  { dispatch, editor, state, tr }: CommandProps,
  options: DocumentContentControlInsertOptions,
): boolean {
  const type = editor.schema.nodes.documentContentControl;
  if (!type) return false;
  const { from, to } = state.selection;
  const $from = state.doc.resolve(from);
  const $to = state.doc.resolve(to);
  if ($from.parent !== $to.parent || !$from.parent.inlineContent) return false;
  if (selectedDocumentContentControl(state)) return false;
  const slice = state.doc.slice(from, to);
  if (
    slice.openStart ||
    slice.openEnd ||
    containsContentControl(slice.content)
  ) {
    return false;
  }
  const properties = normalizeDocumentContentControlProperties({
    ...DOCUMENT_CONTENT_CONTROL_DEFAULTS,
    ...options,
    id: options.id || createWorkId('content-control'),
  });
  const initialText = typeof options.text === 'string' ? options.text : '';
  const content = slice.size
    ? slice.content
    : initialText
      ? Fragment.from(editor.schema.text(initialText))
      : Fragment.empty;
  if (
    properties.type === 'text' &&
    !properties.multiLine &&
    content.textBetween(0, content.size, '\n').includes('\n')
  ) {
    return false;
  }
  if (!type.validContent(content)) return false;
  if (!dispatch) return true;
  const node = type.create(properties, content);
  tr.replaceWith(from, to, node);
  tr.setMeta(DOCUMENT_CONTENT_CONTROL_MUTATION_META, true);
  const cursor = from + 1 + node.content.size;
  tr.setSelection(TextSelection.near(tr.doc.resolve(cursor), -1));
  tr.scrollIntoView();
  return true;
}

function setDocumentContentControlPropertiesCommand(
  { dispatch, editor, state, tr }: CommandProps,
  value: Partial<WorkDocumentContentControlProperties>,
  options: DocumentContentControlCommandOptions,
): boolean {
  const selected = selectedDocumentContentControl(state);
  if (!selected) return false;
  const current = normalizeDocumentContentControlProperties(
    selected.node.attrs,
  );
  if (contentControlLocksShell(current.lock) && !options.allowLocked) {
    return false;
  }
  const attributes = contentControlAttributesForChanges(value, selected.node);
  if (!Object.keys(attributes).length) return false;
  if (!dispatch) return true;
  closeHistory(tr);
  tr.setNodeMarkup(selected.position, undefined, {
    ...selected.node.attrs,
    ...attributes,
  });
  tr.setMeta(DOCUMENT_CONTENT_CONTROL_MUTATION_META, true);
  tr.scrollIntoView();
  if (options.restoreFocus !== false) editor.view.focus();
  return true;
}

function deleteDocumentContentControlCommand(
  { dispatch, editor, state, tr }: CommandProps,
  options: DocumentContentControlCommandOptions,
): boolean {
  const selected = selectedDocumentContentControl(state);
  if (!selected) return false;
  const properties = normalizeDocumentContentControlProperties(
    selected.node.attrs,
  );
  if (contentControlLocksShell(properties.lock) && !options.allowLocked) {
    return false;
  }
  if (!dispatch) return true;
  closeHistory(tr);
  tr.delete(selected.position, selected.position + selected.node.nodeSize);
  tr.setMeta(DOCUMENT_CONTENT_CONTROL_MUTATION_META, true);
  tr.setSelection(
    TextSelection.near(
      tr.doc.resolve(Math.min(selected.position, tr.doc.content.size)),
      -1,
    ),
  );
  tr.scrollIntoView();
  if (options.restoreFocus !== false) editor.view.focus();
  return true;
}

function contentControlAttributesForChanges(
  value: Partial<WorkDocumentContentControlProperties>,
  current: ProseMirrorNode,
): Record<string, unknown> {
  const merged = normalizeDocumentContentControlProperties({
    ...current.attrs,
    ...value,
  });
  const attributes: Record<string, unknown> = {};
  for (const name of CONTENT_CONTROL_MARKER_ATTRIBUTES) {
    if (name in value && merged[name] !== current.attrs[name]) {
      attributes[name] = merged[name];
    }
  }
  return attributes;
}

function contentControlAttributesFromElement(
  element: HTMLElement,
): Record<string, unknown> {
  return {
    id: element.dataset.contentControlId ?? '',
    nativeId: element.dataset.contentControlNativeId,
    type: element.dataset.contentControlType,
    alias: element.dataset.contentControlAlias ?? '',
    tag: element.dataset.contentControlTag ?? '',
    lock: element.dataset.contentControlLock,
    multiLine: element.dataset.contentControlMultiline,
    appearance: element.dataset.contentControlAppearance,
    color: element.dataset.contentControlColor,
  };
}

function filterLockedContentControlTransaction(
  transaction: Transaction,
  state: Editor['state'],
): boolean {
  if (!transaction.docChanged) return true;
  if (transaction.getMeta(DOCUMENT_CONTENT_CONTROL_MUTATION_META)) return true;
  const previous = new Map<string, ProseMirrorNode[]>();
  state.doc.descendants((node) => {
    if (node.type.name !== 'documentContentControl') return;
    const properties = normalizeDocumentContentControlProperties(node.attrs);
    if (properties.id) {
      const controls = previous.get(properties.id) ?? [];
      controls.push(node);
      previous.set(properties.id, controls);
    }
  });
  if (!previous.size) return true;
  const next = new Map<string, ProseMirrorNode[]>();
  transaction.doc.descendants((node) => {
    if (node.type.name !== 'documentContentControl') return;
    const properties = normalizeDocumentContentControlProperties(node.attrs);
    if (properties.id) {
      const controls = next.get(properties.id) ?? [];
      controls.push(node);
      next.set(properties.id, controls);
    }
  });
  for (const [id, controls] of previous) {
    for (const node of controls) {
      const properties = normalizeDocumentContentControlProperties(node.attrs);
      const replacements = next.get(id) ?? [];
      if (!replacements.length) {
        if (contentControlLocksShell(properties.lock)) return false;
        continue;
      }
      // A shell lock protects the control definition itself. A content lock
      // additionally protects its inline content. Explicit typed commands
      // cross this guard through DOCUMENT_CONTENT_CONTROL_MUTATION_META.
      if (contentControlLocksShell(properties.lock)) {
        if (
          replacements.length !== 1 ||
          !sameContentControlAttributes(
            replacements[0]?.attrs ?? {},
            properties,
          )
        ) {
          return false;
        }
      }
      if (contentControlLocksContent(properties.lock)) {
        if (
          replacements.length !== 1 ||
          !sameContentControlAttributes(
            replacements[0]?.attrs ?? {},
            properties,
          ) ||
          JSON.stringify(node.content.toJSON()) !==
            JSON.stringify(replacements[0]?.content.toJSON())
        ) {
          return false;
        }
      }
    }
  }
  return true;
}

function containsContentControl(content: ProseMirrorNode['content']): boolean {
  let found = false;
  content.descendants((node) => {
    if (node.type.name === 'documentContentControl') found = true;
    return !found;
  });
  return found;
}

function sameContentControlAttributes(
  current: Record<string, unknown>,
  next: WorkDocumentContentControlProperties,
): boolean {
  const normalized = normalizeDocumentContentControlProperties(current);
  return CONTENT_CONTROL_MARKER_ATTRIBUTES.every(
    (name) => normalized[name] === next[name],
  );
}

function normalizeContentControlId(value: unknown): string {
  return typeof value === 'string'
    ? value
        .trim()
        .replace(/[\u0000-\u001f\u007f]/gu, '')
        .slice(0, CONTENT_CONTROL_ID_MAX_LENGTH)
    : '';
}

function normalizeContentControlString(value: unknown): string {
  return typeof value === 'string'
    ? value
        .trim()
        .replace(/[\u0000-\u001f\u007f]/gu, '')
        .slice(0, CONTENT_CONTROL_STRING_MAX_LENGTH)
    : '';
}

function nullableContentControlInteger(value: unknown): number | null {
  if (value === null || value === undefined || value === '') return null;
  const number = typeof value === 'number' ? value : Number(value);
  return Number.isSafeInteger(number) &&
    number >= -2_147_483_648 &&
    number <= 2_147_483_647
    ? number
    : null;
}

function normalizeContentControlColor(value: unknown): string | null {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value !== 'string') return null;
  const normalized = value.trim().toLowerCase();
  if (CONTENT_CONTROL_COLOR_PATTERN.test(normalized)) return normalized;
  if (/^[0-9a-f]{6}$/i.test(normalized)) return `#${normalized}`;
  return null;
}

function contentControlLock(value: unknown): WorkDocumentContentControlLock {
  return value === 'contentLocked' ||
    value === 'sdtLocked' ||
    value === 'sdtContentLocked'
    ? value
    : 'unlocked';
}

function contentControlAppearance(
  value: unknown,
): WorkDocumentContentControlAppearance {
  return value === 'hidden' || value === 'tags' ? value : 'boundingBox';
}

function booleanValue(value: unknown): boolean {
  return value === true || value === 'true' || value === '1' || value === 'on';
}

function hiddenAttribute(defaultValue: unknown) {
  return {
    default: defaultValue,
    rendered: false,
  };
}
