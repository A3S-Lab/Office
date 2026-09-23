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
 * Repeating sections use a dedicated block model. Picture controls and other
 * active form semantics beyond the bounded checkbox, drop-down list, combo
 * box, date, and passive custom-XML data-binding slices are intentionally not
 * represented here. A control is an inline, editable container whose metadata
 * can round-trip to a passive Word `w:sdt`.
 */
export type WorkDocumentContentControlType =
  | 'text'
  | 'richText'
  | 'checkbox'
  | 'dropDownList'
  | 'comboBox'
  | 'date';
export type WorkDocumentContentControlLock =
  | 'unlocked'
  | 'contentLocked'
  | 'sdtLocked'
  | 'sdtContentLocked';
export type WorkDocumentContentControlAppearance =
  | 'boundingBox'
  | 'hidden'
  | 'tags';

export type WorkDocumentContentControlDateMapping =
  | 'date'
  | 'dateTime'
  | 'text';

export interface WorkDocumentContentControlListItem {
  displayText: string;
  value: string;
}

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
  /** Checked state for `checkbox` controls; ignored otherwise. */
  checked: boolean;
  /** Ordered list items for `dropDownList` and `comboBox` controls. */
  options: WorkDocumentContentControlListItem[];
  /**
   * Selected option value for `dropDownList` and `comboBox` controls.
   * Combo boxes may leave this empty when the visible text is free-form.
   */
  selectedValue: string;
  /** Last known date for `date` controls (`xsd:dateTime`, UTC). */
  fullDate: string;
  /** Display mask for `date` controls (Word `w:dateFormat`). */
  dateFormat: string;
  /** Language id for `date` controls (Word `w:lid`). */
  dateLanguage: string;
  /** Custom-XML mapping hint for `date` controls (Word `w:storeMappedDataAs`). */
  dateMapping: WorkDocumentContentControlDateMapping;
  /**
   * Passive `w:dataBinding` store item GUID when present. Empty means unbound.
   * Live custom-XML sync stays host-owned; Writer only round-trips metadata.
   */
  bindingStoreItemId: string;
  /** Absolute XPath for `w:dataBinding`; required when a store item id is set. */
  bindingXPath: string;
  /** Optional `w:prefixMappings` for the binding XPath namespaces. */
  bindingPrefixMappings: string;
}

/** Ballot-box glyphs used for bounded checkbox display and DOCX export. */
export const DOCUMENT_CONTENT_CONTROL_CHECKBOX_UNCHECKED = '\u2610';
export const DOCUMENT_CONTENT_CONTROL_CHECKBOX_CHECKED = '\u2611';
export const DOCUMENT_CONTENT_CONTROL_MAX_LIST_ITEMS = 32;
export const DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_FORMAT = 'yyyy年M月d日';
export const DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_LANGUAGE = 'zh-CN';
export const DOCUMENT_CONTENT_CONTROL_DATE_FORMATS = [
  'yyyy年M月d日',
  'yyyy-MM-dd',
  'MMMM d, yyyy',
  'dddd, MMMM d, yyyy',
] as const;

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
    checked: false,
    options: [],
    selectedValue: '',
    fullDate: '',
    dateFormat: DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_FORMAT,
    dateLanguage: DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_LANGUAGE,
    dateMapping: 'dateTime',
    bindingStoreItemId: '',
    bindingXPath: '',
    bindingPrefixMappings: '',
  };

/** Transaction meta used by the typed commands to cross the lock guard. */
export const DOCUMENT_CONTENT_CONTROL_MUTATION_META =
  'documentContentControlMutation';

const CONTENT_CONTROL_ID_MAX_LENGTH = 160;
const CONTENT_CONTROL_STRING_MAX_LENGTH = 255;
const CONTENT_CONTROL_BINDING_PATH_MAX_LENGTH = 1024;
const CONTENT_CONTROL_BINDING_STORE_ITEM_PATTERN =
  /^\{[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}$/i;
const CONTENT_CONTROL_BINDING_XPATH_PATTERN =
  /^\/[A-Za-z0-9_.:\[\]@='"\/\-]+$/u;
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
  'checked',
  'options',
  'selectedValue',
  'fullDate',
  'dateFormat',
  'dateLanguage',
  'dateMapping',
  'bindingStoreItemId',
  'bindingXPath',
  'bindingPrefixMappings',
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
      checked: hiddenAttribute(false),
      options: hiddenAttribute([]),
      selectedValue: hiddenAttribute(''),
      fullDate: hiddenAttribute(''),
      dateFormat: hiddenAttribute(
        DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_FORMAT,
      ),
      dateLanguage: hiddenAttribute(
        DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_LANGUAGE,
      ),
      dateMapping: hiddenAttribute('dateTime'),
      bindingStoreItemId: hiddenAttribute(''),
      bindingXPath: hiddenAttribute(''),
      bindingPrefixMappings: hiddenAttribute(''),
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
    const isCheckbox = properties.type === 'checkbox';
    const isDropDown = properties.type === 'dropDownList';
    const isComboBox = properties.type === 'comboBox';
    const isDate = properties.type === 'date';
    return [
      'span',
      mergeAttributes(HTMLAttributes, contentControlDomAttributes(properties), {
        class: 'work-document-content-control',
        role: isCheckbox
          ? 'checkbox'
          : isDropDown
            ? 'listbox'
            : isComboBox
              ? 'combobox'
              : isDate
                ? 'textbox'
                : 'textbox',
        'aria-label': label,
        'aria-checked': isCheckbox
          ? properties.checked
            ? 'true'
            : 'false'
          : undefined,
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
  const type = contentControlType(value.type);
  const checked = type === 'checkbox' ? booleanValue(value.checked) : false;
  const isList = isListStyleContentControl(type);
  const isDate = type === 'date';
  const options = isList ? normalizeContentControlListItems(value.options) : [];
  const selectedValue = isList
    ? normalizeContentControlSelectedValue(
        value.selectedValue,
        options,
        type === 'comboBox',
      )
    : '';
  const fullDate = isDate ? normalizeContentControlFullDate(value.fullDate) : '';
  const dateFormat = isDate
    ? normalizeContentControlDateFormat(value.dateFormat)
    : DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_FORMAT;
  const dateLanguage = isDate
    ? normalizeContentControlDateLanguage(value.dateLanguage)
    : DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_LANGUAGE;
  const dateMapping = isDate
    ? normalizeContentControlDateMapping(value.dateMapping)
    : 'dateTime';
  const binding = normalizeContentControlDataBinding({
    bindingStoreItemId: value.bindingStoreItemId,
    bindingXPath: value.bindingXPath,
    bindingPrefixMappings: value.bindingPrefixMappings,
  });
  return {
    id: normalizeContentControlId(value.id),
    nativeId: nullableContentControlInteger(value.nativeId),
    type,
    alias: normalizeContentControlString(value.alias),
    tag: normalizeContentControlString(value.tag),
    lock: contentControlLock(value.lock),
    multiLine:
      isList || type === 'checkbox' || isDate
        ? false
        : booleanValue(value.multiLine),
    appearance: contentControlAppearance(value.appearance),
    color: normalizeContentControlColor(value.color),
    checked,
    options,
    selectedValue,
    fullDate,
    dateFormat,
    dateLanguage,
    dateMapping,
    bindingStoreItemId: binding.bindingStoreItemId,
    bindingXPath: binding.bindingXPath,
    bindingPrefixMappings: binding.bindingPrefixMappings,
  };
}

export function isListStyleContentControl(
  type: WorkDocumentContentControlType | unknown,
): type is 'comboBox' | 'dropDownList' {
  return type === 'dropDownList' || type === 'comboBox';
}

export function isStructuredFormContentControl(
  type: WorkDocumentContentControlType | unknown,
): type is 'checkbox' | 'comboBox' | 'date' | 'dropDownList' {
  return (
    type === 'checkbox' ||
    type === 'dropDownList' ||
    type === 'comboBox' ||
    type === 'date'
  );
}

export function documentContentControlCheckboxGlyph(checked: boolean): string {
  return checked
    ? DOCUMENT_CONTENT_CONTROL_CHECKBOX_CHECKED
    : DOCUMENT_CONTENT_CONTROL_CHECKBOX_UNCHECKED;
}

export function documentContentControlDropDownDisplay(
  properties: Pick<
    WorkDocumentContentControlProperties,
    'options' | 'selectedValue'
  >,
): string {
  const selected = properties.options.find(
    (item) => item.value === properties.selectedValue,
  );
  return selected?.displayText ?? properties.options[0]?.displayText ?? '';
}

export function documentContentControlDateDisplay(
  properties: Pick<
    WorkDocumentContentControlProperties,
    'fullDate' | 'dateFormat'
  >,
): string {
  const date = parseContentControlFullDate(properties.fullDate);
  if (!date) return '';
  return formatContentControlDate(date, properties.dateFormat);
}

export function contentControlDateInputValue(fullDate: string): string {
  const date = parseContentControlFullDate(fullDate);
  if (!date) return '';
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const day = String(date.getUTCDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

export function contentControlFullDateFromInput(value: string): string {
  return normalizeContentControlFullDate(value) || defaultContentControlFullDate();
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
    'data-content-control-checked':
      properties.type === 'checkbox'
        ? properties.checked
          ? 'true'
          : 'false'
        : undefined,
    'data-content-control-options': isListStyleContentControl(properties.type)
      ? JSON.stringify(properties.options)
      : undefined,
    'data-content-control-selected-value': isListStyleContentControl(
      properties.type,
    )
      ? properties.selectedValue || undefined
      : undefined,
    'data-content-control-full-date':
      properties.type === 'date' ? properties.fullDate || undefined : undefined,
    'data-content-control-date-format':
      properties.type === 'date' ? properties.dateFormat : undefined,
    'data-content-control-date-language':
      properties.type === 'date' ? properties.dateLanguage : undefined,
    'data-content-control-date-mapping':
      properties.type === 'date' ? properties.dateMapping : undefined,
    'data-content-control-binding-store-item-id':
      properties.bindingStoreItemId || undefined,
    'data-content-control-binding-xpath':
      properties.bindingXPath || undefined,
    'data-content-control-binding-prefix-mappings':
      properties.bindingPrefixMappings || undefined,
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
    checked: element.getAttribute('data-content-control-checked'),
    options: element.getAttribute('data-content-control-options'),
    selectedValue: element.getAttribute('data-content-control-selected-value'),
    fullDate: element.getAttribute('data-content-control-full-date'),
    dateFormat: element.getAttribute('data-content-control-date-format'),
    dateLanguage: element.getAttribute('data-content-control-date-language'),
    dateMapping: element.getAttribute('data-content-control-date-mapping'),
    bindingStoreItemId: element.getAttribute(
      'data-content-control-binding-store-item-id',
    ),
    bindingXPath: element.getAttribute('data-content-control-binding-xpath'),
    bindingPrefixMappings: element.getAttribute(
      'data-content-control-binding-prefix-mappings',
    ),
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
  let properties = normalizeDocumentContentControlProperties({
    ...DOCUMENT_CONTENT_CONTROL_DEFAULTS,
    ...options,
    id: options.id || createWorkId('content-control'),
  });
  if (properties.type === 'date' && !properties.fullDate) {
    properties = {
      ...properties,
      fullDate: defaultContentControlFullDate(),
    };
  }
  if (
    isStructuredFormContentControl(properties.type) &&
    slice.size
  ) {
    return false;
  }
  if (
    isListStyleContentControl(properties.type) &&
    !properties.options.length
  ) {
    return false;
  }
  if (properties.type === 'date' && !properties.fullDate) {
    return false;
  }
  const initialText =
    typeof options.text === 'string'
      ? options.text
      : properties.type === 'checkbox'
        ? documentContentControlCheckboxGlyph(properties.checked)
        : isListStyleContentControl(properties.type)
          ? documentContentControlDropDownDisplay(properties)
          : properties.type === 'date'
            ? documentContentControlDateDisplay(properties)
            : '';
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
  if (
    isStructuredFormContentControl(properties.type) &&
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
  const next = normalizeDocumentContentControlProperties({
    ...selected.node.attrs,
    ...value,
  });
  if (isListStyleContentControl(next.type) && !next.options.length) {
    return false;
  }
  const shouldSyncCheckboxGlyph =
    next.type === 'checkbox' &&
    (current.type !== 'checkbox' ||
      next.checked !== current.checked ||
      ('type' in value && value.type === 'checkbox'));
  const shouldSyncListDisplay =
    isListStyleContentControl(next.type) &&
    (next.type === 'dropDownList' || Boolean(next.selectedValue)) &&
    (current.type !== next.type ||
      next.selectedValue !== current.selectedValue ||
      JSON.stringify(next.options) !== JSON.stringify(current.options) ||
      ('type' in value && isListStyleContentControl(value.type)));
  const shouldSyncDateDisplay =
    next.type === 'date' &&
    (current.type !== 'date' ||
      next.fullDate !== current.fullDate ||
      next.dateFormat !== current.dateFormat ||
      ('type' in value && value.type === 'date'));
  if (
    !Object.keys(attributes).length &&
    !shouldSyncCheckboxGlyph &&
    !shouldSyncListDisplay &&
    !shouldSyncDateDisplay
  ) {
    return false;
  }
  if (!dispatch) return true;
  closeHistory(tr);
  const nextAttrs = {
    ...selected.node.attrs,
    ...attributes,
  };
  if (shouldSyncCheckboxGlyph) {
    const glyph = documentContentControlCheckboxGlyph(next.checked);
    const content = Fragment.from(editor.schema.text(glyph));
    const replacement = selected.node.type.create(nextAttrs, content);
    tr.replaceWith(
      selected.position,
      selected.position + selected.node.nodeSize,
      replacement,
    );
  } else if (shouldSyncListDisplay) {
    const display = documentContentControlDropDownDisplay(next);
    const content = display
      ? Fragment.from(editor.schema.text(display))
      : Fragment.empty;
    const replacement = selected.node.type.create(nextAttrs, content);
    tr.replaceWith(
      selected.position,
      selected.position + selected.node.nodeSize,
      replacement,
    );
  } else if (shouldSyncDateDisplay) {
    const display = documentContentControlDateDisplay(next);
    const content = display
      ? Fragment.from(editor.schema.text(display))
      : Fragment.empty;
    const replacement = selected.node.type.create(nextAttrs, content);
    tr.replaceWith(
      selected.position,
      selected.position + selected.node.nodeSize,
      replacement,
    );
  } else {
    tr.setNodeMarkup(selected.position, undefined, nextAttrs);
  }
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
  if (!Object.keys(value).length) return {};
  const merged = normalizeDocumentContentControlProperties({
    ...current.attrs,
    ...value,
  });
  const attributes: Record<string, unknown> = {};
  for (const name of CONTENT_CONTROL_MARKER_ATTRIBUTES) {
    if (name === 'options') {
      if (
        JSON.stringify(merged.options) !==
        JSON.stringify(
          normalizeContentControlListItems(current.attrs.options),
        )
      ) {
        attributes.options = merged.options;
      }
      continue;
    }
    if (merged[name] !== current.attrs[name]) {
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
    checked: element.dataset.contentControlChecked,
    options: element.dataset.contentControlOptions,
    selectedValue: element.dataset.contentControlSelectedValue,
    fullDate: element.dataset.contentControlFullDate,
    dateFormat: element.dataset.contentControlDateFormat,
    dateLanguage: element.dataset.contentControlDateLanguage,
    dateMapping: element.dataset.contentControlDateMapping,
    bindingStoreItemId: element.dataset.contentControlBindingStoreItemId,
    bindingXPath: element.dataset.contentControlBindingXpath,
    bindingPrefixMappings: element.dataset.contentControlBindingPrefixMappings,
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
  return CONTENT_CONTROL_MARKER_ATTRIBUTES.every((name) => {
    if (name === 'options') {
      return (
        JSON.stringify(normalized.options) === JSON.stringify(next.options)
      );
    }
    return normalized[name] === next[name];
  });
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

function normalizeContentControlDataBinding(value: {
  bindingStoreItemId: unknown;
  bindingXPath: unknown;
  bindingPrefixMappings: unknown;
}): Pick<
  WorkDocumentContentControlProperties,
  'bindingStoreItemId' | 'bindingXPath' | 'bindingPrefixMappings'
> {
  const empty = {
    bindingStoreItemId: '',
    bindingXPath: '',
    bindingPrefixMappings: '',
  };
  const storeItemId =
    typeof value.bindingStoreItemId === 'string'
      ? value.bindingStoreItemId
          .trim()
          .replace(/[\u0000-\u001f\u007f]/gu, '')
          .slice(0, 64)
      : '';
  const xpath =
    typeof value.bindingXPath === 'string'
      ? value.bindingXPath
          .trim()
          .replace(/[\u0000-\u001f\u007f]/gu, '')
          .slice(0, CONTENT_CONTROL_BINDING_PATH_MAX_LENGTH)
      : '';
  const prefixMappings =
    typeof value.bindingPrefixMappings === 'string'
      ? value.bindingPrefixMappings
          .trim()
          .replace(/[\u0000-\u001f\u007f]/gu, '')
          .slice(0, CONTENT_CONTROL_BINDING_PATH_MAX_LENGTH)
      : '';
  if (!storeItemId && !xpath && !prefixMappings) return empty;
  if (
    !CONTENT_CONTROL_BINDING_STORE_ITEM_PATTERN.test(storeItemId) ||
    !CONTENT_CONTROL_BINDING_XPATH_PATTERN.test(xpath)
  ) {
    return empty;
  }
  return {
    bindingStoreItemId: storeItemId,
    bindingXPath: xpath,
    bindingPrefixMappings: prefixMappings,
  };
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

function contentControlType(value: unknown): WorkDocumentContentControlType {
  return value === 'richText' ||
    value === 'checkbox' ||
    value === 'dropDownList' ||
    value === 'comboBox' ||
    value === 'date'
    ? value
    : 'text';
}

function normalizeContentControlListItems(
  value: unknown,
): WorkDocumentContentControlListItem[] {
  let raw: unknown = value;
  if (typeof value === 'string') {
    try {
      raw = JSON.parse(value);
    } catch {
      return [];
    }
  }
  if (!Array.isArray(raw)) return [];
  const items: WorkDocumentContentControlListItem[] = [];
  const seen = new Set<string>();
  for (const entry of raw) {
    if (items.length >= DOCUMENT_CONTENT_CONTROL_MAX_LIST_ITEMS) break;
    if (!entry || typeof entry !== 'object') continue;
    const record = entry as Record<string, unknown>;
    const displayText = normalizeContentControlString(record.displayText);
    const itemValue =
      normalizeContentControlString(record.value) || displayText;
    if (!displayText || !itemValue || seen.has(itemValue)) continue;
    seen.add(itemValue);
    items.push({ displayText, value: itemValue });
  }
  return items;
}

function normalizeContentControlSelectedValue(
  value: unknown,
  options: WorkDocumentContentControlListItem[],
  allowEmpty = false,
): string {
  if (!options.length) return '';
  const selected = normalizeContentControlString(value);
  if (selected && options.some((item) => item.value === selected)) {
    return selected;
  }
  return allowEmpty ? '' : (options[0]?.value ?? '');
}

function normalizeContentControlFullDate(value: unknown): string {
  if (typeof value !== 'string') return '';
  const trimmed = value.trim();
  if (!trimmed) return '';
  const match = trimmed.match(
    /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2}):(\d{2})(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?)?$/u,
  );
  if (!match) return '';
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = Number(match[4] ?? '0');
  const minute = Number(match[5] ?? '0');
  const second = Number(match[6] ?? '0');
  if (
    !Number.isSafeInteger(year) ||
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > 31 ||
    hour > 23 ||
    minute > 59 ||
    second > 59
  ) {
    return '';
  }
  const date = new Date(Date.UTC(year, month - 1, day, hour, minute, second));
  if (
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() + 1 !== month ||
    date.getUTCDate() !== day
  ) {
    return '';
  }
  return `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}T${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}:${String(second).padStart(2, '0')}Z`;
}

function defaultContentControlFullDate(): string {
  const now = new Date();
  return normalizeContentControlFullDate(
    `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`,
  );
}

function normalizeContentControlDateFormat(value: unknown): string {
  const format = normalizeContentControlString(value);
  return format || DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_FORMAT;
}

function normalizeContentControlDateLanguage(value: unknown): string {
  if (typeof value !== 'string') {
    return DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_LANGUAGE;
  }
  const language = value.trim();
  if (!/^[A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,8})*$/u.test(language)) {
    return DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_LANGUAGE;
  }
  return language.slice(0, 32);
}

function normalizeContentControlDateMapping(
  value: unknown,
): WorkDocumentContentControlDateMapping {
  return value === 'date' || value === 'text' ? value : 'dateTime';
}

function parseContentControlFullDate(value: string): Date | null {
  const normalized = normalizeContentControlFullDate(value);
  if (!normalized) return null;
  const date = new Date(normalized);
  return Number.isFinite(date.getTime()) ? date : null;
}

function formatContentControlDate(date: Date, format: string): string {
  const hour12 = date.getUTCHours() % 12 || 12;
  const replacements: Record<string, string> = {
    'AM/PM': date.getUTCHours() < 12 ? 'AM' : 'PM',
    'am/pm': date.getUTCHours() < 12 ? 'am' : 'pm',
    yyyy: String(date.getUTCFullYear()).padStart(4, '0'),
    yy: String(date.getUTCFullYear() % 100).padStart(2, '0'),
    MMMM: new Intl.DateTimeFormat('zh-CN', {
      month: 'long',
      timeZone: 'UTC',
    }).format(date),
    MMM: new Intl.DateTimeFormat('zh-CN', {
      month: 'short',
      timeZone: 'UTC',
    }).format(date),
    MM: String(date.getUTCMonth() + 1).padStart(2, '0'),
    M: String(date.getUTCMonth() + 1),
    dddd: new Intl.DateTimeFormat('zh-CN', {
      weekday: 'long',
      timeZone: 'UTC',
    }).format(date),
    ddd: new Intl.DateTimeFormat('zh-CN', {
      weekday: 'short',
      timeZone: 'UTC',
    }).format(date),
    dd: String(date.getUTCDate()).padStart(2, '0'),
    d: String(date.getUTCDate()),
    HH: String(date.getUTCHours()).padStart(2, '0'),
    H: String(date.getUTCHours()),
    hh: String(hour12).padStart(2, '0'),
    h: String(hour12),
    mm: String(date.getUTCMinutes()).padStart(2, '0'),
    m: String(date.getUTCMinutes()),
    ss: String(date.getUTCSeconds()).padStart(2, '0'),
    s: String(date.getUTCSeconds()),
  };
  return format.replace(
    /AM\/PM|am\/pm|yyyy|MMMM|dddd|MMM|ddd|yy|MM|dd|HH|hh|mm|ss|M|d|H|h|m|s/g,
    (token) => replacements[token] ?? token,
  );
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
