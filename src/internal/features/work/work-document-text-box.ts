import {
  type CommandProps,
  type Editor,
  mergeAttributes,
  Node,
} from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { NodeSelection, Plugin, TextSelection } from '@tiptap/pm/state';
import { activeDocumentSectionFromState } from './work-document-section-editor';
import { createWorkId } from './work-templates';

/** The intentionally small text-box surface supported by Writer. */
export type WorkDocumentTextBoxLayout = 'inline' | 'floating';
export type WorkDocumentTextBoxHorizontalReference =
  | 'column'
  | 'margin'
  | 'page';
export type WorkDocumentTextBoxVerticalReference =
  | 'paragraph'
  | 'margin'
  | 'page';
export type WorkDocumentTextBoxVerticalAlign = 'top' | 'center' | 'bottom';

export interface WorkDocumentTextBoxProperties {
  id: string;
  width: number;
  height: number;
  layout: WorkDocumentTextBoxLayout;
  horizontalOffset: number | null;
  verticalOffset: number | null;
  horizontalReference: WorkDocumentTextBoxHorizontalReference;
  verticalReference: WorkDocumentTextBoxVerticalReference;
  fill: string;
  borderColor: string;
  borderWidth: number;
  padding: number;
  verticalAlign: WorkDocumentTextBoxVerticalAlign;
  docPropertiesId: number | null;
}

export interface DocumentTextBoxCommandOptions {
  restoreFocus?: boolean;
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentTextBox: {
      insertDocumentTextBox: (
        text?: string,
        options?: Partial<WorkDocumentTextBoxProperties>,
      ) => ReturnType;
      setDocumentTextBoxProperties: (
        value: Partial<WorkDocumentTextBoxProperties>,
        options?: DocumentTextBoxCommandOptions,
      ) => ReturnType;
      deleteDocumentTextBox: (
        options?: DocumentTextBoxCommandOptions,
      ) => ReturnType;
    };
  }
}

export const DOCUMENT_TEXT_BOX_DEFAULTS: WorkDocumentTextBoxProperties = {
  id: '',
  width: 120,
  height: 45,
  layout: 'inline',
  horizontalOffset: null,
  verticalOffset: null,
  horizontalReference: 'column',
  verticalReference: 'paragraph',
  fill: '#fff2cc',
  borderColor: '#4472c4',
  borderWidth: 0.35,
  padding: 3,
  verticalAlign: 'top',
  docPropertiesId: null,
};

export const DOCUMENT_TEXT_BOX_LIMITS = {
  width: { min: 20, max: 558.7 },
  height: { min: 10, max: 558.7 },
  offset: { min: -558.7, max: 558.7 },
  borderWidth: { min: 0, max: 10 },
  padding: { min: 0, max: 25 },
} as const;

const TEXT_BOX_ID_MAX_LENGTH = 160;
const TEXT_BOX_COLOR_PATTERN = /^#[0-9a-f]{6}$/i;
const TEXT_BOX_MARKER_ATTRIBUTES = [
  'id',
  'width',
  'height',
  'layout',
  'horizontalOffset',
  'verticalOffset',
  'horizontalReference',
  'verticalReference',
  'fill',
  'borderColor',
  'borderWidth',
  'padding',
  'verticalAlign',
  'docPropertiesId',
] as const;

export const DocumentTextBox = Node.create({
  name: 'documentTextBox',
  group: 'block',
  content: 'inline*',
  defining: true,
  isolating: true,
  selectable: true,

  addAttributes() {
    return {
      id: dataAttribute(''),
      width: dataAttribute(DOCUMENT_TEXT_BOX_DEFAULTS.width),
      height: dataAttribute(DOCUMENT_TEXT_BOX_DEFAULTS.height),
      layout: dataAttribute(DOCUMENT_TEXT_BOX_DEFAULTS.layout),
      horizontalOffset: nullableDataAttribute(),
      verticalOffset: nullableDataAttribute(),
      horizontalReference: dataAttribute(
        DOCUMENT_TEXT_BOX_DEFAULTS.horizontalReference,
      ),
      verticalReference: dataAttribute(
        DOCUMENT_TEXT_BOX_DEFAULTS.verticalReference,
      ),
      fill: dataAttribute(DOCUMENT_TEXT_BOX_DEFAULTS.fill),
      borderColor: dataAttribute(DOCUMENT_TEXT_BOX_DEFAULTS.borderColor),
      borderWidth: dataAttribute(DOCUMENT_TEXT_BOX_DEFAULTS.borderWidth),
      padding: dataAttribute(DOCUMENT_TEXT_BOX_DEFAULTS.padding),
      verticalAlign: dataAttribute(DOCUMENT_TEXT_BOX_DEFAULTS.verticalAlign),
      docPropertiesId: nullableDataAttribute(),
    };
  },

  parseHTML() {
    return [
      {
        tag: 'div[data-document-text-box]',
        getAttrs: (element) => {
          if (!(element instanceof HTMLElement)) return false;
          return textBoxAttributesFromElement(element);
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const properties = normalizeDocumentTextBoxProperties(node.attrs);
    return [
      'div',
      mergeAttributes(HTMLAttributes, textBoxDomAttributes(properties), {
        class: 'work-document-text-box',
        contenteditable: undefined,
        style: textBoxCss(properties),
        role: 'textbox',
        'aria-label': '文本框',
      }),
      0,
    ];
  },

  renderText({ node }) {
    return node.textContent;
  },

  addCommands() {
    return {
      insertDocumentTextBox:
        (text = '', options = {}) =>
        ({ dispatch, editor, state, tr }: CommandProps) =>
          insertDocumentTextBoxCommand(
            { dispatch, editor, state, tr },
            text,
            options,
          ),
      setDocumentTextBoxProperties:
        (value, options = {}) =>
        ({ chain, state, tr }) => {
          if (!selectedDocumentTextBox(state)) return false;
          const attributes = textBoxAttributesForChanges(value);
          if (!Object.keys(attributes).length) return false;
          closeHistory(tr);
          let commandChain = chain();
          if (options.restoreFocus !== false)
            commandChain = commandChain.focus();
          return commandChain.updateAttributes(this.name, attributes).run();
        },
      deleteDocumentTextBox:
        (options = {}) =>
        ({ dispatch, editor, state, tr }) => {
          const selected = selectedDocumentTextBox(state);
          if (!selected) return false;
          if (!dispatch) return true;
          closeHistory(tr);
          tr.delete(
            selected.position,
            selected.position + selected.node.nodeSize,
          );
          tr.setSelection(
            TextSelection.near(
              tr.doc.resolve(Math.min(selected.position, tr.doc.content.size)),
              -1,
            ),
          );
          dispatch(tr.scrollIntoView());
          if (options.restoreFocus !== false) editor.view.focus();
          return true;
        },
    };
  },

  addKeyboardShortcuts() {
    return {
      Backspace: () => deleteEmptyDocumentTextBox(this.editor),
      Delete: () => deleteEmptyDocumentTextBox(this.editor),
    };
  },

  addProseMirrorPlugins() {
    return [
      new Plugin({
        appendTransaction: (transactions, _oldState, state) => {
          if (!transactions.some((transaction) => transaction.docChanged)) {
            return null;
          }
          const seen = new Set<string>();
          const updates: Array<{ position: number; id: string }> = [];
          state.doc.descendants((node, position) => {
            if (node.type.name !== this.name) return;
            const current = normalizeDocumentTextBoxId(node.attrs.id);
            const id =
              current && !seen.has(current)
                ? current
                : createWorkId('text-box');
            seen.add(id);
            if (id !== node.attrs.id) updates.push({ position, id });
          });
          if (!updates.length) return null;
          const transaction = state.tr;
          for (const update of updates) {
            const node = state.doc.nodeAt(update.position);
            if (!node) continue;
            transaction.setNodeMarkup(update.position, undefined, {
              ...node.attrs,
              id: update.id,
            });
          }
          transaction.setMeta('addToHistory', false);
          return transaction;
        },
      }),
    ];
  },
});

export function documentTextBoxProperties(
  editor: Editor,
): WorkDocumentTextBoxProperties {
  return normalizeDocumentTextBoxProperties(
    editor.getAttributes('documentTextBox') as Record<string, unknown>,
  );
}

export function setDocumentTextBoxProperties(
  editor: Editor,
  value: Partial<WorkDocumentTextBoxProperties>,
  options: DocumentTextBoxCommandOptions = {},
): boolean {
  return editor.commands.setDocumentTextBoxProperties(value, options);
}

export function normalizeDocumentTextBoxProperties(
  value: Partial<Record<keyof WorkDocumentTextBoxProperties, unknown>>,
): WorkDocumentTextBoxProperties {
  return {
    id: normalizeDocumentTextBoxId(value.id),
    width: boundedNumber(
      value.width,
      DOCUMENT_TEXT_BOX_DEFAULTS.width,
      DOCUMENT_TEXT_BOX_LIMITS.width.min,
      DOCUMENT_TEXT_BOX_LIMITS.width.max,
    ),
    height: boundedNumber(
      value.height,
      DOCUMENT_TEXT_BOX_DEFAULTS.height,
      DOCUMENT_TEXT_BOX_LIMITS.height.min,
      DOCUMENT_TEXT_BOX_LIMITS.height.max,
    ),
    layout: value.layout === 'floating' ? 'floating' : 'inline',
    horizontalOffset: nullableBoundedNumber(
      value.horizontalOffset,
      DOCUMENT_TEXT_BOX_LIMITS.offset.min,
      DOCUMENT_TEXT_BOX_LIMITS.offset.max,
    ),
    verticalOffset: nullableBoundedNumber(
      value.verticalOffset,
      DOCUMENT_TEXT_BOX_LIMITS.offset.min,
      DOCUMENT_TEXT_BOX_LIMITS.offset.max,
    ),
    horizontalReference: textBoxHorizontalReference(value.horizontalReference),
    verticalReference: textBoxVerticalReference(value.verticalReference),
    fill: normalizeTextBoxFill(value.fill),
    borderColor: normalizeTextBoxBorderColor(value.borderColor),
    borderWidth: boundedNumber(
      value.borderWidth,
      DOCUMENT_TEXT_BOX_DEFAULTS.borderWidth,
      DOCUMENT_TEXT_BOX_LIMITS.borderWidth.min,
      DOCUMENT_TEXT_BOX_LIMITS.borderWidth.max,
    ),
    padding: boundedNumber(
      value.padding,
      DOCUMENT_TEXT_BOX_DEFAULTS.padding,
      DOCUMENT_TEXT_BOX_LIMITS.padding.min,
      DOCUMENT_TEXT_BOX_LIMITS.padding.max,
    ),
    verticalAlign: textBoxVerticalAlign(value.verticalAlign),
    docPropertiesId: nullableInteger(value.docPropertiesId, 0, 0xffff_ffff),
  };
}

export function textBoxCss(
  value: Partial<Record<keyof WorkDocumentTextBoxProperties, unknown>>,
): string {
  const properties = normalizeDocumentTextBoxProperties(value);
  return [
    `--work-document-text-box-width:${formatNumber(properties.width)}mm`,
    `--work-document-text-box-height:${formatNumber(properties.height)}mm`,
    `--work-document-text-box-padding:${formatNumber(properties.padding)}mm`,
    `--work-document-text-box-fill:${properties.fill}`,
    `--work-document-text-box-border-color:${properties.borderColor === 'none' ? 'transparent' : properties.borderColor}`,
    `--work-document-text-box-border-width:${formatNumber(properties.borderWidth)}mm`,
    `--work-document-text-box-vertical-align:${properties.verticalAlign}`,
    ...(properties.layout === 'floating'
      ? [
          `--work-document-text-box-horizontal-offset:${formatNullableNumber(properties.horizontalOffset)}mm`,
          `--work-document-text-box-vertical-offset:${formatNullableNumber(properties.verticalOffset)}mm`,
        ]
      : []),
  ].join(';');
}

export function textBoxDomAttributes(
  value: Partial<Record<keyof WorkDocumentTextBoxProperties, unknown>>,
): Record<string, string | undefined> {
  const properties = normalizeDocumentTextBoxProperties(value);
  return {
    'data-document-text-box': 'true',
    'data-text-box-id': properties.id || undefined,
    'data-text-box-width': formatNumber(properties.width),
    'data-text-box-height': formatNumber(properties.height),
    'data-text-box-layout': properties.layout,
    'data-text-box-horizontal-offset':
      properties.horizontalOffset === null
        ? undefined
        : formatNumber(properties.horizontalOffset),
    'data-text-box-vertical-offset':
      properties.verticalOffset === null
        ? undefined
        : formatNumber(properties.verticalOffset),
    'data-text-box-horizontal-reference': properties.horizontalReference,
    'data-text-box-vertical-reference': properties.verticalReference,
    'data-text-box-fill': properties.fill,
    'data-text-box-border-color': properties.borderColor,
    'data-text-box-border-width': formatNumber(properties.borderWidth),
    'data-text-box-padding': formatNumber(properties.padding),
    'data-text-box-vertical-align': properties.verticalAlign,
    'data-text-box-doc-properties-id':
      properties.docPropertiesId === null
        ? undefined
        : String(properties.docPropertiesId),
  };
}

function insertDocumentTextBoxCommand(
  {
    dispatch,
    editor,
    state,
    tr,
  }: Pick<CommandProps, 'dispatch' | 'editor' | 'state' | 'tr'>,
  text: string,
  options: Partial<WorkDocumentTextBoxProperties>,
): boolean {
  const section = activeDocumentSectionFromState(state);
  const textBoxType = editor.schema.nodes.documentTextBox;
  const paragraphType = editor.schema.nodes.paragraph;
  if (!section || !textBoxType || !paragraphType) return false;
  const child = activeSectionChild(section, state.selection.from);
  if (!child) return false;
  if (!dispatch) return true;
  const properties = normalizeDocumentTextBoxProperties({
    ...DOCUMENT_TEXT_BOX_DEFAULTS,
    ...options,
    id: options.id || createWorkId('text-box'),
  });
  const content = text ? editor.schema.text(text) : undefined;
  const textBox = textBoxType.create(properties, content);
  const insertPosition = section.position + 1 + child.offset + child.nodeSize;
  tr.insert(insertPosition, textBox);
  const selectionPosition = insertPosition + 1;
  if (child.index === section.node.childCount - 1) {
    const paragraphPosition = insertPosition + textBox.nodeSize;
    tr.insert(paragraphPosition, paragraphType.create());
  }
  tr.setSelection(TextSelection.near(tr.doc.resolve(selectionPosition)));
  tr.scrollIntoView();
  return true;
}

function selectedDocumentTextBox(
  state: Editor['state'],
): { node: ProseMirrorNode; position: number } | null {
  const { $from } = state.selection;
  for (let depth = $from.depth; depth > 0; depth -= 1) {
    const node = $from.node(depth);
    if (node.type.name === 'documentTextBox') {
      return { node, position: $from.before(depth) };
    }
  }
  if (
    state.selection instanceof NodeSelection &&
    state.selection.node.type.name === 'documentTextBox'
  ) {
    return { node: state.selection.node, position: state.selection.from };
  }
  return null;
}

function deleteEmptyDocumentTextBox(editor: Editor): boolean {
  const selected = selectedDocumentTextBox(editor.state);
  if (!selected || selected.node.content.size > 0) return false;
  return editor.commands.deleteDocumentTextBox();
}

function textBoxAttributesForChanges(
  value: Partial<WorkDocumentTextBoxProperties>,
): Record<string, unknown> {
  const attributes: Record<string, unknown> = {};
  for (const name of TEXT_BOX_MARKER_ATTRIBUTES) {
    if (name in value) {
      attributes[name] = normalizeDocumentTextBoxProperties({
        ...DOCUMENT_TEXT_BOX_DEFAULTS,
        ...value,
      })[name];
    }
  }
  return attributes;
}

function textBoxAttributesFromElement(
  element: HTMLElement,
): Record<string, unknown> {
  return {
    id: element.dataset.textBoxId ?? '',
    width: element.dataset.textBoxWidth,
    height: element.dataset.textBoxHeight,
    layout: element.dataset.textBoxLayout,
    horizontalOffset: element.dataset.textBoxHorizontalOffset,
    verticalOffset: element.dataset.textBoxVerticalOffset,
    horizontalReference: element.dataset.textBoxHorizontalReference,
    verticalReference: element.dataset.textBoxVerticalReference,
    fill: element.dataset.textBoxFill,
    borderColor: element.dataset.textBoxBorderColor,
    borderWidth: element.dataset.textBoxBorderWidth,
    padding: element.dataset.textBoxPadding,
    verticalAlign: element.dataset.textBoxVerticalAlign,
    docPropertiesId: element.dataset.textBoxDocPropertiesId,
  };
}

export function documentTextBoxPropertiesFromElement(
  element: Element,
): WorkDocumentTextBoxProperties {
  return normalizeDocumentTextBoxProperties({
    id: element.getAttribute('data-text-box-id'),
    width: element.getAttribute('data-text-box-width'),
    height: element.getAttribute('data-text-box-height'),
    layout: element.getAttribute('data-text-box-layout'),
    horizontalOffset: element.getAttribute('data-text-box-horizontal-offset'),
    verticalOffset: element.getAttribute('data-text-box-vertical-offset'),
    horizontalReference: element.getAttribute(
      'data-text-box-horizontal-reference',
    ),
    verticalReference: element.getAttribute('data-text-box-vertical-reference'),
    fill: element.getAttribute('data-text-box-fill'),
    borderColor: element.getAttribute('data-text-box-border-color'),
    borderWidth: element.getAttribute('data-text-box-border-width'),
    padding: element.getAttribute('data-text-box-padding'),
    verticalAlign: element.getAttribute('data-text-box-vertical-align'),
    docPropertiesId: element.getAttribute('data-text-box-doc-properties-id'),
  });
}

function dataAttribute(defaultValue: unknown) {
  return {
    default: defaultValue,
    parseHTML: () => defaultValue,
    rendered: false,
  };
}

function nullableDataAttribute() {
  return {
    default: null,
    parseHTML: () => null,
    rendered: false,
  };
}

function normalizeDocumentTextBoxId(value: unknown): string {
  return typeof value === 'string'
    ? value.trim().slice(0, TEXT_BOX_ID_MAX_LENGTH)
    : '';
}

function normalizeTextBoxFill(value: unknown): string {
  if (value === 'transparent') return 'transparent';
  return normalizeTextBoxColor(value, DOCUMENT_TEXT_BOX_DEFAULTS.fill);
}

function normalizeTextBoxBorderColor(value: unknown): string {
  if (value === 'none') return 'none';
  return normalizeTextBoxColor(value, DOCUMENT_TEXT_BOX_DEFAULTS.borderColor);
}

function normalizeTextBoxColor(value: unknown, fallback: string): string {
  if (typeof value !== 'string') return fallback;
  const normalized = value.trim().toLowerCase();
  return TEXT_BOX_COLOR_PATTERN.test(normalized) ? normalized : fallback;
}

function textBoxHorizontalReference(
  value: unknown,
): WorkDocumentTextBoxHorizontalReference {
  return value === 'margin' || value === 'page' ? value : 'column';
}

function textBoxVerticalReference(
  value: unknown,
): WorkDocumentTextBoxVerticalReference {
  return value === 'margin' || value === 'page' ? value : 'paragraph';
}

function textBoxVerticalAlign(
  value: unknown,
): WorkDocumentTextBoxVerticalAlign {
  return value === 'center' || value === 'bottom' ? value : 'top';
}

function boundedNumber(
  value: unknown,
  fallback: number,
  min: number,
  max: number,
): number {
  const number = typeof value === 'number' ? value : Number(value);
  if (!Number.isFinite(number)) return fallback;
  return Number(Math.min(max, Math.max(min, number)).toFixed(2));
}

function nullableBoundedNumber(
  value: unknown,
  min: number,
  max: number,
): number | null {
  if (value === null || value === undefined || value === '') return null;
  return boundedNumber(value, 0, min, max);
}

function nullableInteger(
  value: unknown,
  min: number,
  max: number,
): number | null {
  if (value === null || value === undefined || value === '') return null;
  const number = typeof value === 'number' ? value : Number(value);
  return Number.isSafeInteger(number)
    ? Math.min(max, Math.max(min, number))
    : null;
}

function formatNumber(value: number): string {
  return Number(value.toFixed(2)).toString();
}

function formatNullableNumber(value: number | null): string {
  return value === null ? '0' : formatNumber(value);
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
