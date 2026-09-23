import { mergeAttributes, Node, type CommandProps } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import { TextSelection } from '@tiptap/pm/state';
import { createWorkId } from './work-templates';

/**
 * Bounded body-level Word content controls.
 *
 * Admits a body-level `w:sdt` whose content is one or more plain paragraphs of
 * runs and up to two nesting levels of child `text` / `richText` block
 * controls. Deeper nesting (depth ≥ 3), tables, form controls, repeating
 * sections, and live custom-XML sync remain fail-closed. Passive data-binding
 * metadata may round-trip when present.
 */
export type WorkDocumentBlockContentControlType = 'richText' | 'text';
export type WorkDocumentBlockContentControlLock =
  | 'contentLocked'
  | 'sdtContentLocked'
  | 'sdtLocked'
  | 'unlocked';
export type WorkDocumentBlockContentControlAppearance =
  | 'boundingBox'
  | 'hidden'
  | 'tags';

export interface WorkDocumentBlockContentControlProperties {
  id: string;
  nativeId: number | null;
  type: WorkDocumentBlockContentControlType;
  alias: string;
  tag: string;
  lock: WorkDocumentBlockContentControlLock;
  appearance: WorkDocumentBlockContentControlAppearance;
  color: string | null;
  bindingStoreItemId: string;
  bindingXPath: string;
  bindingPrefixMappings: string;
}

export const DOCUMENT_BLOCK_CONTENT_CONTROL_DEFAULTS: WorkDocumentBlockContentControlProperties =
  {
    id: '',
    nativeId: null,
    type: 'richText',
    alias: '',
    tag: '',
    lock: 'unlocked',
    appearance: 'boundingBox',
    color: null,
    bindingStoreItemId: '',
    bindingXPath: '',
    bindingPrefixMappings: '',
  };

/** Maximum nesting depth of child block controls inside a body-level control. */
export const DOCUMENT_BLOCK_CONTENT_CONTROL_MAX_NEST_DEPTH = 2;

const LOCK_VALUES = new Set<WorkDocumentBlockContentControlLock>([
  'contentLocked',
  'sdtContentLocked',
  'sdtLocked',
  'unlocked',
]);
const APPEARANCE_VALUES = new Set<WorkDocumentBlockContentControlAppearance>([
  'boundingBox',
  'hidden',
  'tags',
]);
const COLOR_PATTERN = /^#[0-9a-fA-F]{6}$/;
const BINDING_STORE_ITEM_PATTERN =
  /^\{[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}$/i;
const BINDING_XPATH_PATTERN = /^\/[A-Za-z0-9_.:\[\]@='"\/\-]+$/u;
const BINDING_PATH_MAX = 1024;

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentBlockContentControl: {
      insertDocumentBlockContentControl: (
        text?: string,
        options?: Partial<WorkDocumentBlockContentControlProperties>,
      ) => ReturnType;
      setDocumentBlockContentControlProperties: (
        value: Partial<WorkDocumentBlockContentControlProperties>,
      ) => ReturnType;
      deleteDocumentBlockContentControl: () => ReturnType;
    };
  }
}

export const DocumentBlockContentControl = Node.create({
  name: 'documentBlockContentControl',
  group: 'block',
  content: '(paragraph|documentBlockContentControl)+',
  defining: true,
  isolating: true,
  selectable: true,

  addAttributes() {
    return {
      id: {
        default: '',
        parseHTML: (element) =>
          element.getAttribute('data-block-content-control-id') ?? '',
        renderHTML: (attributes) => ({
          'data-block-content-control-id': attributes.id || undefined,
        }),
      },
      nativeId: {
        default: null,
        parseHTML: (element) => {
          const value = element.getAttribute(
            'data-block-content-control-native-id',
          );
          if (value === null || value === '') return null;
          const parsed = Number(value);
          return Number.isSafeInteger(parsed) ? parsed : null;
        },
        renderHTML: (attributes) =>
          attributes.nativeId === null || attributes.nativeId === undefined
            ? {}
            : {
                'data-block-content-control-native-id': String(
                  attributes.nativeId,
                ),
              },
      },
      type: {
        default: 'richText',
        parseHTML: (element) =>
          element.getAttribute('data-block-content-control-type') ?? 'richText',
        renderHTML: (attributes) => ({
          'data-block-content-control-type': attributes.type,
        }),
      },
      alias: {
        default: '',
        parseHTML: (element) =>
          element.getAttribute('data-block-content-control-alias') ?? '',
        renderHTML: (attributes) =>
          attributes.alias
            ? { 'data-block-content-control-alias': attributes.alias }
            : {},
      },
      tag: {
        default: '',
        parseHTML: (element) =>
          element.getAttribute('data-block-content-control-tag') ?? '',
        renderHTML: (attributes) =>
          attributes.tag
            ? { 'data-block-content-control-tag': attributes.tag }
            : {},
      },
      lock: {
        default: 'unlocked',
        parseHTML: (element) =>
          element.getAttribute('data-block-content-control-lock') ?? 'unlocked',
        renderHTML: (attributes) => ({
          'data-block-content-control-lock': attributes.lock,
        }),
      },
      appearance: {
        default: 'boundingBox',
        parseHTML: (element) =>
          element.getAttribute('data-block-content-control-appearance') ??
          'boundingBox',
        renderHTML: (attributes) => ({
          'data-block-content-control-appearance': attributes.appearance,
        }),
      },
      color: {
        default: null,
        parseHTML: (element) =>
          element.getAttribute('data-block-content-control-color'),
        renderHTML: (attributes) =>
          attributes.color
            ? { 'data-block-content-control-color': attributes.color }
            : {},
      },
      bindingStoreItemId: {
        default: '',
        parseHTML: (element) =>
          element.getAttribute(
            'data-block-content-control-binding-store-item-id',
          ) ?? '',
        renderHTML: (attributes) =>
          attributes.bindingStoreItemId
            ? {
                'data-block-content-control-binding-store-item-id':
                  attributes.bindingStoreItemId,
              }
            : {},
      },
      bindingXPath: {
        default: '',
        parseHTML: (element) =>
          element.getAttribute('data-block-content-control-binding-xpath') ??
          '',
        renderHTML: (attributes) =>
          attributes.bindingXPath
            ? {
                'data-block-content-control-binding-xpath':
                  attributes.bindingXPath,
              }
            : {},
      },
      bindingPrefixMappings: {
        default: '',
        parseHTML: (element) =>
          element.getAttribute(
            'data-block-content-control-binding-prefix-mappings',
          ) ?? '',
        renderHTML: (attributes) =>
          attributes.bindingPrefixMappings
            ? {
                'data-block-content-control-binding-prefix-mappings':
                  attributes.bindingPrefixMappings,
              }
            : {},
      },
    };
  },

  parseHTML() {
    return [{ tag: 'div[data-document-block-content-control]' }];
  },

  renderHTML({ node, HTMLAttributes }) {
    const properties = normalizeBlockContentControlProperties(node.attrs);
    const label = properties.alias || properties.tag || '块内容控件';
    return [
      'div',
      mergeAttributes(
        HTMLAttributes,
        blockContentControlDomAttributes(properties),
        {
          class: 'work-document-block-content-control',
          role: 'group',
          'aria-label': label,
        },
      ),
      0,
    ];
  },

  addCommands() {
    return {
      insertDocumentBlockContentControl:
        (text = '', options = {}) =>
        (props: CommandProps) =>
          insertDocumentBlockContentControlCommand(props, text, options),
      setDocumentBlockContentControlProperties:
        (value) =>
        ({ dispatch, state, tr }) => {
          const selected = selectedBlockContentControl(state);
          if (!selected) return false;
          const next = normalizeBlockContentControlProperties({
            ...selected.node.attrs,
            ...value,
          });
          if (!dispatch) return true;
          tr.setNodeMarkup(selected.position, undefined, next);
          closeHistory(tr);
          return true;
        },
      deleteDocumentBlockContentControl:
        () =>
        ({ dispatch, state, tr }) => {
          const selected = selectedBlockContentControl(state);
          if (!selected) return false;
          if (!dispatch) return true;
          tr.delete(selected.position, selected.position + selected.node.nodeSize);
          closeHistory(tr);
          return true;
        },
    };
  },
});

export function normalizeBlockContentControlProperties(
  value: Partial<WorkDocumentBlockContentControlProperties> | null | undefined,
): WorkDocumentBlockContentControlProperties {
  const lock = LOCK_VALUES.has(value?.lock as WorkDocumentBlockContentControlLock)
    ? (value?.lock as WorkDocumentBlockContentControlLock)
    : 'unlocked';
  const appearance = APPEARANCE_VALUES.has(
    value?.appearance as WorkDocumentBlockContentControlAppearance,
  )
    ? (value?.appearance as WorkDocumentBlockContentControlAppearance)
    : 'boundingBox';
  const color =
    typeof value?.color === 'string' && COLOR_PATTERN.test(value.color.trim())
      ? value.color.trim().toLowerCase()
      : null;
  const binding = normalizeBlockBinding({
    bindingStoreItemId: value?.bindingStoreItemId,
    bindingXPath: value?.bindingXPath,
    bindingPrefixMappings: value?.bindingPrefixMappings,
  });
  return {
    id: sanitize(value?.id, 160),
    nativeId: nullableInteger(value?.nativeId),
    type: value?.type === 'text' ? 'text' : 'richText',
    alias: sanitize(value?.alias, 255),
    tag: sanitize(value?.tag, 255),
    lock,
    appearance,
    color,
    ...binding,
  };
}

export function blockContentControlDomAttributes(
  properties: WorkDocumentBlockContentControlProperties,
): Record<string, string | undefined> {
  return {
    'data-document-block-content-control': 'true',
    'data-block-content-control-id': properties.id || undefined,
    'data-block-content-control-native-id':
      properties.nativeId === null ? undefined : String(properties.nativeId),
    'data-block-content-control-type': properties.type,
    'data-block-content-control-alias': properties.alias || undefined,
    'data-block-content-control-tag': properties.tag || undefined,
    'data-block-content-control-lock': properties.lock,
    'data-block-content-control-appearance': properties.appearance,
    'data-block-content-control-color': properties.color ?? undefined,
    'data-block-content-control-binding-store-item-id':
      properties.bindingStoreItemId || undefined,
    'data-block-content-control-binding-xpath':
      properties.bindingXPath || undefined,
    'data-block-content-control-binding-prefix-mappings':
      properties.bindingPrefixMappings || undefined,
  };
}

export function blockContentControlPropertiesFromElement(
  element: Element,
): WorkDocumentBlockContentControlProperties {
  return normalizeBlockContentControlProperties({
    id: element.getAttribute('data-block-content-control-id') ?? '',
    nativeId: (() => {
      const value = element.getAttribute('data-block-content-control-native-id');
      if (value === null || value === '') return null;
      const parsed = Number(value);
      return Number.isSafeInteger(parsed) ? parsed : null;
    })(),
    type: element.getAttribute(
      'data-block-content-control-type',
    ) as WorkDocumentBlockContentControlType,
    alias: element.getAttribute('data-block-content-control-alias') ?? '',
    tag: element.getAttribute('data-block-content-control-tag') ?? '',
    lock: element.getAttribute(
      'data-block-content-control-lock',
    ) as WorkDocumentBlockContentControlLock,
    appearance: element.getAttribute(
      'data-block-content-control-appearance',
    ) as WorkDocumentBlockContentControlAppearance,
    color: element.getAttribute('data-block-content-control-color'),
    bindingStoreItemId:
      element.getAttribute('data-block-content-control-binding-store-item-id') ??
      '',
    bindingXPath:
      element.getAttribute('data-block-content-control-binding-xpath') ?? '',
    bindingPrefixMappings:
      element.getAttribute(
        'data-block-content-control-binding-prefix-mappings',
      ) ?? '',
  });
}

function insertDocumentBlockContentControlCommand(
  { dispatch, editor, state, tr }: CommandProps,
  text: string,
  options: Partial<WorkDocumentBlockContentControlProperties>,
): boolean {
  const type = state.schema.nodes.documentBlockContentControl;
  const paragraph = state.schema.nodes.paragraph;
  if (!type || !paragraph) return false;
  if (
    blockContentControlNestingDepth(state) >=
    DOCUMENT_BLOCK_CONTENT_CONTROL_MAX_NEST_DEPTH + 1
  ) {
    return false;
  }
  const properties = normalizeBlockContentControlProperties({
    ...DOCUMENT_BLOCK_CONTENT_CONTROL_DEFAULTS,
    ...options,
    id: options.id?.trim() || createWorkId('block-content-control'),
  });
  const lines = text.length
    ? text
        .slice(0, 10_000)
        .split(/\r\n|\r|\n/u)
        .slice(0, 64)
    : [''];
  const content = lines.map((line) =>
    paragraph.create(null, line ? state.schema.text(line) : undefined),
  );
  const node = type.create(properties, content);
  if (!dispatch) return true;
  closeHistory(tr);
  const { from, to } = state.selection;
  tr.replaceWith(from, to, node);
  const resolved = tr.doc.resolve(Math.min(from + 2, tr.doc.content.size));
  tr.setSelection(TextSelection.near(resolved));
  dispatch(tr.scrollIntoView());
  editor.view.focus();
  return true;
}

function blockContentControlNestingDepth(
  state: CommandProps['state'],
): number {
  const { $from } = state.selection;
  let depth = 0;
  for (let level = $from.depth; level > 0; level -= 1) {
    if ($from.node(level).type.name === 'documentBlockContentControl') {
      depth += 1;
    }
  }
  return depth;
}

function selectedBlockContentControl(state: CommandProps['state']): {
  node: import('@tiptap/pm/model').Node;
  position: number;
} | null {
  const { $from } = state.selection;
  for (let depth = $from.depth; depth > 0; depth -= 1) {
    const node = $from.node(depth);
    if (node.type.name === 'documentBlockContentControl') {
      return { node, position: $from.before(depth) };
    }
  }
  return null;
}

function normalizeBlockBinding(value: {
  bindingStoreItemId: unknown;
  bindingXPath: unknown;
  bindingPrefixMappings: unknown;
}): Pick<
  WorkDocumentBlockContentControlProperties,
  'bindingStoreItemId' | 'bindingXPath' | 'bindingPrefixMappings'
> {
  const empty = {
    bindingStoreItemId: '',
    bindingXPath: '',
    bindingPrefixMappings: '',
  };
  const storeItemId = sanitize(value.bindingStoreItemId, 64);
  const xpath = sanitize(value.bindingXPath, BINDING_PATH_MAX);
  const prefixMappings = sanitize(value.bindingPrefixMappings, BINDING_PATH_MAX);
  if (!storeItemId && !xpath && !prefixMappings) return empty;
  if (
    !BINDING_STORE_ITEM_PATTERN.test(storeItemId) ||
    !BINDING_XPATH_PATTERN.test(xpath)
  ) {
    return empty;
  }
  return {
    bindingStoreItemId: storeItemId,
    bindingXPath: xpath,
    bindingPrefixMappings: prefixMappings,
  };
}

function sanitize(value: unknown, max: number): string {
  return typeof value === 'string'
    ? value
        .trim()
        .replace(/[\u0000-\u001f\u007f]/gu, '')
        .slice(0, max)
    : '';
}

function nullableInteger(value: unknown): number | null {
  if (value === null || value === undefined || value === '') return null;
  const number = typeof value === 'number' ? value : Number(value);
  return Number.isSafeInteger(number) &&
    number >= -2_147_483_648 &&
    number <= 2_147_483_647
    ? number
    : null;
}
