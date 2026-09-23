import { mergeAttributes, Node } from '@tiptap/core';
import type { CommandProps } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import { TextSelection } from '@tiptap/pm/state';
import { createWorkId } from './work-templates';

/**
 * Bounded Word repeating-section content controls.
 *
 * Admits a body-level `w15:repeatingSection` wrapping one or more
 * `w15:repeatingSectionItem` children whose bodies are a single plain
 * paragraph of runs. Add/remove-item UI, data binding, nested repeating
 * sections, multi-paragraph items, and nested form controls remain
 * fail-closed.
 */
export interface WorkDocumentRepeatingSectionProperties {
  id: string;
  nativeId: number | null;
  alias: string;
  tag: string;
  lock: WorkDocumentRepeatingSectionLock;
  appearance: WorkDocumentRepeatingSectionAppearance;
  color: string | null;
}

export interface WorkDocumentRepeatingSectionItemProperties {
  id: string;
  nativeId: number | null;
}

export type WorkDocumentRepeatingSectionLock =
  | 'contentLocked'
  | 'sdtContentLocked'
  | 'sdtLocked'
  | 'unlocked';

export type WorkDocumentRepeatingSectionAppearance =
  | 'boundingBox'
  | 'hidden'
  | 'tags';

export const DOCUMENT_REPEATING_SECTION_DEFAULTS: WorkDocumentRepeatingSectionProperties =
  {
    id: '',
    nativeId: null,
    alias: '',
    tag: '',
    lock: 'unlocked',
    appearance: 'boundingBox',
    color: null,
  };

const LOCK_VALUES = new Set<WorkDocumentRepeatingSectionLock>([
  'contentLocked',
  'sdtContentLocked',
  'sdtLocked',
  'unlocked',
]);

const APPEARANCE_VALUES = new Set<WorkDocumentRepeatingSectionAppearance>([
  'boundingBox',
  'hidden',
  'tags',
]);

const COLOR_PATTERN = /^#[0-9a-fA-F]{6}$/;

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentRepeatingSection: {
      insertDocumentRepeatingSection: (
        text?: string,
        options?: Partial<WorkDocumentRepeatingSectionProperties>,
      ) => ReturnType;
      setDocumentRepeatingSectionProperties: (
        value: Partial<WorkDocumentRepeatingSectionProperties>,
      ) => ReturnType;
      deleteDocumentRepeatingSection: () => ReturnType;
    };
  }
}

export const DocumentRepeatingSectionItem = Node.create({
  name: 'documentRepeatingSectionItem',
  group: 'block',
  content: 'paragraph',
  defining: true,
  isolating: true,
  selectable: true,

  addAttributes() {
    return {
      id: {
        default: '',
        parseHTML: (element) =>
          element.getAttribute('data-repeating-section-item-id') ?? '',
        renderHTML: (attributes) => ({
          'data-repeating-section-item-id': attributes.id || undefined,
        }),
      },
      nativeId: {
        default: null,
        parseHTML: (element) => {
          const value = element.getAttribute(
            'data-repeating-section-item-native-id',
          );
          if (value === null || value === '') return null;
          const parsed = Number(value);
          return Number.isSafeInteger(parsed) ? parsed : null;
        },
        renderHTML: (attributes) =>
          attributes.nativeId === null || attributes.nativeId === undefined
            ? {}
            : {
                'data-repeating-section-item-native-id': String(
                  attributes.nativeId,
                ),
              },
      },
    };
  },

  parseHTML() {
    return [
      {
        tag: 'div[data-document-repeating-section-item]',
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const properties = normalizeRepeatingSectionItemProperties(node.attrs);
    return [
      'div',
      mergeAttributes(HTMLAttributes, repeatingSectionItemDomAttributes(properties), {
        class: 'work-document-repeating-section-item',
      }),
      0,
    ];
  },
});

export const DocumentRepeatingSection = Node.create({
  name: 'documentRepeatingSection',
  group: 'block',
  content: 'documentRepeatingSectionItem+',
  defining: true,
  isolating: true,
  selectable: true,

  addAttributes() {
    return {
      id: {
        default: '',
        parseHTML: (element) =>
          element.getAttribute('data-repeating-section-id') ?? '',
        renderHTML: (attributes) => ({
          'data-repeating-section-id': attributes.id || undefined,
        }),
      },
      nativeId: {
        default: null,
        parseHTML: (element) => {
          const value = element.getAttribute('data-repeating-section-native-id');
          if (value === null || value === '') return null;
          const parsed = Number(value);
          return Number.isSafeInteger(parsed) ? parsed : null;
        },
        renderHTML: (attributes) =>
          attributes.nativeId === null || attributes.nativeId === undefined
            ? {}
            : {
                'data-repeating-section-native-id': String(attributes.nativeId),
              },
      },
      alias: {
        default: '',
        parseHTML: (element) =>
          element.getAttribute('data-repeating-section-alias') ?? '',
        renderHTML: (attributes) =>
          attributes.alias
            ? { 'data-repeating-section-alias': attributes.alias }
            : {},
      },
      tag: {
        default: '',
        parseHTML: (element) =>
          element.getAttribute('data-repeating-section-tag') ?? '',
        renderHTML: (attributes) =>
          attributes.tag
            ? { 'data-repeating-section-tag': attributes.tag }
            : {},
      },
      lock: {
        default: DOCUMENT_REPEATING_SECTION_DEFAULTS.lock,
        parseHTML: (element) =>
          element.getAttribute('data-repeating-section-lock') ??
          DOCUMENT_REPEATING_SECTION_DEFAULTS.lock,
        renderHTML: (attributes) => ({
          'data-repeating-section-lock': attributes.lock,
        }),
      },
      appearance: {
        default: DOCUMENT_REPEATING_SECTION_DEFAULTS.appearance,
        parseHTML: (element) =>
          element.getAttribute('data-repeating-section-appearance') ??
          DOCUMENT_REPEATING_SECTION_DEFAULTS.appearance,
        renderHTML: (attributes) => ({
          'data-repeating-section-appearance': attributes.appearance,
        }),
      },
      color: {
        default: null,
        parseHTML: (element) =>
          element.getAttribute('data-repeating-section-color'),
        renderHTML: (attributes) =>
          attributes.color
            ? { 'data-repeating-section-color': attributes.color }
            : {},
      },
    };
  },

  parseHTML() {
    return [
      {
        tag: 'div[data-document-repeating-section]',
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const properties = normalizeRepeatingSectionProperties(node.attrs);
    const label = properties.alias || properties.tag || '重复部分';
    return [
      'div',
      mergeAttributes(HTMLAttributes, repeatingSectionDomAttributes(properties), {
        class: 'work-document-repeating-section',
        role: 'group',
        'aria-label': label,
      }),
      0,
    ];
  },

  addCommands() {
    return {
      insertDocumentRepeatingSection:
        (text = '项目', options = {}) =>
        (props: CommandProps) =>
          insertDocumentRepeatingSectionCommand(props, text, options),
      setDocumentRepeatingSectionProperties:
        (value) =>
        ({ chain, state }) => {
          const selected = selectedRepeatingSection(state);
          if (!selected) return false;
          const next = normalizeRepeatingSectionProperties({
            ...selected.node.attrs,
            ...value,
            id: selected.node.attrs.id,
          });
          return chain()
            .focus()
            .command(({ tr, dispatch }) => {
              if (!dispatch) return true;
              closeHistory(tr);
              tr.setNodeMarkup(selected.position, undefined, next);
              dispatch(tr);
              return true;
            })
            .run();
        },
      deleteDocumentRepeatingSection:
        () =>
        ({ dispatch, editor, state, tr }) => {
          const selected = selectedRepeatingSection(state);
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
          editor.view.focus();
          return true;
        },
    };
  },

  addProseMirrorPlugins() {
    return [];
  },
});

export function normalizeRepeatingSectionProperties(
  value: Partial<WorkDocumentRepeatingSectionProperties> | null | undefined,
): WorkDocumentRepeatingSectionProperties {
  const lock = LOCK_VALUES.has(value?.lock as WorkDocumentRepeatingSectionLock)
    ? (value?.lock as WorkDocumentRepeatingSectionLock)
    : DOCUMENT_REPEATING_SECTION_DEFAULTS.lock;
  const appearance = APPEARANCE_VALUES.has(
    value?.appearance as WorkDocumentRepeatingSectionAppearance,
  )
    ? (value?.appearance as WorkDocumentRepeatingSectionAppearance)
    : DOCUMENT_REPEATING_SECTION_DEFAULTS.appearance;
  const color =
    typeof value?.color === 'string' && COLOR_PATTERN.test(value.color)
      ? value.color.toLowerCase()
      : null;
  const nativeId =
    typeof value?.nativeId === 'number' &&
    Number.isSafeInteger(value.nativeId) &&
    value.nativeId >= -2_147_483_648 &&
    value.nativeId <= 2_147_483_647
      ? value.nativeId
      : null;
  return {
    id:
      typeof value?.id === 'string' && value.id.trim()
        ? value.id.trim().slice(0, 128)
        : '',
    nativeId,
    alias:
      typeof value?.alias === 'string'
        ? value.alias.replace(/[\u0000-\u001f\u007f]/gu, '').slice(0, 255)
        : '',
    tag:
      typeof value?.tag === 'string'
        ? value.tag.replace(/[\u0000-\u001f\u007f]/gu, '').slice(0, 255)
        : '',
    lock,
    appearance,
    color,
  };
}

export function normalizeRepeatingSectionItemProperties(
  value: Partial<WorkDocumentRepeatingSectionItemProperties> | null | undefined,
): WorkDocumentRepeatingSectionItemProperties {
  const nativeId =
    typeof value?.nativeId === 'number' &&
    Number.isSafeInteger(value.nativeId) &&
    value.nativeId >= -2_147_483_648 &&
    value.nativeId <= 2_147_483_647
      ? value.nativeId
      : null;
  return {
    id:
      typeof value?.id === 'string' && value.id.trim()
        ? value.id.trim().slice(0, 128)
        : '',
    nativeId,
  };
}

export function repeatingSectionDomAttributes(
  properties: WorkDocumentRepeatingSectionProperties,
): Record<string, string | undefined> {
  return {
    'data-document-repeating-section': 'true',
    'data-repeating-section-id': properties.id || undefined,
    'data-repeating-section-native-id':
      properties.nativeId === null ? undefined : String(properties.nativeId),
    'data-repeating-section-alias': properties.alias || undefined,
    'data-repeating-section-tag': properties.tag || undefined,
    'data-repeating-section-lock': properties.lock,
    'data-repeating-section-appearance': properties.appearance,
    'data-repeating-section-color': properties.color ?? undefined,
  };
}

export function repeatingSectionItemDomAttributes(
  properties: WorkDocumentRepeatingSectionItemProperties,
): Record<string, string | undefined> {
  return {
    'data-document-repeating-section-item': 'true',
    'data-repeating-section-item-id': properties.id || undefined,
    'data-repeating-section-item-native-id':
      properties.nativeId === null ? undefined : String(properties.nativeId),
  };
}

export function repeatingSectionPropertiesFromElement(
  element: HTMLElement,
): WorkDocumentRepeatingSectionProperties {
  return normalizeRepeatingSectionProperties({
    id: element.getAttribute('data-repeating-section-id') ?? '',
    nativeId: (() => {
      const value = element.getAttribute('data-repeating-section-native-id');
      if (value === null || value === '') return null;
      const parsed = Number(value);
      return Number.isSafeInteger(parsed) ? parsed : null;
    })(),
    alias: element.getAttribute('data-repeating-section-alias') ?? '',
    tag: element.getAttribute('data-repeating-section-tag') ?? '',
    lock: element.getAttribute(
      'data-repeating-section-lock',
    ) as WorkDocumentRepeatingSectionLock,
    appearance: element.getAttribute(
      'data-repeating-section-appearance',
    ) as WorkDocumentRepeatingSectionAppearance,
    color: element.getAttribute('data-repeating-section-color'),
  });
}

export function repeatingSectionItemPropertiesFromElement(
  element: HTMLElement,
): WorkDocumentRepeatingSectionItemProperties {
  return normalizeRepeatingSectionItemProperties({
    id: element.getAttribute('data-repeating-section-item-id') ?? '',
    nativeId: (() => {
      const value = element.getAttribute(
        'data-repeating-section-item-native-id',
      );
      if (value === null || value === '') return null;
      const parsed = Number(value);
      return Number.isSafeInteger(parsed) ? parsed : null;
    })(),
  });
}

function insertDocumentRepeatingSectionCommand(
  { dispatch, editor, state, tr }: CommandProps,
  text: string,
  options: Partial<WorkDocumentRepeatingSectionProperties>,
): boolean {
  const sectionType = state.schema.nodes.documentRepeatingSection;
  const itemType = state.schema.nodes.documentRepeatingSectionItem;
  const paragraphType = state.schema.nodes.paragraph;
  if (!sectionType || !itemType || !paragraphType) return false;
  const properties = normalizeRepeatingSectionProperties({
    ...DOCUMENT_REPEATING_SECTION_DEFAULTS,
    ...options,
    id: options.id?.trim() || createWorkId('repeating-section'),
  });
  const item = itemType.create(
    {
      id: createWorkId('repeating-section-item'),
      nativeId: null,
    },
    paragraphType.create(
      null,
      text ? state.schema.text(text.slice(0, 10_000)) : undefined,
    ),
  );
  const section = sectionType.create(properties, item);
  if (!dispatch) return true;
  closeHistory(tr);
  const { from, to } = state.selection;
  tr.replaceWith(from, to, section);
  const resolved = tr.doc.resolve(
    Math.min(from + 2, tr.doc.content.size),
  );
  tr.setSelection(TextSelection.near(resolved));
  dispatch(tr.scrollIntoView());
  editor.view.focus();
  return true;
}

function selectedRepeatingSection(state: CommandProps['state']): {
  node: import('@tiptap/pm/model').Node;
  position: number;
} | null {
  const { $from } = state.selection;
  for (let depth = $from.depth; depth >= 0; depth -= 1) {
    const node = $from.node(depth);
    if (node.type.name === 'documentRepeatingSection') {
      return { node, position: $from.before(depth) };
    }
  }
  return null;
}
