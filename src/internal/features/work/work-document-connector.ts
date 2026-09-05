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

/** The bounded straight connector subset shared by the Writer and DOCX paths. */
export type WorkDocumentConnectorLayout = 'inline' | 'floating';
export type WorkDocumentConnectorArrow = 'none' | 'triangle';
export type WorkDocumentConnectorReference = 'column' | 'margin' | 'page';
export type WorkDocumentConnectorVerticalReference =
  | 'paragraph'
  | 'margin'
  | 'page';

export interface WorkDocumentConnectorProperties {
  id: string;
  width: number;
  height: number;
  layout: WorkDocumentConnectorLayout;
  horizontalOffset: number | null;
  verticalOffset: number | null;
  horizontalReference: WorkDocumentConnectorReference;
  verticalReference: WorkDocumentConnectorVerticalReference;
  startX: number;
  startY: number;
  endX: number;
  endY: number;
  lineColor: string;
  lineWidth: number;
  startArrow: WorkDocumentConnectorArrow;
  endArrow: WorkDocumentConnectorArrow;
  docPropertiesId: number | null;
}

export interface DocumentConnectorCommandOptions {
  restoreFocus?: boolean;
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentConnector: {
      insertDocumentConnector: (
        options?: Partial<WorkDocumentConnectorProperties>,
      ) => ReturnType;
      setDocumentConnectorProperties: (
        value: Partial<WorkDocumentConnectorProperties>,
        options?: DocumentConnectorCommandOptions,
      ) => ReturnType;
      deleteDocumentConnector: (
        options?: DocumentConnectorCommandOptions,
      ) => ReturnType;
    };
  }
}

export const DOCUMENT_CONNECTOR_DEFAULTS: WorkDocumentConnectorProperties = {
  id: '',
  width: 120,
  height: 25,
  layout: 'inline',
  horizontalOffset: null,
  verticalOffset: null,
  horizontalReference: 'column',
  verticalReference: 'paragraph',
  startX: 0,
  startY: 50,
  endX: 100,
  endY: 50,
  lineColor: '#c00000',
  lineWidth: 0.35,
  startArrow: 'none',
  endArrow: 'none',
  docPropertiesId: null,
};

export const DOCUMENT_CONNECTOR_LIMITS = {
  width: { min: 10, max: 558.7 },
  height: { min: 0.5, max: 558.7 },
  offset: { min: -558.7, max: 558.7 },
  endpoint: { min: 0, max: 100 },
  lineWidth: { min: 0.1, max: 10 },
} as const;

const CONNECTOR_ID_MAX_LENGTH = 160;
const CONNECTOR_COLOR_PATTERN = /^#[0-9a-f]{6}$/i;
const CONNECTOR_MARKER_ATTRIBUTES = [
  'id',
  'width',
  'height',
  'layout',
  'horizontalOffset',
  'verticalOffset',
  'horizontalReference',
  'verticalReference',
  'startX',
  'startY',
  'endX',
  'endY',
  'lineColor',
  'lineWidth',
  'startArrow',
  'endArrow',
  'docPropertiesId',
] as const;

export const DocumentConnector = Node.create({
  name: 'documentConnector',
  group: 'block',
  atom: true,
  defining: true,
  isolating: true,
  selectable: true,
  draggable: true,

  addAttributes() {
    return {
      id: dataAttribute('id', ''),
      width: dataAttribute('width', DOCUMENT_CONNECTOR_DEFAULTS.width),
      height: dataAttribute('height', DOCUMENT_CONNECTOR_DEFAULTS.height),
      layout: dataAttribute('layout', DOCUMENT_CONNECTOR_DEFAULTS.layout),
      horizontalOffset: nullableDataAttribute('horizontal-offset'),
      verticalOffset: nullableDataAttribute('vertical-offset'),
      horizontalReference: dataAttribute(
        'horizontal-reference',
        DOCUMENT_CONNECTOR_DEFAULTS.horizontalReference,
      ),
      verticalReference: dataAttribute(
        'vertical-reference',
        DOCUMENT_CONNECTOR_DEFAULTS.verticalReference,
      ),
      startX: dataAttribute('start-x', DOCUMENT_CONNECTOR_DEFAULTS.startX),
      startY: dataAttribute('start-y', DOCUMENT_CONNECTOR_DEFAULTS.startY),
      endX: dataAttribute('end-x', DOCUMENT_CONNECTOR_DEFAULTS.endX),
      endY: dataAttribute('end-y', DOCUMENT_CONNECTOR_DEFAULTS.endY),
      lineColor: dataAttribute(
        'line-color',
        DOCUMENT_CONNECTOR_DEFAULTS.lineColor,
      ),
      lineWidth: dataAttribute(
        'line-width',
        DOCUMENT_CONNECTOR_DEFAULTS.lineWidth,
      ),
      startArrow: dataAttribute(
        'start-arrow',
        DOCUMENT_CONNECTOR_DEFAULTS.startArrow,
      ),
      endArrow: dataAttribute(
        'end-arrow',
        DOCUMENT_CONNECTOR_DEFAULTS.endArrow,
      ),
      docPropertiesId: nullableDataAttribute('doc-properties-id'),
    };
  },

  parseHTML() {
    return [
      {
        tag: 'div[data-document-connector]',
        getAttrs: (element) => {
          if (!(element instanceof HTMLElement)) return false;
          return connectorAttributesFromElement(element);
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const properties = normalizeDocumentConnectorProperties(node.attrs);
    const markerId = connectorMarkerId(properties.id);
    const startMarker = `${markerId}-start`;
    const endMarker = `${markerId}-end`;
    return [
      'div',
      mergeAttributes(HTMLAttributes, connectorDomAttributes(properties), {
        class: 'work-document-connector',
        contenteditable: undefined,
        style: connectorCss(properties),
        role: 'img',
        'aria-label': '直线连接符',
      }),
      [
        'svg',
        {
          class: 'work-document-connector-svg',
          viewBox: '0 0 100 100',
          preserveAspectRatio: 'none',
          focusable: 'false',
          'aria-hidden': 'true',
        },
        [
          'defs',
          ...(properties.startArrow === 'triangle'
            ? [connectorMarker(startMarker, properties.lineColor)]
            : []),
          ...(properties.endArrow === 'triangle'
            ? [connectorMarker(endMarker, properties.lineColor)]
            : []),
        ],
        [
          'line',
          {
            class: 'work-document-connector-line',
            x1: formatNumber(properties.startX),
            y1: formatNumber(properties.startY),
            x2: formatNumber(properties.endX),
            y2: formatNumber(properties.endY),
            stroke: properties.lineColor,
            'stroke-width': connectorStrokeWidth(properties.lineWidth),
            'stroke-linecap': 'round',
            'marker-start':
              properties.startArrow === 'triangle'
                ? `url(#${startMarker})`
                : undefined,
            'marker-end':
              properties.endArrow === 'triangle'
                ? `url(#${endMarker})`
                : undefined,
          },
        ],
      ],
    ];
  },

  renderText() {
    return '';
  },

  addCommands() {
    return {
      insertDocumentConnector:
        (options = {}) =>
        ({ dispatch, editor, state, tr }: CommandProps) =>
          insertDocumentConnectorCommand(
            { dispatch, editor, state, tr },
            options,
          ),
      setDocumentConnectorProperties:
        (value, options = {}) =>
        ({ chain, state, tr }) => {
          if (!selectedDocumentConnector(state)) return false;
          const attributes = connectorAttributesForChanges(value);
          if (!Object.keys(attributes).length) return false;
          closeHistory(tr);
          let commandChain = chain();
          if (options.restoreFocus !== false)
            commandChain = commandChain.focus();
          return commandChain.updateAttributes(this.name, attributes).run();
        },
      deleteDocumentConnector:
        (options = {}) =>
        ({ dispatch, editor, state, tr }) => {
          const selected = selectedDocumentConnector(state);
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
            const current = normalizeDocumentConnectorId(node.attrs.id);
            const id =
              current && !seen.has(current)
                ? current
                : createWorkId('connector');
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

export function documentConnectorProperties(
  editor: Editor,
): WorkDocumentConnectorProperties {
  return normalizeDocumentConnectorProperties(
    editor.getAttributes('documentConnector') as Record<string, unknown>,
  );
}

export function setDocumentConnectorProperties(
  editor: Editor,
  value: Partial<WorkDocumentConnectorProperties>,
  options: DocumentConnectorCommandOptions = {},
): boolean {
  return editor.commands.setDocumentConnectorProperties(value, options);
}

export function normalizeDocumentConnectorProperties(
  value: Partial<Record<keyof WorkDocumentConnectorProperties, unknown>>,
): WorkDocumentConnectorProperties {
  return {
    id: normalizeDocumentConnectorId(value.id),
    width: boundedNumber(
      value.width,
      DOCUMENT_CONNECTOR_DEFAULTS.width,
      DOCUMENT_CONNECTOR_LIMITS.width.min,
      DOCUMENT_CONNECTOR_LIMITS.width.max,
    ),
    height: boundedNumber(
      value.height,
      DOCUMENT_CONNECTOR_DEFAULTS.height,
      DOCUMENT_CONNECTOR_LIMITS.height.min,
      DOCUMENT_CONNECTOR_LIMITS.height.max,
    ),
    layout: value.layout === 'floating' ? 'floating' : 'inline',
    horizontalOffset: nullableBoundedNumber(
      value.horizontalOffset,
      DOCUMENT_CONNECTOR_LIMITS.offset.min,
      DOCUMENT_CONNECTOR_LIMITS.offset.max,
    ),
    verticalOffset: nullableBoundedNumber(
      value.verticalOffset,
      DOCUMENT_CONNECTOR_LIMITS.offset.min,
      DOCUMENT_CONNECTOR_LIMITS.offset.max,
    ),
    horizontalReference: connectorHorizontalReference(
      value.horizontalReference,
    ),
    verticalReference: connectorVerticalReference(value.verticalReference),
    startX: boundedNumber(
      value.startX,
      DOCUMENT_CONNECTOR_DEFAULTS.startX,
      DOCUMENT_CONNECTOR_LIMITS.endpoint.min,
      DOCUMENT_CONNECTOR_LIMITS.endpoint.max,
    ),
    startY: boundedNumber(
      value.startY,
      DOCUMENT_CONNECTOR_DEFAULTS.startY,
      DOCUMENT_CONNECTOR_LIMITS.endpoint.min,
      DOCUMENT_CONNECTOR_LIMITS.endpoint.max,
    ),
    endX: boundedNumber(
      value.endX,
      DOCUMENT_CONNECTOR_DEFAULTS.endX,
      DOCUMENT_CONNECTOR_LIMITS.endpoint.min,
      DOCUMENT_CONNECTOR_LIMITS.endpoint.max,
    ),
    endY: boundedNumber(
      value.endY,
      DOCUMENT_CONNECTOR_DEFAULTS.endY,
      DOCUMENT_CONNECTOR_LIMITS.endpoint.min,
      DOCUMENT_CONNECTOR_LIMITS.endpoint.max,
    ),
    lineColor: normalizeConnectorColor(value.lineColor),
    lineWidth: boundedNumber(
      value.lineWidth,
      DOCUMENT_CONNECTOR_DEFAULTS.lineWidth,
      DOCUMENT_CONNECTOR_LIMITS.lineWidth.min,
      DOCUMENT_CONNECTOR_LIMITS.lineWidth.max,
    ),
    startArrow: connectorArrow(value.startArrow),
    endArrow: connectorArrow(value.endArrow),
    docPropertiesId: nullableInteger(value.docPropertiesId, 0, 0xffff_ffff),
  };
}

export function connectorCss(
  value: Partial<Record<keyof WorkDocumentConnectorProperties, unknown>>,
): string {
  const properties = normalizeDocumentConnectorProperties(value);
  return [
    `--work-document-connector-width:${formatNumber(properties.width)}mm`,
    `--work-document-connector-height:${formatNumber(properties.height)}mm`,
    `--work-document-connector-line-color:${properties.lineColor}`,
    `--work-document-connector-line-width:${formatNumber(properties.lineWidth)}mm`,
    ...(properties.layout === 'floating'
      ? [
          `--work-document-connector-horizontal-offset:${formatNullableNumber(properties.horizontalOffset)}mm`,
          `--work-document-connector-vertical-offset:${formatNullableNumber(properties.verticalOffset)}mm`,
        ]
      : []),
  ].join(';');
}

export function connectorDomAttributes(
  value: Partial<Record<keyof WorkDocumentConnectorProperties, unknown>>,
): Record<string, string | undefined> {
  const properties = normalizeDocumentConnectorProperties(value);
  return {
    'data-document-connector': 'true',
    'data-connector-id': properties.id || undefined,
    'data-connector-width': formatNumber(properties.width),
    'data-connector-height': formatNumber(properties.height),
    'data-connector-layout': properties.layout,
    'data-connector-horizontal-offset':
      properties.horizontalOffset === null
        ? undefined
        : formatNumber(properties.horizontalOffset),
    'data-connector-vertical-offset':
      properties.verticalOffset === null
        ? undefined
        : formatNumber(properties.verticalOffset),
    'data-connector-horizontal-reference': properties.horizontalReference,
    'data-connector-vertical-reference': properties.verticalReference,
    'data-connector-start-x': formatNumber(properties.startX),
    'data-connector-start-y': formatNumber(properties.startY),
    'data-connector-end-x': formatNumber(properties.endX),
    'data-connector-end-y': formatNumber(properties.endY),
    'data-connector-line-color': properties.lineColor,
    'data-connector-line-width': formatNumber(properties.lineWidth),
    'data-connector-start-arrow': properties.startArrow,
    'data-connector-end-arrow': properties.endArrow,
    'data-connector-doc-properties-id':
      properties.docPropertiesId === null
        ? undefined
        : String(properties.docPropertiesId),
  };
}

export function documentConnectorPropertiesFromElement(
  element: Element,
): WorkDocumentConnectorProperties {
  return normalizeDocumentConnectorProperties({
    id: element.getAttribute('data-connector-id'),
    width: element.getAttribute('data-connector-width'),
    height: element.getAttribute('data-connector-height'),
    layout: element.getAttribute('data-connector-layout'),
    horizontalOffset: element.getAttribute('data-connector-horizontal-offset'),
    verticalOffset: element.getAttribute('data-connector-vertical-offset'),
    horizontalReference: element.getAttribute(
      'data-connector-horizontal-reference',
    ),
    verticalReference: element.getAttribute(
      'data-connector-vertical-reference',
    ),
    startX: element.getAttribute('data-connector-start-x'),
    startY: element.getAttribute('data-connector-start-y'),
    endX: element.getAttribute('data-connector-end-x'),
    endY: element.getAttribute('data-connector-end-y'),
    lineColor: element.getAttribute('data-connector-line-color'),
    lineWidth: element.getAttribute('data-connector-line-width'),
    startArrow: element.getAttribute('data-connector-start-arrow'),
    endArrow: element.getAttribute('data-connector-end-arrow'),
    docPropertiesId: element.getAttribute('data-connector-doc-properties-id'),
  });
}

function insertDocumentConnectorCommand(
  {
    dispatch,
    editor,
    state,
    tr,
  }: Pick<CommandProps, 'dispatch' | 'editor' | 'state' | 'tr'>,
  options: Partial<WorkDocumentConnectorProperties>,
): boolean {
  const section = activeDocumentSectionFromState(state);
  const connectorType = editor.schema.nodes.documentConnector;
  const paragraphType = editor.schema.nodes.paragraph;
  if (!section || !connectorType || !paragraphType) return false;
  const child = activeSectionChild(section, state.selection.from);
  if (!child) return false;
  if (!dispatch) return true;
  const properties = normalizeDocumentConnectorProperties({
    ...DOCUMENT_CONNECTOR_DEFAULTS,
    ...options,
    id: options.id || createWorkId('connector'),
  });
  const connector = connectorType.create(properties);
  const insertPosition = section.position + 1 + child.offset + child.nodeSize;
  tr.insert(insertPosition, connector);
  if (child.index === section.node.childCount - 1) {
    tr.insert(insertPosition + connector.nodeSize, paragraphType.create());
  }
  tr.setSelection(NodeSelection.create(tr.doc, insertPosition));
  tr.scrollIntoView();
  return true;
}

function selectedDocumentConnector(
  state: Editor['state'],
): { node: ProseMirrorNode; position: number } | null {
  const { $from } = state.selection;
  for (let depth = $from.depth; depth > 0; depth -= 1) {
    const node = $from.node(depth);
    if (node.type.name === 'documentConnector') {
      return { node, position: $from.before(depth) };
    }
  }
  if (
    state.selection instanceof NodeSelection &&
    state.selection.node.type.name === 'documentConnector'
  ) {
    return { node: state.selection.node, position: state.selection.from };
  }
  return null;
}

function connectorAttributesForChanges(
  value: Partial<WorkDocumentConnectorProperties>,
): Record<string, unknown> {
  const attributes: Record<string, unknown> = {};
  for (const name of CONNECTOR_MARKER_ATTRIBUTES) {
    if (name in value) {
      attributes[name] = normalizeDocumentConnectorProperties({
        ...DOCUMENT_CONNECTOR_DEFAULTS,
        ...value,
      })[name];
    }
  }
  return attributes;
}

function connectorAttributesFromElement(
  element: HTMLElement,
): Record<string, unknown> {
  return {
    id: element.dataset.connectorId ?? '',
    width: element.dataset.connectorWidth,
    height: element.dataset.connectorHeight,
    layout: element.dataset.connectorLayout,
    horizontalOffset: element.dataset.connectorHorizontalOffset,
    verticalOffset: element.dataset.connectorVerticalOffset,
    horizontalReference: element.dataset.connectorHorizontalReference,
    verticalReference: element.dataset.connectorVerticalReference,
    startX: element.dataset.connectorStartX,
    startY: element.dataset.connectorStartY,
    endX: element.dataset.connectorEndX,
    endY: element.dataset.connectorEndY,
    lineColor: element.dataset.connectorLineColor,
    lineWidth: element.dataset.connectorLineWidth,
    startArrow: element.dataset.connectorStartArrow,
    endArrow: element.dataset.connectorEndArrow,
    docPropertiesId: element.dataset.connectorDocPropertiesId,
  };
}

function connectorMarker(marker: string, color: string) {
  return [
    'marker',
    {
      id: marker,
      markerWidth: '6',
      markerHeight: '6',
      refX: '5',
      refY: '3',
      orient: 'auto-start-reverse',
      markerUnits: 'strokeWidth',
      viewBox: '0 0 6 6',
    },
    ['path', { d: 'M 0 0 L 6 3 L 0 6 z', fill: color }],
  ];
}

function connectorMarkerId(id: string): string {
  const safe = id.replace(/[^a-zA-Z0-9_-]/g, '-').slice(0, 80);
  return `a3s-connector-${safe || 'default'}`;
}

function connectorStrokeWidth(width: number): string {
  return formatNumber(Math.max(0.8, Math.min(8, width * 2.8)));
}

function normalizeDocumentConnectorId(value: unknown): string {
  return typeof value === 'string'
    ? value.trim().slice(0, CONNECTOR_ID_MAX_LENGTH)
    : '';
}

function connectorHorizontalReference(
  value: unknown,
): WorkDocumentConnectorReference {
  return value === 'margin' || value === 'page' ? value : 'column';
}

function connectorVerticalReference(
  value: unknown,
): WorkDocumentConnectorVerticalReference {
  return value === 'margin' || value === 'page' ? value : 'paragraph';
}

function connectorArrow(value: unknown): WorkDocumentConnectorArrow {
  return value === 'triangle' ? 'triangle' : 'none';
}

function normalizeConnectorColor(value: unknown): string {
  if (typeof value !== 'string') return DOCUMENT_CONNECTOR_DEFAULTS.lineColor;
  const normalized = value.trim().toLowerCase();
  return CONNECTOR_COLOR_PATTERN.test(normalized)
    ? normalized
    : DOCUMENT_CONNECTOR_DEFAULTS.lineColor;
}

function boundedNumber(
  value: unknown,
  fallback: number,
  min: number,
  max: number,
): number {
  const number = typeof value === 'number' ? value : Number(value);
  if (!Number.isFinite(number)) return fallback;
  return Math.min(max, Math.max(min, number));
}

function nullableBoundedNumber(
  value: unknown,
  min: number,
  max: number,
): number | null {
  if (value === null || value === undefined || value === '') return null;
  const number = Number(value);
  return Number.isFinite(number) ? Math.min(max, Math.max(min, number)) : null;
}

function nullableInteger(
  value: unknown,
  min: number,
  max: number,
): number | null {
  if (value === null || value === undefined || value === '') return null;
  const number = Number(value);
  return Number.isSafeInteger(number) && number >= min && number <= max
    ? number
    : null;
}

function formatNumber(value: number): string {
  return Number(value.toFixed(2)).toString();
}

function formatNullableNumber(value: number | null): string {
  return formatNumber(value ?? 0);
}

function dataAttribute(dataName: string, defaultValue: unknown) {
  return {
    default: defaultValue,
    parseHTML: (element: Element) =>
      element.getAttribute(`data-connector-${dataName}`) ?? defaultValue,
    rendered: false,
  };
}

function nullableDataAttribute(dataName: string) {
  return {
    default: null,
    parseHTML: (element: Element) =>
      element.getAttribute(`data-connector-${dataName}`),
    rendered: false,
  };
}

interface ActiveSectionChild {
  node: ProseMirrorNode;
  offset: number;
  nodeSize: number;
  index: number;
}

function activeSectionChild(
  section: { node: ProseMirrorNode; position: number },
  position: number,
): ActiveSectionChild | null {
  let offset = 0;
  for (let index = 0; index < section.node.childCount; index += 1) {
    const node = section.node.child(index);
    const start = section.position + 1 + offset;
    if (position >= start && position <= start + node.nodeSize) {
      return { node, offset, nodeSize: node.nodeSize, index };
    }
    offset += node.nodeSize;
  }
  const lastIndex = section.node.childCount - 1;
  const last = section.node.child(lastIndex);
  return last
    ? {
        node: last,
        offset: Math.max(0, section.node.content.size - last.nodeSize),
        nodeSize: last.nodeSize,
        index: lastIndex,
      }
    : null;
}
