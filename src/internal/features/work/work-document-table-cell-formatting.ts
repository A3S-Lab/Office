import { type CommandProps, Extension } from '@tiptap/core';
import { TableCell, TableHeader } from '@tiptap/extension-table';
import type { Node as ProseMirrorNode, ResolvedPos } from '@tiptap/pm/model';
import type { EditorState, Selection } from '@tiptap/pm/state';
import { CellSelection } from '@tiptap/pm/tables';

export type DocumentTableVerticalAlign = 'top' | 'middle' | 'bottom';
export type DocumentTableHorizontalAlign = 'left' | 'center' | 'right';
export type DocumentTableBorderStyle =
  | 'solid'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'none';

export interface DocumentTableCellFormat {
  backgroundColor: string;
  verticalAlign: DocumentTableVerticalAlign;
  borderColor: string;
  borderStyle: DocumentTableBorderStyle;
  borderWidth: number;
}

export type DocumentTableCellFormatPatch = Partial<DocumentTableCellFormat>;

export type DocumentTableStyleId =
  | 'grid'
  | 'blueStripe'
  | 'greenStripe'
  | 'grayStripe'
  | 'clean';

export interface DocumentTableStyleOption {
  id: DocumentTableStyleId;
  label: string;
  headerColor: string;
  bodyColor: string;
  alternateColor: string;
  borderColor: string;
  borderStyle: DocumentTableBorderStyle;
  borderWidth: number;
}

export const DOCUMENT_TABLE_STYLE_OPTIONS: readonly DocumentTableStyleOption[] =
  [
    {
      id: 'grid',
      label: '网格',
      headerColor: '#f1f4f9',
      bodyColor: '#ffffff',
      alternateColor: '#ffffff',
      borderColor: '#cfd5df',
      borderStyle: 'solid',
      borderWidth: 1,
    },
    {
      id: 'blueStripe',
      label: '蓝色条纹',
      headerColor: '#d9eaf7',
      bodyColor: '#ffffff',
      alternateColor: '#f7fbff',
      borderColor: '#9fbad0',
      borderStyle: 'solid',
      borderWidth: 1,
    },
    {
      id: 'greenStripe',
      label: '绿色条纹',
      headerColor: '#ddeee7',
      bodyColor: '#ffffff',
      alternateColor: '#f4faf7',
      borderColor: '#abc8be',
      borderStyle: 'solid',
      borderWidth: 1,
    },
    {
      id: 'grayStripe',
      label: '灰色条纹',
      headerColor: '#e7e9ed',
      bodyColor: '#ffffff',
      alternateColor: '#f6f7f9',
      borderColor: '#b9c0ca',
      borderStyle: 'solid',
      borderWidth: 1,
    },
    {
      id: 'clean',
      label: '简洁',
      headerColor: '#f1f4f9',
      bodyColor: '#ffffff',
      alternateColor: '#ffffff',
      borderColor: '#cfd5df',
      borderStyle: 'none',
      borderWidth: 0,
    },
  ];

export const DEFAULT_DOCUMENT_TABLE_CELL_FORMAT: DocumentTableCellFormat = {
  backgroundColor: '#ffffff',
  verticalAlign: 'top',
  borderColor: '#cfd5df',
  borderStyle: 'solid',
  borderWidth: 1,
};

export const DEFAULT_DOCUMENT_TABLE_HEADER_FORMAT: DocumentTableCellFormat = {
  ...DEFAULT_DOCUMENT_TABLE_CELL_FORMAT,
  backgroundColor: '#f1f4f9',
};

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentTableFormatting: {
      setDocumentTableCellFormat: (
        format: DocumentTableCellFormatPatch,
      ) => ReturnType;
      setDocumentTableHorizontalAlignment: (
        alignment: DocumentTableHorizontalAlign,
      ) => ReturnType;
      applyDocumentTableStyle: (style: DocumentTableStyleId) => ReturnType;
    };
  }
}

export const DocumentTableCell = TableCell.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      ...documentTableCellAttributes(DEFAULT_DOCUMENT_TABLE_CELL_FORMAT),
    };
  },
});

export const DocumentTableHeader = TableHeader.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      ...documentTableCellAttributes(DEFAULT_DOCUMENT_TABLE_HEADER_FORMAT),
    };
  },
});

export const DocumentTableFormatting = Extension.create({
  name: 'documentTableFormatting',

  addCommands() {
    return {
      setDocumentTableCellFormat: (format) => (props) =>
        setSelectedCellFormat(props, format),
      setDocumentTableHorizontalAlignment: (alignment) => (props) =>
        setSelectedCellAlignment(props, alignment),
      applyDocumentTableStyle: (style) => (props) =>
        applyTableStyle(props, style),
    };
  },
});

export function documentTableCellFormat(
  state: EditorState,
): DocumentTableCellFormat | null {
  const position = selectedTableCellPositions(state.selection)[0];
  const node = position === undefined ? null : state.doc.nodeAt(position);
  return node && isTableCell(node) ? formatFromAttributes(node.attrs) : null;
}

export function documentTableHorizontalAlignment(
  state: EditorState,
): DocumentTableHorizontalAlign {
  const position = selectedTableCellPositions(state.selection)[0];
  const cell = position === undefined ? null : state.doc.nodeAt(position);
  let alignment: DocumentTableHorizontalAlign = 'left';
  cell?.descendants((node) => {
    if (node.type.name !== 'paragraph' && node.type.name !== 'heading') {
      return true;
    }
    if (node.attrs.textAlign === 'center' || node.attrs.textAlign === 'right') {
      alignment = node.attrs.textAlign;
    }
    return false;
  });
  return alignment;
}

export function activeDocumentTableStyle(
  state: EditorState,
): DocumentTableStyleId | null {
  const context = selectedTable(state.selection);
  if (!context) return null;
  for (const style of DOCUMENT_TABLE_STYLE_OPTIONS) {
    let matches = true;
    context.node.forEach((row, _rowOffset, rowIndex) => {
      row.forEach((cell) => {
        const expected = tableStyleCellFormat(style, cell, rowIndex);
        const actual = formatFromAttributes(cell.attrs);
        if (
          expected.backgroundColor !== actual.backgroundColor ||
          expected.borderColor !== actual.borderColor ||
          expected.borderStyle !== actual.borderStyle ||
          expected.borderWidth !== actual.borderWidth
        ) {
          matches = false;
        }
      });
    });
    if (matches) return style.id;
  }
  return null;
}

function documentTableCellAttributes(defaults: DocumentTableCellFormat) {
  return {
    backgroundColor: {
      default: defaults.backgroundColor,
      parseHTML: (element: HTMLElement) =>
        normalizeTableColor(
          element.dataset.officeCellFill || element.style.backgroundColor,
        ) ?? defaults.backgroundColor,
      renderHTML: (attributes: Record<string, unknown>) => {
        const color =
          normalizeTableColor(String(attributes.backgroundColor ?? '')) ??
          defaults.backgroundColor;
        return {
          'data-office-cell-fill': color,
          style: `background-color: ${color}`,
        };
      },
    },
    verticalAlign: {
      default: defaults.verticalAlign,
      parseHTML: (element: HTMLElement) =>
        normalizeDocumentTableVerticalAlign(
          element.dataset.officeCellVerticalAlign ||
            element.style.verticalAlign,
        ) ?? defaults.verticalAlign,
      renderHTML: (attributes: Record<string, unknown>) => {
        const alignment =
          normalizeDocumentTableVerticalAlign(
            String(attributes.verticalAlign ?? ''),
          ) ?? defaults.verticalAlign;
        return {
          'data-office-cell-vertical-align': alignment,
          style: `vertical-align: ${alignment}`,
        };
      },
    },
    borderColor: {
      default: defaults.borderColor,
      parseHTML: (element: HTMLElement) =>
        normalizeTableColor(
          element.dataset.officeCellBorderColor || element.style.borderColor,
        ) ?? defaults.borderColor,
      renderHTML: (attributes: Record<string, unknown>) => {
        const format = formatFromAttributes(attributes, defaults);
        return {
          'data-office-cell-border-color': format.borderColor,
          'data-office-cell-border-style': format.borderStyle,
          'data-office-cell-border-width': String(format.borderWidth),
          style:
            format.borderStyle === 'none' || format.borderWidth === 0
              ? 'border: 0 none transparent'
              : `border: ${format.borderWidth}px ${format.borderStyle} ${format.borderColor}`,
        };
      },
    },
    borderStyle: {
      default: defaults.borderStyle,
      parseHTML: (element: HTMLElement) =>
        normalizeDocumentTableBorderStyle(
          element.dataset.officeCellBorderStyle || element.style.borderStyle,
        ) ?? defaults.borderStyle,
      renderHTML: () => ({}),
    },
    borderWidth: {
      default: defaults.borderWidth,
      parseHTML: (element: HTMLElement) =>
        normalizeDocumentTableBorderWidth(
          element.dataset.officeCellBorderWidth || element.style.borderWidth,
        ) ?? defaults.borderWidth,
      renderHTML: () => ({}),
    },
  };
}

function setSelectedCellFormat(
  { dispatch, state }: CommandProps,
  patch: DocumentTableCellFormatPatch,
): boolean {
  const normalized = normalizeFormatPatch(patch);
  if (!normalized) return false;
  const positions = selectedTableCellPositions(state.selection);
  if (!positions.length) return false;
  const transaction = state.tr;
  for (const position of positions) {
    const cell = transaction.doc.nodeAt(position);
    if (!cell || !isTableCell(cell)) continue;
    transaction.setNodeMarkup(position, undefined, {
      ...cell.attrs,
      ...normalized,
    });
  }
  if (dispatch && transaction.docChanged)
    dispatch(transaction.scrollIntoView());
  return true;
}

function setSelectedCellAlignment(
  { dispatch, state }: CommandProps,
  alignment: DocumentTableHorizontalAlign,
): boolean {
  if (alignment !== 'left' && alignment !== 'center' && alignment !== 'right') {
    return false;
  }
  const positions = selectedTableCellPositions(state.selection);
  if (!positions.length) return false;
  const transaction = state.tr;
  for (const position of positions) {
    const cell = transaction.doc.nodeAt(position);
    if (!cell || !isTableCell(cell)) continue;
    cell.descendants((node, offset) => {
      if (node.type.name !== 'paragraph' && node.type.name !== 'heading') {
        return true;
      }
      transaction.setNodeMarkup(position + 1 + offset, undefined, {
        ...node.attrs,
        textAlign: alignment,
      });
      return false;
    });
  }
  if (dispatch && transaction.docChanged)
    dispatch(transaction.scrollIntoView());
  return true;
}

function applyTableStyle(
  { dispatch, state }: CommandProps,
  styleId: DocumentTableStyleId,
): boolean {
  const style = DOCUMENT_TABLE_STYLE_OPTIONS.find(
    (candidate) => candidate.id === styleId,
  );
  const context = selectedTable(state.selection);
  if (!style || !context) return false;
  const transaction = state.tr;
  context.node.forEach((row, rowOffset, rowIndex) => {
    row.forEach((cell, cellOffset) => {
      const position = context.position + 2 + rowOffset + cellOffset;
      transaction.setNodeMarkup(position, undefined, {
        ...cell.attrs,
        ...tableStyleCellFormat(style, cell, rowIndex),
      });
    });
  });
  if (dispatch && transaction.docChanged)
    dispatch(transaction.scrollIntoView());
  return true;
}

function selectedTableCellPositions(selection: Selection): number[] {
  if (selection instanceof CellSelection) {
    const positions: number[] = [];
    selection.forEachCell((_node, position) => positions.push(position));
    return positions;
  }
  const position = ancestorPosition(selection.$from, isTableCell);
  return position === null ? [] : [position];
}

function selectedTable(
  selection: Selection,
): { node: ProseMirrorNode; position: number } | null {
  for (let depth = selection.$from.depth; depth > 0; depth -= 1) {
    const node = selection.$from.node(depth);
    if (node.type.spec.tableRole === 'table') {
      return { node, position: selection.$from.before(depth) };
    }
  }
  return null;
}

function ancestorPosition(
  position: ResolvedPos,
  predicate: (node: ProseMirrorNode) => boolean,
): number | null {
  for (let depth = position.depth; depth > 0; depth -= 1) {
    if (predicate(position.node(depth))) return position.before(depth);
  }
  return null;
}

function isTableCell(node: ProseMirrorNode): boolean {
  return (
    node.type.spec.tableRole === 'cell' ||
    node.type.spec.tableRole === 'header_cell'
  );
}

function tableStyleCellFormat(
  style: DocumentTableStyleOption,
  cell: ProseMirrorNode,
  rowIndex: number,
): DocumentTableCellFormatPatch {
  return {
    backgroundColor:
      cell.type.spec.tableRole === 'header_cell'
        ? style.headerColor
        : rowIndex % 2 === 1
          ? style.alternateColor
          : style.bodyColor,
    borderColor: style.borderColor,
    borderStyle: style.borderStyle,
    borderWidth: style.borderWidth,
  };
}

function normalizeFormatPatch(
  patch: DocumentTableCellFormatPatch,
): DocumentTableCellFormatPatch | null {
  const normalized: DocumentTableCellFormatPatch = {};
  if (patch.backgroundColor !== undefined) {
    const color = normalizeTableColor(patch.backgroundColor);
    if (!color) return null;
    normalized.backgroundColor = color;
  }
  if (patch.verticalAlign !== undefined) {
    const alignment = normalizeDocumentTableVerticalAlign(patch.verticalAlign);
    if (!alignment) return null;
    normalized.verticalAlign = alignment;
  }
  if (patch.borderColor !== undefined) {
    const color = normalizeTableColor(patch.borderColor);
    if (!color) return null;
    normalized.borderColor = color;
  }
  if (patch.borderStyle !== undefined) {
    const style = normalizeDocumentTableBorderStyle(patch.borderStyle);
    if (!style) return null;
    normalized.borderStyle = style;
  }
  if (patch.borderWidth !== undefined) {
    const width = normalizeDocumentTableBorderWidth(patch.borderWidth);
    if (width === null) return null;
    normalized.borderWidth = width;
  }
  return normalized;
}

function formatFromAttributes(
  attributes: Record<string, unknown>,
  defaults = DEFAULT_DOCUMENT_TABLE_CELL_FORMAT,
): DocumentTableCellFormat {
  return {
    backgroundColor:
      normalizeTableColor(String(attributes.backgroundColor ?? '')) ??
      defaults.backgroundColor,
    verticalAlign:
      normalizeDocumentTableVerticalAlign(
        String(attributes.verticalAlign ?? ''),
      ) ?? defaults.verticalAlign,
    borderColor:
      normalizeTableColor(String(attributes.borderColor ?? '')) ??
      defaults.borderColor,
    borderStyle:
      normalizeDocumentTableBorderStyle(String(attributes.borderStyle ?? '')) ??
      defaults.borderStyle,
    borderWidth:
      normalizeDocumentTableBorderWidth(attributes.borderWidth) ??
      defaults.borderWidth,
  };
}

export function normalizeTableColor(
  value: string | null | undefined,
): string | null {
  const trimmed = value?.trim().toLowerCase();
  if (!trimmed) return null;
  if (/^#[0-9a-f]{3}$/.test(trimmed)) {
    return `#${trimmed
      .slice(1)
      .split('')
      .map((part) => `${part}${part}`)
      .join('')}`;
  }
  if (/^#[0-9a-f]{6}$/.test(trimmed)) return trimmed;
  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/.exec(trimmed);
  if (!rgb) return null;
  const channels = rgb.slice(1, 4).map(Number);
  if (channels.some((channel) => channel < 0 || channel > 255)) return null;
  return `#${channels
    .map((channel) => channel.toString(16).padStart(2, '0'))
    .join('')}`;
}

export function normalizeDocumentTableVerticalAlign(
  value: string | null | undefined,
): DocumentTableVerticalAlign | null {
  if (value === 'top' || value === 'middle' || value === 'bottom') return value;
  if (value === 'center') return 'middle';
  return null;
}

export function normalizeDocumentTableBorderStyle(
  value: string | null | undefined,
): DocumentTableBorderStyle | null {
  if (
    value === 'solid' ||
    value === 'dashed' ||
    value === 'dotted' ||
    value === 'double' ||
    value === 'none'
  ) {
    return value;
  }
  return null;
}

export function normalizeDocumentTableBorderWidth(
  value: unknown,
): number | null {
  const width =
    typeof value === 'number'
      ? value
      : Number.parseFloat(String(value ?? '').replace(/px$/i, ''));
  if (!Number.isFinite(width) || width < 0 || width > 6) return null;
  return Math.round(width * 2) / 2;
}
