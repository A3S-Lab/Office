import { type CommandProps, Extension } from '@tiptap/core';
import { TableCell, TableHeader } from '@tiptap/extension-table';
import type { Node as ProseMirrorNode, ResolvedPos } from '@tiptap/pm/model';
import {
  type EditorState,
  NodeSelection,
  type Selection,
} from '@tiptap/pm/state';
import { CellSelection } from '@tiptap/pm/tables';
import {
  type DocumentTableBorder,
  type DocumentTableBorderStyle,
  type DocumentTableBorderTarget,
  type DocumentTableCellBorders,
  documentTableBordersFromAttributes,
  documentTableBordersFromElement,
  normalizeDocumentTableBorderStyle,
  normalizeDocumentTableBorderWidth,
  normalizeTableColor,
  renderDocumentTableBorders,
  setSelectedDocumentTableBorders,
  uniformDocumentTableBorder,
  uniformDocumentTableBorders,
} from './work-document-table-borders';
import {
  documentTableCellMarginOverridesFromElement,
  normalizeDocumentTableCellMarginOverrides,
  renderDocumentTableCellMarginOverrides,
  type DocumentTableCellMarginOverrides,
} from './work-document-table-geometry';
import {
  parseDocxThemeReference,
  serializeDocxThemeReference,
} from './work-docx-theme-reference';

export {
  documentTableBordersFromElement,
  normalizeDocumentTableBorder,
  normalizeDocumentTableBorderStyle,
  normalizeDocumentTableBorderWidth,
  normalizeTableColor,
  renderDocumentTableBorders,
  uniformDocumentTableBorder,
  uniformDocumentTableBorders,
} from './work-document-table-borders';
export type {
  DocumentTableBorder,
  DocumentTableBorderEdge,
  DocumentTableBorderStyle,
  DocumentTableBorderTarget,
  DocumentTableCellBorders,
} from './work-document-table-borders';

export type DocumentTableVerticalAlign = 'top' | 'middle' | 'bottom';
export type DocumentTableHorizontalAlign = 'left' | 'center' | 'right';
export interface DocumentTableCellFormat {
  backgroundColor: string;
  verticalAlign: DocumentTableVerticalAlign;
  borderColor: string;
  borderStyle: DocumentTableBorderStyle;
  borderWidth: number;
  borders: DocumentTableCellBorders;
  margins: DocumentTableCellMarginOverrides | null;
}

export type DocumentTableCellFormatPatch = Partial<
  Omit<DocumentTableCellFormat, 'borders'>
>;

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
  borders: uniformDocumentTableBorders({
    color: '#cfd5df',
    style: 'solid',
    width: 1,
  }),
  margins: null,
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
      setDocumentTableBorders: (
        target: DocumentTableBorderTarget,
        border: DocumentTableBorder,
      ) => ReturnType;
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
      setDocumentTableBorders: (target, border) => (props) =>
        setSelectedDocumentTableBorders(props, target, border),
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
        const expectedBorder = {
          color: style.borderColor,
          style: style.borderStyle,
          width: style.borderWidth,
        } satisfies DocumentTableBorder;
        const actualBorder = uniformDocumentTableBorder(actual.borders);
        if (
          expected.backgroundColor !== actual.backgroundColor ||
          !actualBorder ||
          actualBorder.color !== expectedBorder.color ||
          actualBorder.style !== expectedBorder.style ||
          actualBorder.width !== expectedBorder.width
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
    themeFill: {
      default: null,
      parseHTML: (element: HTMLElement) =>
        serializeDocxThemeReference(
          parseDocxThemeReference(element.dataset.officeCellThemeFill),
        ) ?? null,
      renderHTML: (attributes: Record<string, unknown>) => {
        const value = serializeDocxThemeReference(
          parseDocxThemeReference(String(attributes.themeFill ?? '')),
        );
        return value ? { 'data-office-cell-theme-fill': value } : {};
      },
    },
    borderThemes: {
      default: null,
      parseHTML: (element: HTMLElement) => serializedCellBorderThemes(element),
      renderHTML: (attributes: Record<string, unknown>) =>
        renderedCellBorderThemes(attributes.borderThemes),
    },
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
      renderHTML: () => ({}),
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
    borders: {
      default: defaults.borders,
      parseHTML: (element: HTMLElement) =>
        documentTableBordersFromElement(element, {
          color: defaults.borderColor,
          style: defaults.borderStyle,
          width: defaults.borderWidth,
        }),
      renderHTML: (attributes: Record<string, unknown>) =>
        renderDocumentTableBorders(
          formatFromAttributes(attributes, defaults).borders,
        ),
    },
    margins: {
      default: defaults.margins,
      parseHTML: (element: HTMLElement) =>
        documentTableCellMarginOverridesFromElement(element),
      renderHTML: (attributes: Record<string, unknown>) =>
        renderDocumentTableCellMarginOverrides(attributes.margins),
    },
  };
}

function serializedCellBorderThemes(element: HTMLElement): string | null {
  const themes: Record<string, unknown> = {};
  for (const edge of ['Top', 'Right', 'Bottom', 'Left'] as const) {
    const reference = parseDocxThemeReference(
      element.dataset[`officeCellBorderTheme${edge}`],
    );
    if (reference) themes[edge.toLowerCase()] = reference;
  }
  return Object.keys(themes).length ? JSON.stringify(themes) : null;
}

function renderedCellBorderThemes(value: unknown): Record<string, string> {
  if (typeof value !== 'string' || !value) return {};
  try {
    const themes = JSON.parse(value) as Record<string, unknown>;
    const rendered: Record<string, string> = {};
    for (const edge of ['top', 'right', 'bottom', 'left'] as const) {
      const reference = serializeDocxThemeReference(
        parseDocxThemeReference(JSON.stringify(themes[edge] ?? null)),
      );
      if (reference) {
        rendered[`data-office-cell-border-theme-${edge}`] = reference;
      }
    }
    return rendered;
  } catch {
    return {};
  }
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
    const current = formatFromAttributes(cell.attrs);
    const changesBorder =
      normalized.borderColor !== undefined ||
      normalized.borderStyle !== undefined ||
      normalized.borderWidth !== undefined;
    const nextBorder: DocumentTableBorder = {
      color: normalized.borderColor ?? current.borderColor,
      style: normalized.borderStyle ?? current.borderStyle,
      width: normalized.borderWidth ?? current.borderWidth,
    };
    if (nextBorder.style === 'none') nextBorder.width = 0;
    transaction.setNodeMarkup(position, undefined, {
      ...cell.attrs,
      ...normalized,
      ...(changesBorder
        ? {
            borderColor: nextBorder.color,
            borderStyle: nextBorder.style,
            borderWidth: nextBorder.width,
            borders: uniformDocumentTableBorders(nextBorder),
          }
        : {}),
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
      const format = tableStyleCellFormat(style, cell, rowIndex);
      const border = {
        color: style.borderColor,
        style: style.borderStyle,
        width: style.borderWidth,
      } satisfies DocumentTableBorder;
      transaction.setNodeMarkup(position, undefined, {
        ...cell.attrs,
        ...format,
        borders: uniformDocumentTableBorders(border),
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
  if (
    selection instanceof NodeSelection &&
    selection.node.type.spec.tableRole === 'table'
  ) {
    const positions: number[] = [];
    selection.node.forEach((row, rowOffset) => {
      row.forEach((_cell, cellOffset) => {
        positions.push(selection.from + 2 + rowOffset + cellOffset);
      });
    });
    return positions;
  }
  const position = ancestorPosition(selection.$from, isTableCell);
  return position === null ? [] : [position];
}

function selectedTable(
  selection: Selection,
): { node: ProseMirrorNode; position: number } | null {
  if (
    selection instanceof NodeSelection &&
    selection.node.type.spec.tableRole === 'table'
  ) {
    return { node: selection.node, position: selection.from };
  }
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
  if (patch.margins !== undefined) {
    if (patch.margins === null) {
      normalized.margins = null;
    } else {
      const margins = normalizeDocumentTableCellMarginOverrides(patch.margins);
      if (!margins) return null;
      normalized.margins = margins;
    }
  }
  return normalized;
}

function formatFromAttributes(
  attributes: Record<string, unknown>,
  defaults = DEFAULT_DOCUMENT_TABLE_CELL_FORMAT,
): DocumentTableCellFormat {
  const fallbackBorder: DocumentTableBorder = {
    color:
      normalizeTableColor(String(attributes.borderColor ?? '')) ??
      defaults.borderColor,
    style:
      normalizeDocumentTableBorderStyle(String(attributes.borderStyle ?? '')) ??
      defaults.borderStyle,
    width:
      normalizeDocumentTableBorderWidth(attributes.borderWidth) ??
      defaults.borderWidth,
  };
  if (fallbackBorder.style === 'none') fallbackBorder.width = 0;
  const borders = documentTableBordersFromAttributes(
    attributes,
    fallbackBorder,
  );
  const representative = uniformDocumentTableBorder(borders) ?? borders.top;
  return {
    backgroundColor:
      normalizeTableColor(String(attributes.backgroundColor ?? '')) ??
      defaults.backgroundColor,
    verticalAlign:
      normalizeDocumentTableVerticalAlign(
        String(attributes.verticalAlign ?? ''),
      ) ?? defaults.verticalAlign,
    borderColor: representative.color,
    borderStyle: representative.style,
    borderWidth: representative.width,
    borders,
    margins: normalizeDocumentTableCellMarginOverrides(attributes.margins),
  };
}

export function normalizeDocumentTableVerticalAlign(
  value: string | null | undefined,
): DocumentTableVerticalAlign | null {
  if (value === 'top' || value === 'middle' || value === 'bottom') return value;
  if (value === 'center') return 'middle';
  return null;
}
