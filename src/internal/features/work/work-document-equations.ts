import { type CommandProps, mergeAttributes, Node } from '@tiptap/core';
import type { DOMOutputSpec } from '@tiptap/pm/model';

export type WorkDocumentEquationDisplay = 'inline' | 'block';
export type WorkDocumentEquationBarPosition = 'top' | 'bottom';
export type WorkDocumentEquationFractionType =
  | 'bar'
  | 'noBar'
  | 'skewed'
  | 'linear';
export type WorkDocumentEquationLimitLocation = 'underOver' | 'subSup';
export type WorkDocumentEquationMatrixAlignment = 'left' | 'center' | 'right';
export type WorkDocumentEquationMatrixBaseAlignment =
  | 'top'
  | 'center'
  | 'bottom';
export type WorkDocumentEquationRowSpacingRule =
  | 'single'
  | 'oneAndHalf'
  | 'double'
  | 'exact'
  | 'multiple';

export interface WorkDocumentEquationManualBreak {
  alignmentAt?: number;
}

export type WorkDocumentEquationExpression =
  | { type: 'run'; text: string }
  | {
      type: 'fraction';
      fractionType: WorkDocumentEquationFractionType;
      numerator: WorkDocumentEquationExpression[];
      denominator: WorkDocumentEquationExpression[];
    }
  | {
      type: 'superscript';
      base: WorkDocumentEquationExpression[];
      superScript: WorkDocumentEquationExpression[];
    }
  | {
      type: 'subscript';
      base: WorkDocumentEquationExpression[];
      subScript: WorkDocumentEquationExpression[];
    }
  | {
      type: 'subSuperScript';
      base: WorkDocumentEquationExpression[];
      subScript: WorkDocumentEquationExpression[];
      superScript: WorkDocumentEquationExpression[];
    }
  | {
      type: 'lowerLimit';
      base: WorkDocumentEquationExpression[];
      limit: WorkDocumentEquationExpression[];
    }
  | {
      type: 'upperLimit';
      base: WorkDocumentEquationExpression[];
      limit: WorkDocumentEquationExpression[];
    }
  | {
      type: 'radical';
      children: WorkDocumentEquationExpression[];
      degree?: WorkDocumentEquationExpression[];
    }
  | {
      type: 'function';
      name: WorkDocumentEquationExpression[];
      children: WorkDocumentEquationExpression[];
    }
  | {
      type: 'nary';
      operator: WorkDocumentEquationNaryOperator;
      limitLocation: WorkDocumentEquationLimitLocation;
      children: WorkDocumentEquationExpression[];
      subScript?: WorkDocumentEquationExpression[];
      superScript?: WorkDocumentEquationExpression[];
    }
  | {
      type: 'accent';
      character: string;
      children: WorkDocumentEquationExpression[];
    }
  | {
      type: 'bar';
      position: WorkDocumentEquationBarPosition;
      children: WorkDocumentEquationExpression[];
    }
  | {
      type: 'groupCharacter';
      character: string;
      position: WorkDocumentEquationBarPosition;
      verticalJustification: WorkDocumentEquationBarPosition;
      children: WorkDocumentEquationExpression[];
    }
  | {
      type: 'borderBox';
      hideTop: boolean;
      hideBottom: boolean;
      hideLeft: boolean;
      hideRight: boolean;
      strikeHorizontal: boolean;
      strikeVertical: boolean;
      strikeBottomLeftToTopRight: boolean;
      strikeTopLeftToBottomRight: boolean;
      children: WorkDocumentEquationExpression[];
    }
  | {
      type: 'box';
      operatorEmulator: boolean;
      noBreak: boolean;
      differential: boolean;
      alignment: boolean;
      manualBreak?: WorkDocumentEquationManualBreak;
      children: WorkDocumentEquationExpression[];
    }
  | {
      type: 'matrix';
      baseAlignment: WorkDocumentEquationMatrixBaseAlignment;
      placeholdersHidden: boolean;
      columnAlignments: WorkDocumentEquationMatrixAlignment[];
      rows: WorkDocumentEquationExpression[][][];
    }
  | {
      type: 'equationArray';
      baseAlignment: WorkDocumentEquationMatrixBaseAlignment;
      maximumDistribution: boolean;
      objectDistribution: boolean;
      rowSpacingRule: WorkDocumentEquationRowSpacingRule;
      rowSpacing: number;
      rows: WorkDocumentEquationExpression[][];
    }
  | {
      type: 'delimiter';
      opening: string;
      closing: string;
      separator: string;
      arguments: WorkDocumentEquationExpression[][];
    };

export type WorkDocumentEquationNaryOperator =
  | '∑'
  | '∏'
  | '∐'
  | '∫'
  | '∬'
  | '∭'
  | '∮'
  | '∯'
  | '∰'
  | '⋂'
  | '⋃';

export interface WorkDocumentEquation {
  version: 1;
  display: WorkDocumentEquationDisplay;
  children: WorkDocumentEquationExpression[];
}

interface EquationNormalizationState {
  depth: number;
  nodes: number;
  textLength: number;
  equationArrayDepth: number;
  equationArrayAlignmentMarkers: number;
}

interface EquationArrayAlignmentState {
  markerIndex: number;
  started: boolean;
}

const EQUATION_SELECTOR = 'span[data-document-equation]';
const MATHML_NAMESPACE = 'http://www.w3.org/1998/Math/MathML';
const MAX_EQUATION_DEPTH = 32;
const MAX_EQUATION_NODES = 4_096;
const MAX_EQUATION_TEXT_LENGTH = 65_536;
const MAX_EQUATION_MODEL_LENGTH = 262_144;
const MAX_DELIMITER_ARGUMENTS = 32;
const MAX_MATRIX_ROWS = 64;
const MAX_MATRIX_COLUMNS = 64;
const MAX_MATRIX_CELLS = 1_024;
const MAX_EQUATION_ARRAY_ROWS = 64;
const MAX_EQUATION_ARRAY_ALIGNMENT_MARKERS = 4_096;
const NARY_OPERATORS = new Set<WorkDocumentEquationNaryOperator>([
  '∑',
  '∏',
  '∐',
  '∫',
  '∬',
  '∭',
  '∮',
  '∯',
  '∰',
  '⋂',
  '⋃',
]);
const FRACTION_TYPES = new Set<WorkDocumentEquationFractionType>([
  'bar',
  'noBar',
  'skewed',
  'linear',
]);
const LIMIT_LOCATIONS = new Set<WorkDocumentEquationLimitLocation>([
  'underOver',
  'subSup',
]);
const MATRIX_ALIGNMENTS = new Set<WorkDocumentEquationMatrixAlignment>([
  'left',
  'center',
  'right',
]);
const MATRIX_BASE_ALIGNMENTS = new Set<WorkDocumentEquationMatrixBaseAlignment>(
  ['top', 'center', 'bottom'],
);
const EQUATION_ARRAY_ROW_SPACING_RULES =
  new Set<WorkDocumentEquationRowSpacingRule>([
    'single',
    'oneAndHalf',
    'double',
    'exact',
    'multiple',
  ]);
const BAR_POSITIONS = new Set<WorkDocumentEquationBarPosition>([
  'top',
  'bottom',
]);
const MATHML_ACCENT_CHARACTERS = new Map([
  ['\u0300', '`'],
  ['\u0301', '\u00b4'],
  ['\u0302', '\u02c6'],
  ['\u0303', '\u02dc'],
  ['\u0304', '\u00af'],
  ['\u0305', '\u203e'],
  ['\u0306', '\u02d8'],
  ['\u0307', '\u02d9'],
  ['\u0308', '\u00a8'],
  ['\u030a', '\u02da'],
  ['\u030b', '\u02dd'],
  ['\u030c', '\u02c7'],
  ['\u20d6', '\u2190'],
  ['\u20d7', '\u2192'],
  ['\u20e1', '\u2194'],
]);

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentEquation: {
      insertDocumentEquation: (equation: WorkDocumentEquation) => ReturnType;
      updateDocumentEquation: (equation: WorkDocumentEquation) => ReturnType;
    };
  }
}

export const DocumentEquation = Node.create({
  name: 'documentEquation',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: true,

  addCommands() {
    return {
      insertDocumentEquation: (equation) => (props) =>
        insertDocumentEquationCommand(props, equation),
      updateDocumentEquation: (equation) => (props) =>
        updateDocumentEquationCommand(props, equation),
    };
  },

  addAttributes() {
    return {
      equation: {
        default: defaultDocumentEquation(),
        rendered: false,
      },
    };
  },

  parseHTML() {
    return [
      {
        tag: EQUATION_SELECTOR,
        priority: 120,
        getAttrs: (node) => {
          if (!(node instanceof HTMLElement)) return false;
          const equation = documentEquationFromElement(node);
          return equation ? { equation } : false;
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const equation =
      normalizeDocumentEquation(node.attrs.equation) ??
      defaultDocumentEquation();
    const label = documentEquationText(equation);
    return [
      'span',
      mergeAttributes(HTMLAttributes, {
        'data-document-equation': 'true',
        'data-equation-display': equation.display,
        'data-equation-model': serializeDocumentEquation(equation),
        'aria-label': label,
        class: `work-document-equation ${equation.display}`,
        contenteditable: 'false',
        role: 'math',
      }),
      equationMathMl(equation),
    ];
  },

  renderText({ node }) {
    return documentEquationText(
      normalizeDocumentEquation(node.attrs.equation) ??
        defaultDocumentEquation(),
    );
  },
});

export function normalizeDocumentEquation(
  source: unknown,
): WorkDocumentEquation | null {
  if (!isRecord(source) || source.version !== 1) return null;
  const display =
    source.display === 'block' || source.display === 'inline'
      ? source.display
      : null;
  if (!display) return null;
  const state: EquationNormalizationState = {
    depth: 0,
    nodes: 0,
    textLength: 0,
    equationArrayDepth: 0,
    equationArrayAlignmentMarkers: 0,
  };
  const children = normalizeExpressionList(source.children, state);
  if (!children) return null;
  const equation = {
    version: 1,
    display,
    children,
  } satisfies WorkDocumentEquation;
  return JSON.stringify(equation).length <= MAX_EQUATION_MODEL_LENGTH
    ? equation
    : null;
}

export function serializeDocumentEquation(
  source: WorkDocumentEquation,
): string {
  const equation = normalizeDocumentEquation(source);
  if (!equation) throw new Error('The document equation model is invalid.');
  const serialized = JSON.stringify(equation);
  if (serialized.length > MAX_EQUATION_MODEL_LENGTH) {
    throw new Error('The document equation model exceeds the size limit.');
  }
  return serialized;
}

export function parseDocumentEquation(
  source: string | undefined,
): WorkDocumentEquation | null {
  if (!source || source.length > MAX_EQUATION_MODEL_LENGTH) return null;
  try {
    return normalizeDocumentEquation(JSON.parse(source));
  } catch {
    return null;
  }
}

export function documentEquationFromElement(
  element: Element,
): WorkDocumentEquation | null {
  if (!(element instanceof HTMLElement)) return null;
  return parseDocumentEquation(element.dataset.equationModel);
}

export function createDocumentEquationElement(
  document: Document,
  source: WorkDocumentEquation,
): HTMLElement {
  const equation = normalizeDocumentEquation(source);
  if (!equation) throw new Error('The document equation model is invalid.');
  const element = document.createElement('span');
  const label = documentEquationText(equation);
  element.dataset.documentEquation = 'true';
  element.dataset.equationDisplay = equation.display;
  element.dataset.equationModel = serializeDocumentEquation(equation);
  element.className = `work-document-equation ${equation.display}`;
  element.contentEditable = 'false';
  element.setAttribute('role', 'math');
  element.setAttribute('aria-label', label);
  element.append(
    createEquationDom(element.ownerDocument, equationMathMl(equation)),
  );
  return element;
}

export function documentEquationText(source: WorkDocumentEquation): string {
  return expressionListText(source.children).trim() || 'Equation';
}

export function defaultDocumentEquation(): WorkDocumentEquation {
  return {
    version: 1,
    display: 'inline',
    children: [{ type: 'run', text: 'x' }],
  };
}

function insertDocumentEquationCommand(
  { dispatch, editor, tr }: CommandProps,
  source: WorkDocumentEquation,
): boolean {
  const equation = normalizeDocumentEquation(source);
  const equationType = editor.schema.nodes.documentEquation;
  if (!equation || !equationType) return false;
  if (!dispatch) return true;
  tr.replaceSelectionWith(equationType.create({ equation }), false);
  tr.scrollIntoView();
  return true;
}

function updateDocumentEquationCommand(
  { dispatch, editor, state, tr }: CommandProps,
  source: WorkDocumentEquation,
): boolean {
  const equation = normalizeDocumentEquation(source);
  const equationType = editor.schema.nodes.documentEquation;
  const position = state.selection.from;
  const selected = state.doc.nodeAt(position);
  if (!equation || !equationType || selected?.type !== equationType) {
    return false;
  }
  if (!dispatch) return true;
  tr.setNodeMarkup(position, equationType, { equation });
  tr.scrollIntoView();
  return true;
}

function normalizeExpressionList(
  source: unknown,
  state: EquationNormalizationState,
  optional = false,
): WorkDocumentEquationExpression[] | null {
  if (!Array.isArray(source) || source.length > MAX_EQUATION_NODES) return null;
  if (!optional && source.length === 0) return null;
  const children: WorkDocumentEquationExpression[] = [];
  for (const child of source) {
    const normalized = normalizeExpression(child, state);
    if (!normalized) return null;
    children.push(normalized);
  }
  return children;
}

function normalizeExpression(
  source: unknown,
  state: EquationNormalizationState,
): WorkDocumentEquationExpression | null {
  if (
    !isRecord(source) ||
    state.depth >= MAX_EQUATION_DEPTH ||
    state.nodes >= MAX_EQUATION_NODES
  ) {
    return null;
  }
  state.nodes += 1;
  state.depth += 1;
  try {
    if (source.type === 'run') {
      if (
        typeof source.text !== 'string' ||
        source.text.length === 0 ||
        !validXmlText(source.text)
      )
        return null;
      state.textLength += source.text.length;
      if (state.textLength > MAX_EQUATION_TEXT_LENGTH) return null;
      if (state.equationArrayDepth > 0) {
        state.equationArrayAlignmentMarkers +=
          source.text.split('&').length - 1;
        if (
          state.equationArrayAlignmentMarkers >
          MAX_EQUATION_ARRAY_ALIGNMENT_MARKERS
        ) {
          return null;
        }
      }
      return { type: 'run', text: source.text };
    }
    if (source.type === 'fraction') {
      const fractionType = FRACTION_TYPES.has(
        source.fractionType as WorkDocumentEquationFractionType,
      )
        ? (source.fractionType as WorkDocumentEquationFractionType)
        : null;
      const numerator = normalizeExpressionList(source.numerator, state);
      const denominator = normalizeExpressionList(source.denominator, state);
      return fractionType && numerator && denominator
        ? { type: 'fraction', fractionType, numerator, denominator }
        : null;
    }
    if (source.type === 'superscript') {
      const base = normalizeExpressionList(source.base, state);
      const superScript = normalizeExpressionList(source.superScript, state);
      return base && superScript
        ? { type: 'superscript', base, superScript }
        : null;
    }
    if (source.type === 'subscript') {
      const base = normalizeExpressionList(source.base, state);
      const subScript = normalizeExpressionList(source.subScript, state);
      return base && subScript ? { type: 'subscript', base, subScript } : null;
    }
    if (source.type === 'subSuperScript') {
      const base = normalizeExpressionList(source.base, state);
      const subScript = normalizeExpressionList(source.subScript, state);
      const superScript = normalizeExpressionList(source.superScript, state);
      return base && subScript && superScript
        ? { type: 'subSuperScript', base, subScript, superScript }
        : null;
    }
    if (source.type === 'lowerLimit' || source.type === 'upperLimit') {
      const base = normalizeExpressionList(source.base, state);
      const limit = normalizeExpressionList(source.limit, state);
      return base && limit ? { type: source.type, base, limit } : null;
    }
    if (source.type === 'radical') {
      const children = normalizeExpressionList(source.children, state);
      const degree =
        source.degree === undefined
          ? undefined
          : normalizeExpressionList(source.degree, state);
      return children && degree !== null
        ? {
            type: 'radical',
            children,
            ...(degree ? { degree } : {}),
          }
        : null;
    }
    if (source.type === 'function') {
      const name = normalizeExpressionList(source.name, state);
      const children = normalizeExpressionList(source.children, state);
      return name && children ? { type: 'function', name, children } : null;
    }
    if (source.type === 'nary') {
      const operator = NARY_OPERATORS.has(
        source.operator as WorkDocumentEquationNaryOperator,
      )
        ? (source.operator as WorkDocumentEquationNaryOperator)
        : null;
      const limitLocation = LIMIT_LOCATIONS.has(
        source.limitLocation as WorkDocumentEquationLimitLocation,
      )
        ? (source.limitLocation as WorkDocumentEquationLimitLocation)
        : null;
      const children = normalizeExpressionList(source.children, state);
      const subScript =
        source.subScript === undefined
          ? undefined
          : normalizeExpressionList(source.subScript, state);
      const superScript =
        source.superScript === undefined
          ? undefined
          : normalizeExpressionList(source.superScript, state);
      return operator &&
        limitLocation &&
        children &&
        subScript !== null &&
        superScript !== null
        ? {
            type: 'nary',
            operator,
            limitLocation,
            children,
            ...(subScript ? { subScript } : {}),
            ...(superScript ? { superScript } : {}),
          }
        : null;
    }
    if (source.type === 'accent') {
      const character = accentCharacter(source.character);
      const children = normalizeExpressionList(source.children, state);
      return character && children
        ? { type: 'accent', character, children }
        : null;
    }
    if (source.type === 'bar') {
      const position = BAR_POSITIONS.has(
        source.position as WorkDocumentEquationBarPosition,
      )
        ? (source.position as WorkDocumentEquationBarPosition)
        : null;
      const children = normalizeExpressionList(source.children, state);
      return position && children ? { type: 'bar', position, children } : null;
    }
    if (source.type === 'groupCharacter') {
      const character = mathCharacter(source.character);
      const position = BAR_POSITIONS.has(
        source.position as WorkDocumentEquationBarPosition,
      )
        ? (source.position as WorkDocumentEquationBarPosition)
        : null;
      const verticalJustification = BAR_POSITIONS.has(
        source.verticalJustification as WorkDocumentEquationBarPosition,
      )
        ? (source.verticalJustification as WorkDocumentEquationBarPosition)
        : null;
      const children = normalizeExpressionList(source.children, state);
      return character !== null && position && verticalJustification && children
        ? {
            type: 'groupCharacter',
            character,
            position,
            verticalJustification,
            children,
          }
        : null;
    }
    if (source.type === 'borderBox') {
      if (
        typeof source.hideTop !== 'boolean' ||
        typeof source.hideBottom !== 'boolean' ||
        typeof source.hideLeft !== 'boolean' ||
        typeof source.hideRight !== 'boolean' ||
        typeof source.strikeHorizontal !== 'boolean' ||
        typeof source.strikeVertical !== 'boolean' ||
        typeof source.strikeBottomLeftToTopRight !== 'boolean' ||
        typeof source.strikeTopLeftToBottomRight !== 'boolean'
      ) {
        return null;
      }
      const children = normalizeExpressionList(source.children, state);
      return children
        ? {
            type: 'borderBox',
            hideTop: source.hideTop,
            hideBottom: source.hideBottom,
            hideLeft: source.hideLeft,
            hideRight: source.hideRight,
            strikeHorizontal: source.strikeHorizontal,
            strikeVertical: source.strikeVertical,
            strikeBottomLeftToTopRight: source.strikeBottomLeftToTopRight,
            strikeTopLeftToBottomRight: source.strikeTopLeftToBottomRight,
            children,
          }
        : null;
    }
    if (source.type === 'box') {
      const manualBreak =
        source.manualBreak === undefined
          ? undefined
          : normalizeManualBreak(source.manualBreak);
      const children = normalizeExpressionList(source.children, state);
      return typeof source.operatorEmulator === 'boolean' &&
        typeof source.noBreak === 'boolean' &&
        typeof source.differential === 'boolean' &&
        typeof source.alignment === 'boolean' &&
        manualBreak !== null &&
        children
        ? {
            type: 'box',
            operatorEmulator: source.operatorEmulator,
            noBreak: source.noBreak,
            differential: source.differential,
            alignment: source.alignment,
            ...(manualBreak ? { manualBreak } : {}),
            children,
          }
        : null;
    }
    if (source.type === 'matrix') {
      const baseAlignment = MATRIX_BASE_ALIGNMENTS.has(
        source.baseAlignment as WorkDocumentEquationMatrixBaseAlignment,
      )
        ? (source.baseAlignment as WorkDocumentEquationMatrixBaseAlignment)
        : null;
      if (
        !baseAlignment ||
        typeof source.placeholdersHidden !== 'boolean' ||
        !Array.isArray(source.rows) ||
        source.rows.length === 0 ||
        source.rows.length > MAX_MATRIX_ROWS ||
        !Array.isArray(source.columnAlignments)
      ) {
        return null;
      }
      const columnCount = Array.isArray(source.rows[0])
        ? source.rows[0].length
        : 0;
      if (
        columnCount === 0 ||
        columnCount > MAX_MATRIX_COLUMNS ||
        source.rows.length * columnCount > MAX_MATRIX_CELLS ||
        source.columnAlignments.length !== columnCount
      ) {
        return null;
      }
      const columnAlignments = source.columnAlignments.map((alignment) =>
        MATRIX_ALIGNMENTS.has(alignment as WorkDocumentEquationMatrixAlignment)
          ? (alignment as WorkDocumentEquationMatrixAlignment)
          : null,
      );
      if (
        !columnAlignments.every(
          (alignment): alignment is WorkDocumentEquationMatrixAlignment =>
            alignment !== null,
        )
      ) {
        return null;
      }
      const rows: WorkDocumentEquationExpression[][][] = [];
      for (const row of source.rows) {
        if (!Array.isArray(row) || row.length !== columnCount) return null;
        const normalizedRow: WorkDocumentEquationExpression[][] = [];
        for (const cell of row) {
          const normalizedCell = normalizeExpressionList(cell, state, true);
          if (!normalizedCell) return null;
          normalizedRow.push(normalizedCell);
        }
        rows.push(normalizedRow);
      }
      return {
        type: 'matrix',
        baseAlignment,
        placeholdersHidden: source.placeholdersHidden,
        columnAlignments,
        rows,
      };
    }
    if (source.type === 'equationArray') {
      const baseAlignment = MATRIX_BASE_ALIGNMENTS.has(
        source.baseAlignment as WorkDocumentEquationMatrixBaseAlignment,
      )
        ? (source.baseAlignment as WorkDocumentEquationMatrixBaseAlignment)
        : null;
      const rowSpacingRule = EQUATION_ARRAY_ROW_SPACING_RULES.has(
        source.rowSpacingRule as WorkDocumentEquationRowSpacingRule,
      )
        ? (source.rowSpacingRule as WorkDocumentEquationRowSpacingRule)
        : null;
      if (
        !baseAlignment ||
        typeof source.maximumDistribution !== 'boolean' ||
        typeof source.objectDistribution !== 'boolean' ||
        !rowSpacingRule ||
        typeof source.rowSpacing !== 'number' ||
        !Number.isInteger(source.rowSpacing) ||
        source.rowSpacing < 0 ||
        source.rowSpacing > 65_535 ||
        !Array.isArray(source.rows) ||
        source.rows.length === 0 ||
        source.rows.length > MAX_EQUATION_ARRAY_ROWS
      ) {
        return null;
      }
      const rows: WorkDocumentEquationExpression[][] = [];
      state.equationArrayDepth += 1;
      try {
        for (const row of source.rows) {
          const normalizedRow = normalizeExpressionList(row, state, true);
          if (!normalizedRow) return null;
          rows.push(normalizedRow);
        }
      } finally {
        state.equationArrayDepth -= 1;
      }
      return {
        type: 'equationArray',
        baseAlignment,
        maximumDistribution: source.maximumDistribution,
        objectDistribution: source.objectDistribution,
        rowSpacingRule,
        rowSpacing: source.rowSpacing,
        rows,
      };
    }
    if (source.type === 'delimiter') {
      const opening = mathCharacter(source.opening);
      const closing = mathCharacter(source.closing);
      const separator = mathCharacter(source.separator);
      if (
        opening === null ||
        closing === null ||
        separator === null ||
        !Array.isArray(source.arguments) ||
        source.arguments.length === 0 ||
        source.arguments.length > MAX_DELIMITER_ARGUMENTS
      ) {
        return null;
      }
      const arguments_ = source.arguments.map((argument) =>
        normalizeExpressionList(argument, state),
      );
      return arguments_.every(
        (argument): argument is WorkDocumentEquationExpression[] =>
          Boolean(argument),
      )
        ? {
            type: 'delimiter',
            opening,
            closing,
            separator,
            arguments: arguments_,
          }
        : null;
    }
    return null;
  } finally {
    state.depth -= 1;
  }
}

function expressionListText(
  expressions: readonly WorkDocumentEquationExpression[],
  hideAlignmentMarkers = false,
): string {
  return expressions
    .map((expression) => expressionText(expression, hideAlignmentMarkers))
    .join('');
}

function expressionText(
  expression: WorkDocumentEquationExpression,
  hideAlignmentMarkers = false,
): string {
  if (expression.type === 'run') {
    return hideAlignmentMarkers
      ? expression.text.replaceAll('&', '')
      : expression.text;
  }
  if (expression.type === 'fraction') {
    return `(${expressionListText(
      expression.numerator,
      hideAlignmentMarkers,
    )})/(${expressionListText(expression.denominator, hideAlignmentMarkers)})`;
  }
  if (expression.type === 'superscript') {
    return `${expressionListText(
      expression.base,
      hideAlignmentMarkers,
    )}^(${expressionListText(expression.superScript, hideAlignmentMarkers)})`;
  }
  if (expression.type === 'subscript') {
    return `${expressionListText(
      expression.base,
      hideAlignmentMarkers,
    )}_(${expressionListText(expression.subScript, hideAlignmentMarkers)})`;
  }
  if (expression.type === 'subSuperScript') {
    return `${expressionListText(
      expression.base,
      hideAlignmentMarkers,
    )}_(${expressionListText(
      expression.subScript,
      hideAlignmentMarkers,
    )})^(${expressionListText(expression.superScript, hideAlignmentMarkers)})`;
  }
  if (expression.type === 'radical') {
    const body = expressionListText(expression.children, hideAlignmentMarkers);
    return expression.degree
      ? `root(${expressionListText(
          expression.degree,
          hideAlignmentMarkers,
        )};${body})`
      : `sqrt(${body})`;
  }
  if (expression.type === 'function') {
    return `${expressionListText(
      expression.name,
      hideAlignmentMarkers,
    )}(${expressionListText(expression.children, hideAlignmentMarkers)})`;
  }
  if (expression.type === 'nary') {
    return `${expression.operator}${
      expression.subScript
        ? `_(${expressionListText(expression.subScript, hideAlignmentMarkers)})`
        : ''
    }${
      expression.superScript
        ? `^(${expressionListText(
            expression.superScript,
            hideAlignmentMarkers,
          )})`
        : ''
    } ${expressionListText(expression.children, hideAlignmentMarkers)}`;
  }
  if (expression.type === 'accent') {
    const codePoint = expression.character.codePointAt(0);
    const label = codePoint?.toString(16).toUpperCase().padStart(4, '0');
    return `accent(U+${label};${expressionListText(
      expression.children,
      hideAlignmentMarkers,
    )})`;
  }
  if (expression.type === 'lowerLimit' || expression.type === 'upperLimit') {
    return `${expression.type === 'lowerLimit' ? 'lower-limit' : 'upper-limit'}(${expressionListText(
      expression.base,
      hideAlignmentMarkers,
    )};${expressionListText(expression.limit, hideAlignmentMarkers)})`;
  }
  if (expression.type === 'bar') {
    const body = expressionListText(expression.children, hideAlignmentMarkers);
    return expression.position === 'top'
      ? `overbar(${body})`
      : `underbar(${body})`;
  }
  if (expression.type === 'groupCharacter') {
    const codePoint = expression.character.codePointAt(0);
    const character =
      codePoint === undefined
        ? 'none'
        : `U+${codePoint.toString(16).toUpperCase().padStart(4, '0')}`;
    return `group-character(${character};position=${expression.position};baseline=${expression.verticalJustification};${expressionListText(
      expression.children,
      hideAlignmentMarkers,
    )})`;
  }
  if (expression.type === 'borderBox') {
    return `borderbox(${borderBoxNotation(expression)};${expressionListText(
      expression.children,
      hideAlignmentMarkers,
    )})`;
  }
  if (expression.type === 'box') {
    const properties = [
      expression.operatorEmulator ? 'operator' : '',
      expression.noBreak ? 'no-break' : '',
      expression.differential ? 'differential' : '',
      expression.manualBreak
        ? expression.manualBreak.alignmentAt
          ? `break@${expression.manualBreak.alignmentAt}`
          : 'break'
        : '',
      expression.alignment ? 'alignment' : '',
    ].filter(Boolean);
    return `box(${properties.join(',') || 'default'};${expressionListText(
      expression.children,
      hideAlignmentMarkers,
    )})`;
  }
  if (expression.type === 'matrix') {
    return `matrix(${expression.rows
      .map((row) =>
        row
          .map((cell) => expressionListText(cell, hideAlignmentMarkers))
          .join(','),
      )
      .join(';')})`;
  }
  if (expression.type === 'equationArray') {
    const properties = [
      expression.baseAlignment,
      expression.maximumDistribution ? 'max-distribution' : '',
      expression.objectDistribution ? 'object-distribution' : '',
      `spacing=${expression.rowSpacingRule}:${expression.rowSpacing}`,
    ].filter(Boolean);
    return `equation-array(${properties.join(',')};${expression.rows
      .map((row) => expressionListText(row, true))
      .join(';')})`;
  }
  return `${expression.opening}${expression.arguments
    .map((argument) => expressionListText(argument, hideAlignmentMarkers))
    .join(expression.separator)}${expression.closing}`;
}

function equationMathMl(equation: WorkDocumentEquation): DOMOutputSpec {
  return domSpec(
    'math',
    {
      display: equation.display === 'block' ? 'block' : 'inline',
      xmlns: MATHML_NAMESPACE,
    },
    [mathRow(equation.children)],
  );
}

function createEquationDom(document: Document, spec: DOMOutputSpec): Element {
  if (!Array.isArray(spec) || typeof spec[0] !== 'string') {
    throw new Error('The document equation rendering model is invalid.');
  }
  const element = document.createElementNS(MATHML_NAMESPACE, spec[0]);
  const attributes = spec[1];
  if (
    attributes &&
    typeof attributes === 'object' &&
    !Array.isArray(attributes)
  ) {
    for (const [name, value] of Object.entries(attributes)) {
      if (typeof value === 'string') element.setAttribute(name, value);
    }
  }
  for (const child of spec.slice(2)) {
    if (typeof child === 'string') {
      element.append(document.createTextNode(child));
    } else if (Array.isArray(child) && typeof child[0] === 'string') {
      element.append(
        createEquationDom(document, child as unknown as DOMOutputSpec),
      );
    }
  }
  return element;
}

function mathRow(
  expressions: readonly WorkDocumentEquationExpression[],
  alignmentState?: EquationArrayAlignmentState,
): DOMOutputSpec {
  return domSpec(
    'mrow',
    {},
    expressions.map((expression) =>
      expressionMathMl(expression, alignmentState),
    ),
  );
}

function expressionMathMl(
  expression: WorkDocumentEquationExpression,
  alignmentState?: EquationArrayAlignmentState,
): DOMOutputSpec {
  if (expression.type === 'run') {
    if (!alignmentState) return domSpec('mtext', {}, [expression.text]);
    const children: DOMOutputSpec[] = [];
    if (!alignmentState.started) {
      children.push(domSpec('maligngroup', {}, []));
      alignmentState.started = true;
    }
    const parts = expression.text.split('&');
    parts.forEach((part, index) => {
      if (part) children.push(domSpec('mtext', {}, [part]));
      if (index === parts.length - 1) return;
      alignmentState.markerIndex += 1;
      children.push(
        domSpec(
          alignmentState.markerIndex % 2 === 1 ? 'malignmark' : 'maligngroup',
          {},
          [],
        ),
      );
    });
    return domSpec('mrow', {}, children);
  }
  if (expression.type === 'fraction') {
    if (expression.fractionType === 'linear') {
      return domSpec('mrow', {}, [
        mathRow(expression.numerator, alignmentState),
        domSpec('mo', {}, ['/']),
        mathRow(expression.denominator, alignmentState),
      ]);
    }
    return domSpec(
      'mfrac',
      {
        ...(expression.fractionType === 'noBar' ? { linethickness: '0' } : {}),
        ...(expression.fractionType === 'skewed' ? { bevelled: 'true' } : {}),
      },
      [
        mathRow(expression.numerator, alignmentState),
        mathRow(expression.denominator, alignmentState),
      ],
    );
  }
  if (expression.type === 'superscript') {
    return domSpec('msup', {}, [
      mathRow(expression.base, alignmentState),
      mathRow(expression.superScript, alignmentState),
    ]);
  }
  if (expression.type === 'subscript') {
    return domSpec('msub', {}, [
      mathRow(expression.base, alignmentState),
      mathRow(expression.subScript, alignmentState),
    ]);
  }
  if (expression.type === 'subSuperScript') {
    return domSpec('msubsup', {}, [
      mathRow(expression.base, alignmentState),
      mathRow(expression.subScript, alignmentState),
      mathRow(expression.superScript, alignmentState),
    ]);
  }
  if (expression.type === 'lowerLimit' || expression.type === 'upperLimit') {
    return domSpec(
      expression.type === 'lowerLimit' ? 'munder' : 'mover',
      expression.type === 'lowerLimit'
        ? { accentunder: 'false' }
        : { accent: 'false' },
      [
        mathRow(expression.base, alignmentState),
        mathRow(expression.limit, alignmentState),
      ],
    );
  }
  if (expression.type === 'radical') {
    return expression.degree
      ? domSpec('mroot', {}, [
          mathRow(expression.children, alignmentState),
          mathRow(expression.degree, alignmentState),
        ])
      : domSpec('msqrt', {}, [mathRow(expression.children, alignmentState)]);
  }
  if (expression.type === 'function') {
    return domSpec('mrow', {}, [
      mathRow(expression.name, alignmentState),
      domSpec('mo', {}, ['\u2061']),
      mathRow(expression.children, alignmentState),
    ]);
  }
  if (expression.type === 'nary') {
    const operator = domSpec('mo', {}, [expression.operator]);
    let decorated = operator;
    if (expression.subScript && expression.superScript) {
      decorated = domSpec(
        expression.limitLocation === 'underOver' ? 'munderover' : 'msubsup',
        {},
        [
          operator,
          mathRow(expression.subScript, alignmentState),
          mathRow(expression.superScript, alignmentState),
        ],
      );
    } else if (expression.subScript) {
      decorated = domSpec(
        expression.limitLocation === 'underOver' ? 'munder' : 'msub',
        {},
        [operator, mathRow(expression.subScript, alignmentState)],
      );
    } else if (expression.superScript) {
      decorated = domSpec(
        expression.limitLocation === 'underOver' ? 'mover' : 'msup',
        {},
        [operator, mathRow(expression.superScript, alignmentState)],
      );
    }
    return domSpec('mrow', {}, [
      decorated,
      mathRow(expression.children, alignmentState),
    ]);
  }
  if (expression.type === 'accent') {
    return domSpec('mover', { accent: 'true' }, [
      mathRow(expression.children, alignmentState),
      domSpec('mo', {}, [
        MATHML_ACCENT_CHARACTERS.get(expression.character) ??
          expression.character,
      ]),
    ]);
  }
  if (expression.type === 'bar') {
    return domSpec(
      expression.position === 'top' ? 'mover' : 'munder',
      expression.position === 'top'
        ? { accent: 'false' }
        : { accentunder: 'false' },
      [
        mathRow(expression.children, alignmentState),
        domSpec('mo', {}, ['\u00af']),
      ],
    );
  }
  if (expression.type === 'groupCharacter') {
    return domSpec(
      expression.position === 'top' ? 'mover' : 'munder',
      expression.position === 'top'
        ? { accent: 'false' }
        : { accentunder: 'false' },
      [
        mathRow(expression.children, alignmentState),
        domSpec('mo', {}, [expression.character]),
      ],
    );
  }
  if (expression.type === 'borderBox') {
    return domSpec('menclose', { notation: borderBoxNotation(expression) }, [
      mathRow(expression.children, alignmentState),
    ]);
  }
  if (expression.type === 'box') {
    return domSpec('mpadded', {}, [
      mathRow(expression.children, alignmentState),
    ]);
  }
  if (expression.type === 'matrix') {
    return domSpec(
      'mtable',
      {
        align: expression.baseAlignment,
        columnalign: expression.columnAlignments.join(' '),
      },
      expression.rows.map((row) =>
        domSpec(
          'mtr',
          {},
          row.map((cell) =>
            domSpec('mtd', {}, [mathRow(cell, alignmentState)]),
          ),
        ),
      ),
    );
  }
  if (expression.type === 'equationArray') {
    return domSpec(
      'mtable',
      {
        align: expression.baseAlignment,
        rowspacing: equationArrayRowSpacing(expression),
      },
      expression.rows.map((row) => {
        const rowAlignmentState: EquationArrayAlignmentState = {
          markerIndex: 0,
          started: false,
        };
        return domSpec('mtr', {}, [
          domSpec('mtd', {}, [mathRow(row, rowAlignmentState)]),
        ]);
      }),
    );
  }
  const children: Array<DOMOutputSpec | string> = [];
  if (expression.opening)
    children.push(domSpec('mo', { fence: 'true' }, [expression.opening]));
  expression.arguments.forEach((argument, index) => {
    if (index > 0 && expression.separator) {
      children.push(
        domSpec('mo', { separator: 'true' }, [expression.separator]),
      );
    }
    children.push(mathRow(argument, alignmentState));
  });
  if (expression.closing)
    children.push(domSpec('mo', { fence: 'true' }, [expression.closing]));
  return domSpec('mrow', {}, children);
}

function equationArrayRowSpacing(
  expression: Extract<
    WorkDocumentEquationExpression,
    { type: 'equationArray' }
  >,
): string {
  if (expression.rowSpacingRule === 'oneAndHalf') return '1.5em';
  if (expression.rowSpacingRule === 'double') return '2em';
  if (expression.rowSpacingRule === 'exact') {
    return `${expression.rowSpacing}pt`;
  }
  if (expression.rowSpacingRule === 'multiple') {
    return `${expression.rowSpacing / 2}em`;
  }
  return '1em';
}

function domSpec(
  name: string,
  attributes: Record<string, string>,
  children: readonly (DOMOutputSpec | string)[],
): DOMOutputSpec {
  return [name, attributes, ...children] as DOMOutputSpec;
}

function mathCharacter(source: unknown): string | null {
  if (typeof source !== 'string') return null;
  if (source === '') return '';
  const characters = Array.from(source);
  if (characters.length !== 1 || /[\p{Cc}\p{Cs}]/u.test(source)) return null;
  return source;
}

function accentCharacter(source: unknown): string | null {
  if (typeof source !== 'string' || Array.from(source).length !== 1) {
    return null;
  }
  const codePoint = source.codePointAt(0);
  return codePoint !== undefined &&
    ((codePoint >= 0x0300 && codePoint <= 0x036f) ||
      (codePoint >= 0x20d0 && codePoint <= 0x20ef))
    ? source
    : null;
}

function normalizeManualBreak(
  source: unknown,
): WorkDocumentEquationManualBreak | null {
  if (!isRecord(source)) return null;
  if (source.alignmentAt === undefined) return {};
  return Number.isInteger(source.alignmentAt) &&
    Number(source.alignmentAt) >= 1 &&
    Number(source.alignmentAt) <= 255
    ? { alignmentAt: Number(source.alignmentAt) }
    : null;
}

function borderBoxNotation(
  expression: Extract<WorkDocumentEquationExpression, { type: 'borderBox' }>,
): string {
  const edges = [
    expression.hideTop ? '' : 'top',
    expression.hideBottom ? '' : 'bottom',
    expression.hideLeft ? '' : 'left',
    expression.hideRight ? '' : 'right',
  ].filter(Boolean);
  const notations = edges.length === 4 ? ['box'] : edges;
  if (expression.strikeHorizontal) notations.push('horizontalstrike');
  if (expression.strikeVertical) notations.push('verticalstrike');
  if (expression.strikeBottomLeftToTopRight) {
    notations.push('updiagonalstrike');
  }
  if (expression.strikeTopLeftToBottomRight) {
    notations.push('downdiagonalstrike');
  }
  return notations.join(' ') || 'none';
}

function validXmlText(source: string): boolean {
  for (const character of Array.from(source)) {
    const codePoint = character.codePointAt(0) ?? 0;
    if (
      (codePoint < 0x20 &&
        codePoint !== 0x09 &&
        codePoint !== 0x0a &&
        codePoint !== 0x0d) ||
      (codePoint >= 0xd800 && codePoint <= 0xdfff) ||
      (codePoint & 0xffff) === 0xfffe ||
      (codePoint & 0xffff) === 0xffff
    ) {
      return false;
    }
  }
  return true;
}

function isRecord(source: unknown): source is Record<string, unknown> {
  return (
    Boolean(source) && typeof source === 'object' && !Array.isArray(source)
  );
}
