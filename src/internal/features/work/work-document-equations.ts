import { type CommandProps, mergeAttributes, Node } from '@tiptap/core';
import type { DOMOutputSpec } from '@tiptap/pm/model';

export type WorkDocumentEquationDisplay = 'inline' | 'block';
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
      type: 'matrix';
      baseAlignment: WorkDocumentEquationMatrixBaseAlignment;
      placeholdersHidden: boolean;
      columnAlignments: WorkDocumentEquationMatrixAlignment[];
      rows: WorkDocumentEquationExpression[][][];
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
): string {
  return expressions.map(expressionText).join('');
}

function expressionText(expression: WorkDocumentEquationExpression): string {
  if (expression.type === 'run') return expression.text;
  if (expression.type === 'fraction') {
    return `(${expressionListText(expression.numerator)})/(${expressionListText(
      expression.denominator,
    )})`;
  }
  if (expression.type === 'superscript') {
    return `${expressionListText(expression.base)}^(${expressionListText(
      expression.superScript,
    )})`;
  }
  if (expression.type === 'subscript') {
    return `${expressionListText(expression.base)}_(${expressionListText(
      expression.subScript,
    )})`;
  }
  if (expression.type === 'subSuperScript') {
    return `${expressionListText(expression.base)}_(${expressionListText(
      expression.subScript,
    )})^(${expressionListText(expression.superScript)})`;
  }
  if (expression.type === 'radical') {
    const body = expressionListText(expression.children);
    return expression.degree
      ? `root(${expressionListText(expression.degree)};${body})`
      : `sqrt(${body})`;
  }
  if (expression.type === 'function') {
    return `${expressionListText(expression.name)}(${expressionListText(
      expression.children,
    )})`;
  }
  if (expression.type === 'nary') {
    return `${expression.operator}${
      expression.subScript
        ? `_(${expressionListText(expression.subScript)})`
        : ''
    }${
      expression.superScript
        ? `^(${expressionListText(expression.superScript)})`
        : ''
    } ${expressionListText(expression.children)}`;
  }
  if (expression.type === 'matrix') {
    return `matrix(${expression.rows
      .map((row) => row.map(expressionListText).join(','))
      .join(';')})`;
  }
  return `${expression.opening}${expression.arguments
    .map(expressionListText)
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
): DOMOutputSpec {
  return domSpec('mrow', {}, expressions.map(expressionMathMl));
}

function expressionMathMl(
  expression: WorkDocumentEquationExpression,
): DOMOutputSpec {
  if (expression.type === 'run') return domSpec('mtext', {}, [expression.text]);
  if (expression.type === 'fraction') {
    if (expression.fractionType === 'linear') {
      return domSpec('mrow', {}, [
        mathRow(expression.numerator),
        domSpec('mo', {}, ['/']),
        mathRow(expression.denominator),
      ]);
    }
    return domSpec(
      'mfrac',
      {
        ...(expression.fractionType === 'noBar' ? { linethickness: '0' } : {}),
        ...(expression.fractionType === 'skewed' ? { bevelled: 'true' } : {}),
      },
      [mathRow(expression.numerator), mathRow(expression.denominator)],
    );
  }
  if (expression.type === 'superscript') {
    return domSpec('msup', {}, [
      mathRow(expression.base),
      mathRow(expression.superScript),
    ]);
  }
  if (expression.type === 'subscript') {
    return domSpec('msub', {}, [
      mathRow(expression.base),
      mathRow(expression.subScript),
    ]);
  }
  if (expression.type === 'subSuperScript') {
    return domSpec('msubsup', {}, [
      mathRow(expression.base),
      mathRow(expression.subScript),
      mathRow(expression.superScript),
    ]);
  }
  if (expression.type === 'radical') {
    return expression.degree
      ? domSpec('mroot', {}, [
          mathRow(expression.children),
          mathRow(expression.degree),
        ])
      : domSpec('msqrt', {}, [mathRow(expression.children)]);
  }
  if (expression.type === 'function') {
    return domSpec('mrow', {}, [
      mathRow(expression.name),
      domSpec('mo', {}, ['\u2061']),
      mathRow(expression.children),
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
          mathRow(expression.subScript),
          mathRow(expression.superScript),
        ],
      );
    } else if (expression.subScript) {
      decorated = domSpec(
        expression.limitLocation === 'underOver' ? 'munder' : 'msub',
        {},
        [operator, mathRow(expression.subScript)],
      );
    } else if (expression.superScript) {
      decorated = domSpec(
        expression.limitLocation === 'underOver' ? 'mover' : 'msup',
        {},
        [operator, mathRow(expression.superScript)],
      );
    }
    return domSpec('mrow', {}, [decorated, mathRow(expression.children)]);
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
          row.map((cell) => domSpec('mtd', {}, [mathRow(cell)])),
        ),
      ),
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
    children.push(mathRow(argument));
  });
  if (expression.closing)
    children.push(domSpec('mo', { fence: 'true' }, [expression.closing]));
  return domSpec('mrow', {}, children);
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
