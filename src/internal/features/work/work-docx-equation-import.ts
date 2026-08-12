import {
  createDocumentEquationElement,
  documentEquationText,
  normalizeDocumentEquation,
  type WorkDocumentEquation,
  type WorkDocumentEquationArgumentProperties,
  type WorkDocumentEquationControlRevision,
  type WorkDocumentEquationDelimiterShape,
  type WorkDocumentEquationExpression,
  type WorkDocumentEquationFractionType,
  type WorkDocumentEquationJustification,
  type WorkDocumentEquationLimitLocation,
  type WorkDocumentEquationManualBreak,
  type WorkDocumentEquationMatrixAlignment,
  type WorkDocumentEquationMatrixBaseAlignment,
  type WorkDocumentEquationMatrixSpacing,
  type WorkDocumentEquationNaryOperator,
  type WorkDocumentEquationRowSpacingRule,
  type WorkDocumentEquationRunScript,
  type WorkDocumentEquationRunStyle,
  type WorkDocumentEquationSpacingRule,
  type WorkDocumentEquationThemeColor,
  type WorkDocumentEquationThemeFont,
  type WorkDocumentEquationUnderlineStyle,
  type WorkDocumentEquationWordColor,
  type WorkDocumentEquationWordCombineBrackets,
  type WorkDocumentEquationWordEastAsianLayout,
  type WorkDocumentEquationWordEmphasisMark,
  type WorkDocumentEquationWordFitText,
  type WorkDocumentEquationWordHighlight,
  type WorkDocumentEquationWordLanguages,
  type WorkDocumentEquationWordLineBorderStyle,
  type WorkDocumentEquationWordRunBorder,
  type WorkDocumentEquationWordRunFonts,
  type WorkDocumentEquationWordRunProperties,
  type WorkDocumentEquationWordShading,
  type WorkDocumentEquationWordShadingPattern,
  type WorkDocumentEquationWordTextEffect,
  type WorkDocumentEquationWordUnderline,
  type WorkDocumentEquationWordVerticalAlignment,
} from './work-document-equations';
import {
  closestDocxEquationLikeRoot,
  docxEquationPlacement,
  docxEquationTextNodes,
  docxEquationWordFallbackForContext,
  docxEquationWordReplacement,
  escapeDocxEquationHtml,
  isDocxEquationLikeRoot,
  replaceDocxEquationTextMarker,
} from './work-docx-equation-story';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import {
  descendants,
  directChildren,
  type OoxmlPackage,
} from './work-ooxml-package';

export interface ImportedDocxEquationMarkers {
  equations: ImportedDocxEquationMarker[];
}

export interface MarkedDocxEquationPart {
  document: Document;
  path: string;
}

export interface DocxEquationInspection {
  equation: WorkDocumentEquation | null;
  text: string;
  status: 'supported' | 'unsupported' | 'spoofed';
}

interface ImportedDocxEquationMarker {
  marker: string;
  equation: WorkDocumentEquation;
}

interface EquationMarkerState {
  markers: ImportedDocxEquationMarker[];
  nextMarker: number;
  occupiedText: string;
}

interface EquationParseState {
  depth: number;
  nodes: number;
  textLength: number;
}

interface ParsedMathRunProperties {
  literal?: boolean;
  normalText?: boolean;
  script?: WorkDocumentEquationRunScript;
  style?: WorkDocumentEquationRunStyle;
  manualBreak?: WorkDocumentEquationManualBreak;
  alignment?: boolean;
}

interface ParsedMathArgument {
  children: WorkDocumentEquationExpression[];
  properties?: WorkDocumentEquationArgumentProperties;
}

interface ParsedMathControlProperties {
  controlProperties?: WorkDocumentEquationWordRunProperties;
  controlRevision?: WorkDocumentEquationControlRevision;
}

interface ParsedMathControlRevision {
  revision: WorkDocumentEquationControlRevision;
  controlProperties?: WorkDocumentEquationWordRunProperties;
}

const TRANSITIONAL_MATH_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/math';
const STRICT_MATH_NAMESPACE = 'http://purl.oclc.org/ooxml/officeDocument/math';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const DOCX_MATH_NAMESPACES = new Set([
  TRANSITIONAL_MATH_NAMESPACE,
  STRICT_MATH_NAMESPACE,
]);
const WORD_DATE_UTC_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2023/wordml/word16du';
const WORD_RUN_PROPERTY_ORDER = [
  'rFonts',
  'b',
  'bCs',
  'i',
  'iCs',
  'caps',
  'smallCaps',
  'strike',
  'dstrike',
  'outline',
  'shadow',
  'emboss',
  'imprint',
  'noProof',
  'snapToGrid',
  'vanish',
  'webHidden',
  'color',
  'spacing',
  'w',
  'kern',
  'position',
  'sz',
  'szCs',
  'highlight',
  'u',
  'effect',
  'bdr',
  'shd',
  'fitText',
  'vertAlign',
  'rtl',
  'cs',
  'em',
  'lang',
  'eastAsianLayout',
] as const;
const WORD_THEME_FONTS = new Set<WorkDocumentEquationThemeFont>([
  'majorEastAsia',
  'majorBidi',
  'majorAscii',
  'majorHAnsi',
  'minorEastAsia',
  'minorBidi',
  'minorAscii',
  'minorHAnsi',
]);
const WORD_THEME_COLORS = new Set<WorkDocumentEquationThemeColor>([
  'dark1',
  'light1',
  'dark2',
  'light2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hyperlink',
  'followedHyperlink',
  'none',
  'background1',
  'text1',
  'background2',
  'text2',
]);
const WORD_UNDERLINE_STYLES = new Set<WorkDocumentEquationUnderlineStyle>([
  'none',
  'words',
  'single',
  'double',
  'thick',
  'dotted',
  'dottedHeavy',
  'dash',
  'dashedHeavy',
  'dashLong',
  'dashLongHeavy',
  'dotDash',
  'dashDotHeavy',
  'dotDotDash',
  'dashDotDotHeavy',
  'wave',
  'wavyHeavy',
  'wavyDouble',
]);
const WORD_TEXT_EFFECTS = new Set<WorkDocumentEquationWordTextEffect>([
  'blinkBackground',
  'lights',
  'antsBlack',
  'antsRed',
  'shimmer',
  'sparkle',
  'none',
]);
const WORD_LINE_BORDER_STYLES =
  new Set<WorkDocumentEquationWordLineBorderStyle>([
    'nil',
    'none',
    'single',
    'thick',
    'double',
    'dotted',
    'dashed',
    'dotDash',
    'dotDotDash',
    'triple',
    'thinThickSmallGap',
    'thickThinSmallGap',
    'thinThickThinSmallGap',
    'thinThickMediumGap',
    'thickThinMediumGap',
    'thinThickThinMediumGap',
    'thinThickLargeGap',
    'thickThinLargeGap',
    'thinThickThinLargeGap',
    'wave',
    'doubleWave',
    'dashSmallGap',
    'dashDotStroked',
    'threeDEmboss',
    'threeDEngrave',
    'outset',
    'inset',
  ]);
const WORD_HIGHLIGHT_COLORS = new Set<WorkDocumentEquationWordHighlight>([
  'black',
  'blue',
  'cyan',
  'green',
  'magenta',
  'red',
  'yellow',
  'white',
  'darkBlue',
  'darkCyan',
  'darkGreen',
  'darkMagenta',
  'darkRed',
  'darkYellow',
  'darkGray',
  'lightGray',
  'none',
]);
const WORD_SHADING_PATTERNS = new Set<WorkDocumentEquationWordShadingPattern>([
  'nil',
  'clear',
  'solid',
  'horzStripe',
  'vertStripe',
  'reverseDiagStripe',
  'diagStripe',
  'horzCross',
  'diagCross',
  'thinHorzStripe',
  'thinVertStripe',
  'thinReverseDiagStripe',
  'thinDiagStripe',
  'thinHorzCross',
  'thinDiagCross',
  'pct5',
  'pct10',
  'pct12',
  'pct15',
  'pct20',
  'pct25',
  'pct30',
  'pct35',
  'pct37',
  'pct40',
  'pct45',
  'pct50',
  'pct55',
  'pct60',
  'pct62',
  'pct65',
  'pct70',
  'pct75',
  'pct80',
  'pct85',
  'pct87',
  'pct90',
  'pct95',
]);
const WORD_VERTICAL_ALIGNMENTS =
  new Set<WorkDocumentEquationWordVerticalAlignment>([
    'baseline',
    'superscript',
    'subscript',
  ]);
const WORD_EMPHASIS_MARKS = new Set<WorkDocumentEquationWordEmphasisMark>([
  'none',
  'dot',
  'comma',
  'circle',
  'underDot',
]);
const WORD_COMBINE_BRACKETS = new Set<WorkDocumentEquationWordCombineBrackets>([
  'none',
  'round',
  'square',
  'angle',
  'curly',
]);

const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const MAX_IMPORTED_EQUATIONS = 4_096;
const MAX_MATH_DEPTH = 32;
const MAX_MATH_NODES = 4_096;
const MAX_MATH_TEXT_LENGTH = 65_536;
const MAX_DELIMITER_ARGUMENTS = 32;
const MAX_MATRIX_ROWS = 64;
const MAX_MATRIX_COLUMNS = 64;
const MAX_MATRIX_CELLS = 1_024;
const MAX_MATRIX_SPACING = 65_535;
const MAX_MATRIX_MINIMUM_COLUMN_WIDTH = 31_680;
const MAX_EQUATION_ARRAY_ROWS = 64;
const MAX_WORD_FONT_NAME_LENGTH = 127;
const MAX_WORD_LANGUAGE_LENGTH = 85;
const MAX_WORD_HALF_POINT_SIZE = 1_024;
const MAX_WORD_CHARACTER_SPACING_TWIPS = 31_680;
const MAX_WORD_CHARACTER_SCALE_PERCENT = 600;
const MAX_WORD_KERNING_THRESHOLD_HALF_POINTS = 3_277;
const MIN_WORD_LINE_BORDER_EIGHTH_POINTS = 2;
const MAX_WORD_LINE_BORDER_EIGHTH_POINTS = 96;
const MAX_WORD_BORDER_SPACING_POINTS = 31;
const MAX_WORD_FIT_TEXT_WIDTH_TWIPS = 31_680;
const MIN_WORD_FIT_TEXT_ID = -2_147_483_648;
const MAX_WORD_FIT_TEXT_ID = 2_147_483_647;
const MIN_WORD_EAST_ASIAN_LAYOUT_ID = -2_147_483_648;
const MAX_WORD_EAST_ASIAN_LAYOUT_ID = 2_147_483_647;
const MIN_WORD_POSITION_HALF_POINTS = -2_147_483_648;
const MAX_WORD_POSITION_HALF_POINTS = 2_147_483_647;
const WORD_UNIVERSAL_HALF_POINT_RATIOS: Readonly<
  Record<string, readonly [bigint, bigint]>
> = {
  mm: [720n, 127n],
  cm: [7_200n, 127n],
  in: [144n, 1n],
  pt: [2n, 1n],
  pc: [24n, 1n],
  pi: [24n, 1n],
};
const WORD_UNIVERSAL_TWIP_RATIOS: Readonly<
  Record<string, readonly [bigint, bigint]>
> = {
  mm: [7_200n, 127n],
  cm: [72_000n, 127n],
  in: [1_440n, 1n],
  pt: [20n, 1n],
  pc: [240n, 1n],
  pi: [240n, 1n],
};
const MAX_WORD_MATH_CONTROL_REVISION_ID = 2_147_483_647;
const MAX_WORD_MATH_CONTROL_REVISION_AUTHOR_LENGTH = 255;
const MAX_WORD_MATH_CONTROL_REVISION_DATE_LENGTH = 64;
const WORD_MATH_CONTROL_REVISION_NAMES = new Set([
  'ins',
  'del',
  'moveFrom',
  'moveTo',
]);
const WORD_MATH_INSERTION_CHILD_NAMES = new Set(['del']);
const WORD_MATH_MOVE_CHILD_NAMES = new Set(['ins', 'del']);
const DEFAULT_ACCENT_CHARACTER = '\u0302';
const DEFAULT_DELIMITER_SEPARATOR = '\u2502';
const DEFAULT_GROUP_CHARACTER = '\u23df';
const DEFAULT_NARY_OPERATOR = '\u222b';
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
const NARY_INTEGRALS = new Set<WorkDocumentEquationNaryOperator>([
  '∫',
  '∬',
  '∭',
  '∮',
  '∯',
  '∰',
]);
export async function markDocxPackageEquations(
  archive: OoxmlPackage,
  document: Document,
  existingParts: readonly MarkedDocxEquationPart[] = [],
): Promise<{
  markers: ImportedDocxEquationMarkers;
  parts: MarkedDocxEquationPart[];
  changed: boolean;
}> {
  const parts = new Map(existingParts.map((part) => [part.path, part]));
  for (const config of [
    { path: 'word/footnotes.xml', root: 'footnotes' },
    { path: 'word/endnotes.xml', root: 'endnotes' },
  ] as const) {
    if (!archive.has(config.path) || parts.has(config.path)) continue;
    const part = {
      document: await archive.xml(config.path),
      path: config.path,
    } satisfies MarkedDocxEquationPart;
    const root = part.document.documentElement;
    if (
      root.localName === config.root &&
      DOCX_WORDPROCESSING_NAMESPACES.has(root.namespaceURI ?? '')
    ) {
      parts.set(config.path, part);
    }
  }
  const state: EquationMarkerState = {
    markers: [],
    nextMarker: 1,
    occupiedText: [
      document.documentElement.textContent ?? '',
      ...Array.from(
        parts.values(),
        (part) => part.document.documentElement.textContent ?? '',
      ),
    ].join(''),
  };
  let changed = markDocxEquationDocument(document, state);
  for (const config of [
    { path: 'word/footnotes.xml', root: 'footnotes' },
    { path: 'word/endnotes.xml', root: 'endnotes' },
  ] as const) {
    const part = parts.get(config.path);
    if (!part) continue;
    const root = part.document.documentElement;
    if (
      root.localName !== config.root ||
      !DOCX_WORDPROCESSING_NAMESPACES.has(root.namespaceURI ?? '')
    ) {
      continue;
    }
    changed = markDocxEquationDocument(part.document, state) || changed;
  }
  return {
    markers: { equations: state.markers },
    parts: [...parts.values()],
    changed,
  };
}

export function applyImportedDocxEquationMarkers(
  document: Document,
  markers: ImportedDocxEquationMarkers,
): void {
  for (const marker of markers.equations) {
    const matches = docxEquationTextNodes(document.body).filter((node) =>
      node.data.includes(marker.marker),
    );
    if (matches.length !== 1) {
      for (const match of matches) {
        match.data = match.data.replaceAll(
          marker.marker,
          documentEquationText(marker.equation),
        );
      }
      continue;
    }
    replaceDocxEquationTextMarker(
      matches[0],
      marker.marker,
      createDocumentEquationElement(document, marker.equation),
    );
  }
}

export function hasImportedDocxEquationMarkers(
  markers: ImportedDocxEquationMarkers,
): boolean {
  return markers.equations.length > 0;
}

export function inspectDocxEquation(element: Element): DocxEquationInspection {
  const text = docxEquationPlainText(element);
  if (!['oMath', 'oMathPara'].includes(element.localName)) {
    return { equation: null, text, status: 'unsupported' };
  }
  if (!DOCX_MATH_NAMESPACES.has(element.namespaceURI ?? '')) {
    return { equation: null, text, status: 'spoofed' };
  }
  const equation = parseEquationRoot(element);
  return {
    equation,
    text,
    status: equation ? 'supported' : 'unsupported',
  };
}

export function docxEquationPlainText(element: Element): string {
  const texts = descendants(element, 't').filter(
    (candidate) =>
      DOCX_MATH_NAMESPACES.has(candidate.namespaceURI ?? '') ||
      DOCX_WORDPROCESSING_NAMESPACES.has(candidate.namespaceURI ?? ''),
  );
  const text: string[] = [];
  let length = 0;
  for (const candidate of texts) {
    for (const character of candidate.textContent ?? '') {
      if (length + character.length > MAX_MATH_TEXT_LENGTH) {
        return text.join('');
      }
      text.push(character);
      length += character.length;
    }
  }
  return text.join('');
}

export function isSupportedDocxEquationPlacement(element: Element): boolean {
  return (
    isDocxEquationLikeRoot(element) &&
    DOCX_MATH_NAMESPACES.has(element.namespaceURI ?? '') &&
    docxEquationPlacement(element) !== null
  );
}

export function docxEquationHtml(element: Element): string {
  const inspection = inspectDocxEquation(element);
  if (!inspection.equation || !isSupportedDocxEquationPlacement(element)) {
    return escapeDocxEquationHtml(inspection.text || '[Unsupported equation]');
  }
  const document = new DOMParser().parseFromString('', 'text/html');
  return createDocumentEquationElement(document, inspection.equation).outerHTML;
}

function markDocxEquationDocument(
  document: Document,
  state: EquationMarkerState,
): boolean {
  let changed = false;
  const roots = Array.from(document.querySelectorAll('*')).filter(
    (element) =>
      isDocxEquationLikeRoot(element) &&
      !closestDocxEquationLikeRoot(element.parentElement),
  );
  for (const root of roots) {
    const inspection = inspectDocxEquation(root);
    const placement = docxEquationPlacement(root);
    if (
      inspection.equation &&
      placement &&
      state.markers.length < MAX_IMPORTED_EQUATIONS
    ) {
      const marker = nextEquationMarker(state);
      const equation = {
        ...inspection.equation,
        display: placement === 'block' ? 'block' : inspection.equation.display,
      } satisfies WorkDocumentEquation;
      root.replaceWith(
        docxEquationWordReplacement(
          root.ownerDocument,
          root,
          marker,
          placement,
        ),
      );
      state.markers.push({ marker, equation });
      changed = true;
      continue;
    }
    const fallback = docxEquationWordFallbackForContext(
      root.ownerDocument,
      root,
      inspection.text || '[Unsupported equation]',
      placement,
    );
    if (fallback) root.replaceWith(fallback);
    else root.remove();
    changed = true;
  }
  return changed;
}

function parseEquationRoot(element: Element): WorkDocumentEquation | null {
  if (meaningfulAttributes(element).length) return null;
  const state: EquationParseState = { depth: 0, nodes: 0, textLength: 0 };
  if (element.localName === 'oMathPara') {
    if (
      !structuralChildren(element, new Set(['oMathParaPr', 'oMath'])) ||
      !orderedMathChildren(element, ['oMathParaPr', 'oMath'])
    ) {
      return null;
    }
    const properties = uniqueMathChild(element, 'oMathParaPr', false);
    const equation = uniqueMathChild(element, 'oMath');
    if (properties === null || !equation) {
      return null;
    }
    const justification = parseMathParagraphJustification(properties);
    const children = parseExpressionContainer(equation, state);
    return justification
      ? normalizedEquation('block', children, justification)
      : null;
  }
  const children = parseExpressionContainer(element, state);
  return normalizedEquation('inline', children);
}

function normalizedEquation(
  display: WorkDocumentEquation['display'],
  children: WorkDocumentEquationExpression[] | null,
  justification?: WorkDocumentEquationJustification,
): WorkDocumentEquation | null {
  return children
    ? normalizeDocumentEquation({
        version: 1,
        display,
        ...(justification ? { justification } : {}),
        children,
      })
    : null;
}

function parseMathParagraphJustification(
  properties: Element | undefined,
): WorkDocumentEquationJustification | null {
  if (!properties) return 'centerGroup';
  if (!structuralChildren(properties, new Set(['jc']))) return null;
  const justification = uniqueMathChild(properties, 'jc', false);
  if (justification === null) return null;
  return mathJustificationFromOmml(
    justification
      ? mathValueOrDefault(justification, 'centerGroup')
      : 'centerGroup',
  );
}

function parseExpressionContainer(
  container: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression[] | null {
  if (
    !DOCX_MATH_NAMESPACES.has(container.namespaceURI ?? '') ||
    meaningfulAttributes(container).length ||
    hasMeaningfulDirectText(container)
  ) {
    return null;
  }
  const children = directChildren(container);
  if (
    children.some(
      (child) => !DOCX_MATH_NAMESPACES.has(child.namespaceURI ?? ''),
    )
  ) {
    return null;
  }
  const expressions: WorkDocumentEquationExpression[] = [];
  for (const child of children) {
    if (child.localName === 'argPr' || child.localName === 'ctrlPr')
      return null;
    const expression = parseExpression(child, state);
    if (!expression) return null;
    expressions.push(expression);
  }
  return expressions.length ? expressions : null;
}

function parseMathArgument(
  container: Element,
  state: EquationParseState,
): ParsedMathArgument | null {
  if (
    !DOCX_MATH_NAMESPACES.has(container.namespaceURI ?? '') ||
    meaningfulAttributes(container).length ||
    hasMeaningfulDirectText(container)
  ) {
    return null;
  }
  const children = directChildren(container);
  if (
    children.some(
      (child) => !DOCX_MATH_NAMESPACES.has(child.namespaceURI ?? ''),
    )
  ) {
    return null;
  }
  let expressionStart = 0;
  let expressionEnd = children.length;
  let parsedProperties: WorkDocumentEquationArgumentProperties = {};
  const argumentProperties = children[expressionStart];
  if (argumentProperties?.localName === 'argPr') {
    const parsedArgumentProperties =
      parseMathArgumentProperties(argumentProperties);
    if (parsedArgumentProperties === null) return null;
    parsedProperties = parsedArgumentProperties;
    expressionStart += 1;
  }
  const controlProperties = children[expressionEnd - 1];
  if (controlProperties?.localName === 'ctrlPr') {
    const parsedControlProperties =
      parseMathControlProperties(controlProperties);
    if (parsedControlProperties === null) return null;
    parsedProperties = { ...parsedProperties, ...parsedControlProperties };
    expressionEnd -= 1;
  }
  const expressions: WorkDocumentEquationExpression[] = [];
  for (const child of children.slice(expressionStart, expressionEnd)) {
    if (child.localName === 'argPr' || child.localName === 'ctrlPr') {
      return null;
    }
    const expression = parseExpression(child, state);
    if (!expression) return null;
    expressions.push(expression);
  }
  return {
    children: expressions,
    ...(Object.keys(parsedProperties).length
      ? { properties: parsedProperties }
      : {}),
  };
}

function parseMathArgumentProperties(
  properties: Element,
): WorkDocumentEquationArgumentProperties | null {
  if (!structuralChildren(properties, new Set(['argSz']))) return null;
  const size = uniqueMathChild(properties, 'argSz', false);
  if (size === null) return null;
  const value = size ? mathValueOrDefault(size, '0') : '0';
  if (value === null) return null;
  const source = value.trim();
  if (!/^[+-]?\d+$/u.test(source)) return null;
  const parsed = Number(source);
  if (!Number.isSafeInteger(parsed) || parsed < -2 || parsed > 2) return null;
  return parsed === 0 ? {} : { size: parsed as -2 | -1 | 1 | 2 };
}

function parseExpression(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  if (
    !DOCX_MATH_NAMESPACES.has(element.namespaceURI ?? '') ||
    state.depth >= MAX_MATH_DEPTH ||
    state.nodes >= MAX_MATH_NODES
  ) {
    return null;
  }
  state.depth += 1;
  state.nodes += 1;
  try {
    if (element.localName === 'r') return parseRun(element, state);
    if (element.localName === 'acc') return parseAccent(element, state);
    if (element.localName === 'bar') return parseBar(element, state);
    if (element.localName === 'groupChr') {
      return parseGroupCharacter(element, state);
    }
    if (element.localName === 'phant') return parsePhantom(element, state);
    if (element.localName === 'borderBox') {
      return parseBorderBox(element, state);
    }
    if (element.localName === 'box') return parseBox(element, state);
    if (element.localName === 'f') return parseFraction(element, state);
    if (element.localName === 'sSup') return parseSuperScript(element, state);
    if (element.localName === 'sSub') return parseSubScript(element, state);
    if (element.localName === 'sSubSup') {
      return parseSubSuperScript(element, state);
    }
    if (element.localName === 'sPre') {
      return parsePreSubSuperScript(element, state);
    }
    if (element.localName === 'limLow' || element.localName === 'limUpp') {
      return parseLimit(element, state);
    }
    if (element.localName === 'rad') return parseRadical(element, state);
    if (element.localName === 'func') return parseFunction(element, state);
    if (element.localName === 'nary') return parseNary(element, state);
    if (element.localName === 'eqArr') {
      return parseEquationArray(element, state);
    }
    if (element.localName === 'm') return parseMatrix(element, state);
    if (element.localName === 'd') return parseDelimiter(element, state);
    return null;
  } finally {
    state.depth -= 1;
  }
}

function parseRun(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  if (
    meaningfulAttributes(element).length ||
    hasMeaningfulDirectText(element)
  ) {
    return null;
  }
  const children = directChildren(element);
  let mathProperties: Element | undefined;
  let wordProperties: Element | undefined;
  let textElement: Element | undefined;
  let previous = -1;
  for (const child of children) {
    const position =
      child.localName === 'rPr' &&
      DOCX_MATH_NAMESPACES.has(child.namespaceURI ?? '')
        ? 0
        : child.localName === 'rPr' &&
            DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? '')
          ? 1
          : child.localName === 't' &&
              (DOCX_MATH_NAMESPACES.has(child.namespaceURI ?? '') ||
                DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? ''))
            ? 2
            : -1;
    if (position < previous || position < 0) return null;
    previous = position;
    if (position === 0) {
      if (mathProperties) return null;
      mathProperties = child;
    } else if (position === 1) {
      if (wordProperties) return null;
      wordProperties = child;
    } else {
      if (textElement) return null;
      textElement = child;
    }
  }
  if (
    !textElement ||
    directChildren(textElement).length ||
    !safeMathTextAttributes(textElement)
  ) {
    return null;
  }
  const parsedProperties = mathProperties
    ? parseMathRunProperties(mathProperties)
    : {};
  const parsedWordProperties = wordProperties
    ? parseWordRunProperties(wordProperties)
    : undefined;
  if (!parsedProperties || parsedWordProperties === null) return null;
  const text = textElement.textContent ?? '';
  if (!text || text.length > MAX_MATH_TEXT_LENGTH) return null;
  state.textLength += text.length;
  return state.textLength <= MAX_MATH_TEXT_LENGTH
    ? {
        type: 'run',
        text,
        ...parsedProperties,
        ...(parsedWordProperties
          ? { wordRunProperties: parsedWordProperties }
          : {}),
      }
    : null;
}

function parseWordRunProperties(
  properties: Element,
): WorkDocumentEquationWordRunProperties | null | undefined {
  if (
    !DOCX_WORDPROCESSING_NAMESPACES.has(properties.namespaceURI ?? '') ||
    meaningfulAttributes(properties).length ||
    hasMeaningfulDirectText(properties)
  ) {
    return null;
  }
  const order = new Map(
    WORD_RUN_PROPERTY_ORDER.map((name, index) => [name, index]),
  );
  const children = new Map<string, Element>();
  let previous = -1;
  for (const child of directChildren(properties)) {
    const position = order.get(
      child.localName as (typeof WORD_RUN_PROPERTY_ORDER)[number],
    );
    if (
      !DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? '') ||
      position === undefined ||
      position < previous ||
      children.has(child.localName)
    ) {
      return null;
    }
    previous = position;
    children.set(child.localName, child);
  }
  const fonts = parseWordRunFonts(children.get('rFonts'));
  const color = parseWordRunColor(children.get('color'));
  const characterSpacingTwips = parseWordRequiredInteger(
    children.get('spacing'),
    -MAX_WORD_CHARACTER_SPACING_TWIPS,
    MAX_WORD_CHARACTER_SPACING_TWIPS,
  );
  const characterScalePercent = parseWordCharacterScale(children.get('w'));
  const kerningThresholdHalfPoints = parseWordRequiredInteger(
    children.get('kern'),
    0,
    MAX_WORD_KERNING_THRESHOLD_HALF_POINTS,
  );
  const positionHalfPoints = parseWordPosition(children.get('position'));
  const fontSize = parseWordHalfPointSize(children.get('sz'));
  const fontSizeComplexScript = parseWordHalfPointSize(children.get('szCs'));
  const highlight = parseWordHighlight(children.get('highlight'));
  const underline = parseWordUnderline(children.get('u'));
  const textEffect = parseWordTextEffect(children.get('effect'));
  const border = parseWordRunBorder(children.get('bdr'));
  const shading = parseWordShading(children.get('shd'));
  const fitText = parseWordFitText(children.get('fitText'));
  const verticalAlignment = parseWordVerticalAlignment(
    children.get('vertAlign'),
  );
  const emphasisMark = parseWordEmphasisMark(children.get('em'));
  const languages = parseWordLanguages(children.get('lang'));
  const eastAsianLayout = parseWordEastAsianLayout(
    children.get('eastAsianLayout'),
  );
  if (
    fonts === null ||
    color === null ||
    characterSpacingTwips === null ||
    characterScalePercent === null ||
    kerningThresholdHalfPoints === null ||
    positionHalfPoints === null ||
    fontSize === null ||
    fontSizeComplexScript === null ||
    highlight === null ||
    underline === null ||
    textEffect === null ||
    border === null ||
    shading === null ||
    fitText === null ||
    verticalAlignment === null ||
    emphasisMark === null ||
    languages === null ||
    eastAsianLayout === null
  ) {
    return null;
  }
  const booleans = new Map<
    keyof WorkDocumentEquationWordRunProperties,
    boolean
  >();
  for (const [name, key] of [
    ['b', 'bold'],
    ['bCs', 'boldComplexScript'],
    ['i', 'italic'],
    ['iCs', 'italicComplexScript'],
    ['caps', 'allCaps'],
    ['smallCaps', 'smallCaps'],
    ['strike', 'strike'],
    ['dstrike', 'doubleStrike'],
    ['outline', 'outline'],
    ['shadow', 'shadow'],
    ['emboss', 'emboss'],
    ['imprint', 'imprint'],
    ['noProof', 'noProof'],
    ['snapToGrid', 'snapToGrid'],
    ['vanish', 'hidden'],
    ['webHidden', 'webHidden'],
    ['rtl', 'rightToLeft'],
    ['cs', 'complexScript'],
  ] as const) {
    const element = children.get(name);
    if (!element) continue;
    const value = wordOnOff(element);
    if (value === null) return null;
    booleans.set(key, value);
  }
  const normalized: WorkDocumentEquationWordRunProperties = {
    ...(fonts ? { fonts } : {}),
    ...(booleans.has('bold') ? { bold: booleans.get('bold') } : {}),
    ...(booleans.has('boldComplexScript')
      ? { boldComplexScript: booleans.get('boldComplexScript') }
      : {}),
    ...(booleans.has('italic') ? { italic: booleans.get('italic') } : {}),
    ...(booleans.has('italicComplexScript')
      ? { italicComplexScript: booleans.get('italicComplexScript') }
      : {}),
    ...(booleans.has('allCaps') ? { allCaps: booleans.get('allCaps') } : {}),
    ...(booleans.has('smallCaps')
      ? { smallCaps: booleans.get('smallCaps') }
      : {}),
    ...(booleans.has('strike') ? { strike: booleans.get('strike') } : {}),
    ...(booleans.has('doubleStrike')
      ? { doubleStrike: booleans.get('doubleStrike') }
      : {}),
    ...(booleans.has('outline') ? { outline: booleans.get('outline') } : {}),
    ...(booleans.has('shadow') ? { shadow: booleans.get('shadow') } : {}),
    ...(booleans.has('emboss') ? { emboss: booleans.get('emboss') } : {}),
    ...(booleans.has('imprint') ? { imprint: booleans.get('imprint') } : {}),
    ...(booleans.has('noProof') ? { noProof: booleans.get('noProof') } : {}),
    ...(booleans.has('snapToGrid')
      ? { snapToGrid: booleans.get('snapToGrid') }
      : {}),
    ...(booleans.has('hidden') ? { hidden: booleans.get('hidden') } : {}),
    ...(booleans.has('webHidden')
      ? { webHidden: booleans.get('webHidden') }
      : {}),
    ...(color ? { color } : {}),
    ...(characterSpacingTwips !== undefined ? { characterSpacingTwips } : {}),
    ...(characterScalePercent !== undefined ? { characterScalePercent } : {}),
    ...(kerningThresholdHalfPoints !== undefined
      ? { kerningThresholdHalfPoints }
      : {}),
    ...(positionHalfPoints !== undefined ? { positionHalfPoints } : {}),
    ...(fontSize !== undefined ? { fontSize } : {}),
    ...(fontSizeComplexScript !== undefined ? { fontSizeComplexScript } : {}),
    ...(highlight ? { highlight } : {}),
    ...(underline ? { underline } : {}),
    ...(textEffect ? { textEffect } : {}),
    ...(border ? { border } : {}),
    ...(shading ? { shading } : {}),
    ...(fitText ? { fitText } : {}),
    ...(verticalAlignment ? { verticalAlignment } : {}),
    ...(booleans.has('rightToLeft')
      ? { rightToLeft: booleans.get('rightToLeft') }
      : {}),
    ...(booleans.has('complexScript')
      ? { complexScript: booleans.get('complexScript') }
      : {}),
    ...(emphasisMark ? { emphasisMark } : {}),
    ...(languages ? { languages } : {}),
    ...(eastAsianLayout ? { eastAsianLayout } : {}),
  };
  return Object.keys(normalized).length ? normalized : undefined;
}

function parseMathControlProperties(
  controlProperties: Element | undefined,
): ParsedMathControlProperties | null {
  if (!controlProperties) return {};
  if (
    !DOCX_MATH_NAMESPACES.has(controlProperties.namespaceURI ?? '') ||
    meaningfulAttributes(controlProperties).length ||
    hasMeaningfulDirectText(controlProperties)
  ) {
    return null;
  }
  const children = directChildren(controlProperties);
  if (children.length > 1) return null;
  const child = children[0];
  if (!child) return {};
  if (child.localName === 'rPr') {
    if (!DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? '')) {
      return null;
    }
    const parsed = parseWordRunProperties(child);
    return parsed === null ? null : parsed ? { controlProperties: parsed } : {};
  }
  const parsed = parseMathControlRevision(
    child,
    WORD_MATH_CONTROL_REVISION_NAMES,
  );
  return parsed
    ? {
        controlRevision: parsed.revision,
        ...(parsed.controlProperties
          ? { controlProperties: parsed.controlProperties }
          : {}),
      }
    : null;
}

function parseMathControlRevision(
  element: Element,
  allowedNames: ReadonlySet<string>,
): ParsedMathControlRevision | null {
  if (
    !DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '') ||
    !allowedNames.has(element.localName) ||
    hasMeaningfulDirectText(element)
  ) {
    return null;
  }
  const attributes = wordMathControlRevisionAttributes(element);
  if (
    !attributes?.has('id') ||
    !attributes.has('author') ||
    attributes.size > 4
  ) {
    return null;
  }
  const sourceId = attributes.get('id')?.trim() ?? '';
  const author = attributes.get('author')?.trim() ?? '';
  const date = attributes.get('date')?.trim();
  const dateUtc = attributes.get('dateUtc')?.trim();
  if (
    !/^\+?\d+$/u.test(sourceId) ||
    !author ||
    author.length > MAX_WORD_MATH_CONTROL_REVISION_AUTHOR_LENGTH ||
    /[\p{Cc}\p{Cs}]/u.test(author) ||
    (date !== undefined &&
      (!date || date.length > MAX_WORD_MATH_CONTROL_REVISION_DATE_LENGTH)) ||
    (dateUtc !== undefined &&
      (!dateUtc || dateUtc.length > MAX_WORD_MATH_CONTROL_REVISION_DATE_LENGTH))
  ) {
    return null;
  }
  const id = Number(sourceId);
  if (
    !Number.isSafeInteger(id) ||
    id < 0 ||
    id > MAX_WORD_MATH_CONTROL_REVISION_ID
  ) {
    return null;
  }

  const children = directChildren(element);
  if (children.length > 1) return null;
  const child = children[0];
  let nested: ParsedMathControlRevision | undefined;
  let controlProperties: WorkDocumentEquationWordRunProperties | undefined;
  if (child?.localName === 'rPr') {
    if (!DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? '')) {
      return null;
    }
    const parsed = parseWordRunProperties(child);
    if (parsed === null) return null;
    controlProperties = parsed;
  } else if (child) {
    const childNames =
      element.localName === 'ins'
        ? WORD_MATH_INSERTION_CHILD_NAMES
        : element.localName === 'moveFrom' || element.localName === 'moveTo'
          ? WORD_MATH_MOVE_CHILD_NAMES
          : undefined;
    if (!childNames) return null;
    nested = parseMathControlRevision(child, childNames) ?? undefined;
    if (!nested) return null;
    controlProperties = nested.controlProperties;
  }

  const kind =
    element.localName === 'ins'
      ? 'insertion'
      : element.localName === 'del'
        ? 'deletion'
        : element.localName === 'moveFrom'
          ? 'moveFrom'
          : 'moveTo';
  const revision = {
    kind,
    id,
    author,
    ...(date ? { date } : {}),
    ...(dateUtc ? { dateUtc } : {}),
    ...(nested ? { child: nested.revision } : {}),
  } as WorkDocumentEquationControlRevision;
  return {
    revision,
    ...(controlProperties ? { controlProperties } : {}),
  };
}

function parseWordRunFonts(
  element: Element | undefined,
): WorkDocumentEquationWordRunFonts | null | undefined {
  if (!element) return undefined;
  const attributes = wordLeafAttributes(
    element,
    new Set([
      'ascii',
      'hAnsi',
      'eastAsia',
      'cs',
      'asciiTheme',
      'hAnsiTheme',
      'eastAsiaTheme',
      'cstheme',
      'hint',
    ]),
  );
  if (!attributes) return null;
  const ascii = wordFontName(attributes.get('ascii'));
  const highAnsi = wordFontName(attributes.get('hAnsi'));
  const eastAsia = wordFontName(attributes.get('eastAsia'));
  const complexScript = wordFontName(attributes.get('cs'));
  const asciiTheme = wordThemeFont(attributes.get('asciiTheme'));
  const highAnsiTheme = wordThemeFont(attributes.get('hAnsiTheme'));
  const eastAsiaTheme = wordThemeFont(attributes.get('eastAsiaTheme'));
  const complexScriptTheme = wordThemeFont(attributes.get('cstheme'));
  const hint = attributes.get('hint');
  if (
    ascii === null ||
    highAnsi === null ||
    eastAsia === null ||
    complexScript === null ||
    asciiTheme === null ||
    highAnsiTheme === null ||
    eastAsiaTheme === null ||
    complexScriptTheme === null ||
    (hint !== undefined && !['default', 'eastAsia', 'cs'].includes(hint))
  ) {
    return null;
  }
  const fonts: WorkDocumentEquationWordRunFonts = {
    ...(ascii ? { ascii } : {}),
    ...(highAnsi ? { highAnsi } : {}),
    ...(eastAsia ? { eastAsia } : {}),
    ...(complexScript ? { complexScript } : {}),
    ...(asciiTheme ? { asciiTheme } : {}),
    ...(highAnsiTheme ? { highAnsiTheme } : {}),
    ...(eastAsiaTheme ? { eastAsiaTheme } : {}),
    ...(complexScriptTheme ? { complexScriptTheme } : {}),
    ...(hint ? { hint: hint as WorkDocumentEquationWordRunFonts['hint'] } : {}),
  };
  return Object.keys(fonts).length ? fonts : undefined;
}

function parseWordRunColor(
  element: Element | undefined,
): WorkDocumentEquationWordColor | null | undefined {
  if (!element) return undefined;
  const attributes = wordLeafAttributes(
    element,
    new Set(['val', 'themeColor', 'themeTint', 'themeShade']),
  );
  return attributes ? wordColor(attributes, 'val', true) : null;
}

function parseWordRequiredInteger(
  element: Element | undefined,
  minimum: number,
  maximum: number,
): number | null | undefined {
  if (!element) return undefined;
  const attributes = wordLeafAttributes(element, new Set(['val']));
  if (!attributes || attributes.size !== 1) return null;
  return wordIntegerValue(
    attributes.get('val')?.trim() ?? '',
    minimum,
    maximum,
  );
}

function parseWordCharacterScale(
  element: Element | undefined,
): number | null | undefined {
  if (!element) return undefined;
  const attributes = wordLeafAttributes(element, new Set(['val']));
  if (!attributes) return null;
  if (!attributes.size) return 100;
  return wordIntegerValue(
    attributes.get('val')?.trim() ?? '',
    1,
    MAX_WORD_CHARACTER_SCALE_PERCENT,
  );
}

function parseWordPosition(
  element: Element | undefined,
): number | null | undefined {
  if (!element) return undefined;
  const attributes = wordLeafAttributes(element, new Set(['val']));
  if (!attributes || attributes.size !== 1) return null;
  const source = attributes.get('val')?.trim() ?? '';
  const integer = wordIntegerValue(
    source,
    MIN_WORD_POSITION_HALF_POINTS,
    MAX_WORD_POSITION_HALF_POINTS,
  );
  if (integer !== null) return integer;
  return element.namespaceURI === STRICT_WORD_NAMESPACE
    ? strictUniversalHalfPoints(
        source,
        MIN_WORD_POSITION_HALF_POINTS,
        MAX_WORD_POSITION_HALF_POINTS,
      )
    : null;
}

function parseWordHalfPointSize(
  element: Element | undefined,
): number | null | undefined {
  if (!element) return undefined;
  const attributes = wordLeafAttributes(element, new Set(['val']));
  if (!attributes || attributes.size !== 1) return null;
  const source = attributes.get('val')?.trim() ?? '';
  const integer = wordIntegerValue(source, 1, MAX_WORD_HALF_POINT_SIZE);
  const halfPoints =
    integer ??
    (element.namespaceURI === STRICT_WORD_NAMESPACE
      ? strictUniversalHalfPoints(source, 1, MAX_WORD_HALF_POINT_SIZE)
      : null);
  return halfPoints === null ? null : halfPoints / 2;
}

function wordIntegerValue(
  source: string,
  minimum: number,
  maximum: number,
): number | null {
  if (source.length > 32) return null;
  const pattern = minimum < 0 ? /^[+-]?\d+$/u : /^\+?\d+$/u;
  if (!pattern.test(source)) return null;
  const value = Number(source);
  return Number.isSafeInteger(value) && value >= minimum && value <= maximum
    ? Object.is(value, -0)
      ? 0
      : value
    : null;
}

function strictUniversalHalfPoints(
  source: string,
  minimum: number,
  maximum: number,
): number | null {
  return strictUniversalMeasure(
    source,
    WORD_UNIVERSAL_HALF_POINT_RATIOS,
    minimum,
    maximum,
  );
}

function strictUniversalTwips(
  source: string,
  minimum: number,
  maximum: number,
): number | null {
  return strictUniversalMeasure(
    source,
    WORD_UNIVERSAL_TWIP_RATIOS,
    minimum,
    maximum,
  );
}

function strictUniversalMeasure(
  source: string,
  conversions: Readonly<Record<string, readonly [bigint, bigint]>>,
  minimum: number,
  maximum: number,
): number | null {
  if (source.length > 64) return null;
  const match = /^(-?)(\d+)(?:\.(\d+))?(mm|cm|in|pt|pc|pi)$/u.exec(source);
  if (!match) return null;
  const [, sign, whole = '', fraction = '', unit = ''] = match;
  const conversion = conversions[unit];
  if (!conversion) return null;
  const decimalDigits = `${whole}${fraction}`;
  const decimalDenominator = 10n ** BigInt(fraction.length);
  const numerator =
    BigInt(decimalDigits) * conversion[0] * (sign === '-' ? -1n : 1n);
  const denominator = decimalDenominator * conversion[1];
  if (numerator % denominator !== 0n) return null;
  const units = numerator / denominator;
  if (units < BigInt(minimum) || units > BigInt(maximum)) {
    return null;
  }
  const value = Number(units);
  return Object.is(value, -0) ? 0 : value;
}

function parseWordHighlight(
  element: Element | undefined,
): WorkDocumentEquationWordHighlight | null | undefined {
  if (!element) return undefined;
  const attributes = wordLeafAttributes(element, new Set(['val']));
  const value = attributes?.get('val')?.trim();
  return attributes?.size === 1 &&
    WORD_HIGHLIGHT_COLORS.has(value as WorkDocumentEquationWordHighlight)
    ? (value as WorkDocumentEquationWordHighlight)
    : null;
}

function parseWordUnderline(
  element: Element | undefined,
): WorkDocumentEquationWordUnderline | null | undefined {
  if (!element) return undefined;
  const attributes = wordLeafAttributes(
    element,
    new Set(['val', 'color', 'themeColor', 'themeTint', 'themeShade']),
  );
  const style = attributes?.get('val');
  if (
    !attributes ||
    !style ||
    !WORD_UNDERLINE_STYLES.has(style as WorkDocumentEquationUnderlineStyle)
  ) {
    return null;
  }
  const color = wordColor(attributes, 'color', false);
  return color === null
    ? null
    : {
        style: style as WorkDocumentEquationUnderlineStyle,
        ...(color ? { color } : {}),
      };
}

function parseWordTextEffect(
  element: Element | undefined,
): WorkDocumentEquationWordTextEffect | null | undefined {
  if (!element) return undefined;
  const attributes = wordLeafAttributes(element, new Set(['val']));
  const value = attributes?.get('val')?.trim();
  return attributes?.size === 1 &&
    WORD_TEXT_EFFECTS.has(value as WorkDocumentEquationWordTextEffect)
    ? (value as WorkDocumentEquationWordTextEffect)
    : null;
}

function parseWordRunBorder(
  element: Element | undefined,
): WorkDocumentEquationWordRunBorder | null | undefined {
  if (!element) return undefined;
  const attributes = wordLeafAttributes(
    element,
    new Set([
      'val',
      'color',
      'themeColor',
      'themeTint',
      'themeShade',
      'sz',
      'space',
      'shadow',
      'frame',
    ]),
  );
  const style = attributes?.get('val')?.trim();
  if (
    !attributes ||
    !WORD_LINE_BORDER_STYLES.has(
      style as WorkDocumentEquationWordLineBorderStyle,
    )
  ) {
    return null;
  }
  const color = wordColor(attributes, 'color', false);
  const rawSize = attributes.get('sz')?.trim();
  const sizeEighthPoints =
    rawSize === undefined
      ? undefined
      : wordIntegerValue(
          rawSize,
          MIN_WORD_LINE_BORDER_EIGHTH_POINTS,
          MAX_WORD_LINE_BORDER_EIGHTH_POINTS,
        );
  const rawSpacing = attributes.get('space')?.trim();
  const spacingPoints =
    rawSpacing === undefined
      ? undefined
      : wordIntegerValue(rawSpacing, 0, MAX_WORD_BORDER_SPACING_POINTS);
  const shadow = wordOnOffAttribute(attributes.get('shadow'));
  const frame = wordOnOffAttribute(attributes.get('frame'));
  if (
    color === null ||
    sizeEighthPoints === null ||
    spacingPoints === null ||
    shadow === null ||
    frame === null
  ) {
    return null;
  }
  return {
    style: style as WorkDocumentEquationWordLineBorderStyle,
    ...(color ? { color } : {}),
    ...(sizeEighthPoints !== undefined ? { sizeEighthPoints } : {}),
    ...(spacingPoints !== undefined ? { spacingPoints } : {}),
    ...(shadow !== undefined ? { shadow } : {}),
    ...(frame !== undefined ? { frame } : {}),
  };
}

function parseWordShading(
  element: Element | undefined,
): WorkDocumentEquationWordShading | null | undefined {
  if (!element) return undefined;
  const attributes = wordLeafAttributes(
    element,
    new Set([
      'val',
      'color',
      'themeColor',
      'themeTint',
      'themeShade',
      'fill',
      'themeFill',
      'themeFillTint',
      'themeFillShade',
    ]),
  );
  const pattern = attributes?.get('val')?.trim();
  if (
    !attributes ||
    !WORD_SHADING_PATTERNS.has(
      pattern as WorkDocumentEquationWordShadingPattern,
    )
  ) {
    return null;
  }
  const color = wordColor(attributes, 'color', false);
  const fill = wordColor(
    attributes,
    'fill',
    false,
    'themeFill',
    'themeFillTint',
    'themeFillShade',
  );
  return color === null || fill === null
    ? null
    : {
        pattern: pattern as WorkDocumentEquationWordShadingPattern,
        ...(color ? { color } : {}),
        ...(fill ? { fill } : {}),
      };
}

function parseWordFitText(
  element: Element | undefined,
): WorkDocumentEquationWordFitText | null | undefined {
  if (!element) return undefined;
  const attributes = wordLeafAttributes(element, new Set(['val', 'id']));
  if (!attributes?.has('val')) return null;
  const source = attributes.get('val')?.trim() ?? '';
  const integer = wordIntegerValue(source, 0, MAX_WORD_FIT_TEXT_WIDTH_TWIPS);
  const widthTwips =
    integer ??
    (element.namespaceURI === STRICT_WORD_NAMESPACE
      ? strictUniversalTwips(source, 0, MAX_WORD_FIT_TEXT_WIDTH_TWIPS)
      : null);
  const id = attributes.has('id')
    ? wordIntegerValue(
        attributes.get('id')?.trim() ?? '',
        MIN_WORD_FIT_TEXT_ID,
        MAX_WORD_FIT_TEXT_ID,
      )
    : undefined;
  return widthTwips === null || id === null
    ? null
    : {
        widthTwips,
        ...(id !== undefined ? { id } : {}),
      };
}

function parseWordVerticalAlignment(
  element: Element | undefined,
): WorkDocumentEquationWordVerticalAlignment | null | undefined {
  if (!element) return undefined;
  const attributes = wordLeafAttributes(element, new Set(['val']));
  const value = attributes?.get('val')?.trim();
  return attributes?.size === 1 &&
    WORD_VERTICAL_ALIGNMENTS.has(
      value as WorkDocumentEquationWordVerticalAlignment,
    )
    ? (value as WorkDocumentEquationWordVerticalAlignment)
    : null;
}

function parseWordEmphasisMark(
  element: Element | undefined,
): WorkDocumentEquationWordEmphasisMark | null | undefined {
  if (!element) return undefined;
  const attributes = wordLeafAttributes(element, new Set(['val']));
  const value = attributes?.get('val')?.trim();
  return attributes?.size === 1 &&
    WORD_EMPHASIS_MARKS.has(value as WorkDocumentEquationWordEmphasisMark)
    ? (value as WorkDocumentEquationWordEmphasisMark)
    : null;
}

function parseWordLanguages(
  element: Element | undefined,
): WorkDocumentEquationWordLanguages | null | undefined {
  if (!element) return undefined;
  const attributes = wordLeafAttributes(
    element,
    new Set(['val', 'eastAsia', 'bidi']),
  );
  if (!attributes) return null;
  const latin = wordLanguage(attributes.get('val'));
  const eastAsia = wordLanguage(attributes.get('eastAsia'));
  const bidi = wordLanguage(attributes.get('bidi'));
  if (latin === null || eastAsia === null || bidi === null) return null;
  const languages: WorkDocumentEquationWordLanguages = {
    ...(latin ? { latin } : {}),
    ...(eastAsia ? { eastAsia } : {}),
    ...(bidi ? { bidi } : {}),
  };
  return Object.keys(languages).length ? languages : undefined;
}

function parseWordEastAsianLayout(
  element: Element | undefined,
): WorkDocumentEquationWordEastAsianLayout | null | undefined {
  if (!element) return undefined;
  const attributes = wordLeafAttributes(
    element,
    new Set(['id', 'combine', 'combineBrackets', 'vert', 'vertCompress']),
  );
  if (!attributes) return null;
  const id = attributes.has('id')
    ? wordIntegerValue(
        attributes.get('id')?.trim() ?? '',
        MIN_WORD_EAST_ASIAN_LAYOUT_ID,
        MAX_WORD_EAST_ASIAN_LAYOUT_ID,
      )
    : undefined;
  const combine = wordOnOffAttribute(attributes.get('combine'));
  const rawCombineBrackets = attributes.get('combineBrackets')?.trim();
  const combineBrackets =
    rawCombineBrackets === undefined
      ? undefined
      : WORD_COMBINE_BRACKETS.has(
            rawCombineBrackets as WorkDocumentEquationWordCombineBrackets,
          )
        ? (rawCombineBrackets as WorkDocumentEquationWordCombineBrackets)
        : null;
  const vertical = wordOnOffAttribute(attributes.get('vert'));
  const verticalCompress = wordOnOffAttribute(attributes.get('vertCompress'));
  if (
    id === null ||
    combine === null ||
    combineBrackets === null ||
    vertical === null ||
    verticalCompress === null
  ) {
    return null;
  }
  const layout: WorkDocumentEquationWordEastAsianLayout = {
    ...(id !== undefined ? { id } : {}),
    ...(combine !== undefined ? { combine } : {}),
    ...(combineBrackets !== undefined ? { combineBrackets } : {}),
    ...(vertical !== undefined ? { vertical } : {}),
    ...(verticalCompress !== undefined ? { verticalCompress } : {}),
  };
  return Object.keys(layout).length ? layout : undefined;
}

function wordOnOff(element: Element): boolean | null {
  const attributes = wordLeafAttributes(element, new Set(['val']));
  if (!attributes) return null;
  if (!attributes.size) return true;
  return wordOnOffAttribute(attributes.get('val')) ?? null;
}

function wordOnOffAttribute(
  source: string | undefined,
): boolean | null | undefined {
  if (source === undefined) return undefined;
  const value = source.trim().toLowerCase();
  if (value === '1' || value === 'true' || value === 'on') return true;
  if (value === '0' || value === 'false' || value === 'off') return false;
  return null;
}

function wordLeafAttributes(
  element: Element,
  allowed: ReadonlySet<string>,
): Map<string, string> | null {
  if (directChildren(element).length) return null;
  return wordElementAttributes(element, allowed);
}

function wordMathControlRevisionAttributes(
  element: Element,
): Map<string, string> | null {
  const result = new Map<string, string>();
  for (const attribute of meaningfulAttributes(element)) {
    const name = xmlAttributeLocalName(attribute);
    const namespace = xmlAttributeNamespace(element, attribute);
    const supported =
      (DOCX_WORDPROCESSING_NAMESPACES.has(namespace ?? '') &&
        ['id', 'author', 'date'].includes(name)) ||
      (namespace === WORD_DATE_UTC_NAMESPACE && name === 'dateUtc');
    if (!supported || result.has(name)) return null;
    result.set(name, attribute.value);
  }
  return result;
}

function wordElementAttributes(
  element: Element,
  allowed: ReadonlySet<string>,
): Map<string, string> | null {
  if (
    !DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '') ||
    hasMeaningfulDirectText(element)
  ) {
    return null;
  }
  const result = new Map<string, string>();
  for (const attribute of meaningfulAttributes(element)) {
    const name = xmlAttributeLocalName(attribute);
    if (
      !DOCX_WORDPROCESSING_NAMESPACES.has(
        xmlAttributeNamespace(element, attribute) ?? '',
      ) ||
      !allowed.has(name) ||
      result.has(name)
    ) {
      return null;
    }
    result.set(name, attribute.value);
  }
  return result;
}

function wordColor(
  attributes: ReadonlyMap<string, string>,
  valueAttribute: 'val' | 'color' | 'fill',
  required: boolean,
  themeAttribute = 'themeColor',
  tintAttribute = 'themeTint',
  shadeAttribute = 'themeShade',
): WorkDocumentEquationWordColor | null | undefined {
  const rawValue = attributes.get(valueAttribute)?.trim();
  const rawTheme = attributes.get(themeAttribute)?.trim();
  const rawTint = attributes.get(tintAttribute)?.trim();
  const rawShade = attributes.get(shadeAttribute)?.trim();
  if (!rawValue && !rawTheme && !rawTint && !rawShade) {
    return required ? null : undefined;
  }
  const value =
    rawValue === undefined
      ? undefined
      : rawValue === 'auto'
        ? 'auto'
        : /^[0-9a-f]{6}$/iu.test(rawValue)
          ? (`#${rawValue.toLowerCase()}` as const)
          : null;
  const theme =
    rawTheme === undefined
      ? undefined
      : WORD_THEME_COLORS.has(rawTheme as WorkDocumentEquationThemeColor)
        ? (rawTheme as WorkDocumentEquationThemeColor)
        : null;
  const tint = wordByteHex(rawTint);
  const shade = wordByteHex(rawShade);
  if (
    value === null ||
    theme === null ||
    tint === null ||
    shade === null ||
    (!value && (!theme || theme === 'none')) ||
    ((!theme || theme === 'none') && (tint || shade))
  ) {
    return null;
  }
  return {
    ...(value ? { value } : {}),
    ...(theme ? { theme } : {}),
    ...(tint ? { tint } : {}),
    ...(shade ? { shade } : {}),
  };
}

function wordFontName(source: string | undefined): string | null | undefined {
  if (source === undefined) return undefined;
  const value = source.trim();
  return value &&
    value.length <= MAX_WORD_FONT_NAME_LENGTH &&
    !/[\p{Cc}\p{Cs}]/u.test(value)
    ? value
    : null;
}

function wordThemeFont(
  source: string | undefined,
): WorkDocumentEquationThemeFont | null | undefined {
  if (source === undefined) return undefined;
  const value = source.trim();
  return WORD_THEME_FONTS.has(value as WorkDocumentEquationThemeFont)
    ? (value as WorkDocumentEquationThemeFont)
    : null;
}

function wordLanguage(source: string | undefined): string | null | undefined {
  if (source === undefined) return undefined;
  const value = source.trim();
  return value &&
    value.length <= MAX_WORD_LANGUAGE_LENGTH &&
    /^(?:x-none|[a-z0-9]{1,8}(?:-[a-z0-9]{1,8})*)$/iu.test(value)
    ? value
    : null;
}

function wordByteHex(source: string | undefined): string | null | undefined {
  if (source === undefined) return undefined;
  const value = source.trim();
  return /^[0-9a-f]{2}$/iu.test(value) ? value.toUpperCase() : null;
}

function parseMathRunProperties(
  properties: Element,
): ParsedMathRunProperties | null {
  const propertyOrder = ['lit', 'nor', 'scr', 'sty', 'brk', 'aln'];
  if (
    !structuralChildren(properties, new Set(propertyOrder)) ||
    !orderedMathChildren(properties, propertyOrder)
  ) {
    return null;
  }
  const literal = mathOnOffProperty(properties, 'lit');
  const normalText = mathOnOffProperty(properties, 'nor');
  const scriptElement = uniqueMathChild(properties, 'scr', false);
  const styleElement = uniqueMathChild(properties, 'sty', false);
  const breakElement = uniqueMathChild(properties, 'brk', false);
  const alignment = mathOnOffProperty(properties, 'aln');
  if (
    literal === null ||
    normalText === null ||
    scriptElement === null ||
    styleElement === null ||
    breakElement === null ||
    alignment === null
  ) {
    return null;
  }
  const script = runScriptFromOmml(
    scriptElement ? mathValueOrDefault(scriptElement, 'roman') : 'roman',
  );
  const style = runStyleFromOmml(
    styleElement ? mathValueOrDefault(styleElement, 'i') : 'i',
  );
  const manualBreak = breakElement ? parseManualBreak(breakElement) : undefined;
  return script && style && manualBreak !== null
    ? {
        ...(literal ? { literal } : {}),
        ...(normalText ? { normalText } : {}),
        ...(script !== 'roman' ? { script } : {}),
        ...(style !== 'italic' ? { style } : {}),
        ...(manualBreak ? { manualBreak } : {}),
        ...(alignment ? { alignment } : {}),
      }
    : null;
}

function parseAccent(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  if (!structuralChildren(element, new Set(['accPr', 'e']))) return null;
  const properties = uniqueMathChild(element, 'accPr', false);
  const body = uniqueMathChild(element, 'e');
  if (
    properties === null ||
    !body ||
    (properties && directChildren(element)[0] !== properties) ||
    (properties && !structuralChildren(properties, new Set(['chr', 'ctrlPr'])))
  ) {
    return null;
  }
  const characterElement = properties
    ? uniqueMathChild(properties, 'chr', false)
    : undefined;
  const controlProperties = properties
    ? uniqueMathChild(properties, 'ctrlPr', false)
    : undefined;
  const parsedControlProperties =
    controlProperties === null
      ? null
      : parseMathControlProperties(controlProperties);
  if (
    characterElement === null ||
    parsedControlProperties === null ||
    (characterElement &&
      controlProperties &&
      properties &&
      directChildren(properties)[0] !== characterElement)
  ) {
    return null;
  }
  const character = characterElement
    ? mathValue(characterElement)
    : DEFAULT_ACCENT_CHARACTER;
  const parsedBody = parseMathArgument(body, state);
  return character && accentCharacter(character) && parsedBody
    ? {
        type: 'accent',
        ...parsedControlProperties,
        character,
        children: parsedBody.children,
        ...(parsedBody.properties
          ? { childrenProperties: parsedBody.properties }
          : {}),
      }
    : null;
}

function parseBar(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  if (!structuralChildren(element, new Set(['barPr', 'e']))) return null;
  const properties = uniqueMathChild(element, 'barPr', false);
  const body = uniqueMathChild(element, 'e');
  if (
    properties === null ||
    !body ||
    (properties && directChildren(element)[0] !== properties) ||
    (properties && !structuralChildren(properties, new Set(['pos', 'ctrlPr'])))
  ) {
    return null;
  }
  const positionElement = properties
    ? uniqueMathChild(properties, 'pos', false)
    : undefined;
  const controlProperties = properties
    ? uniqueMathChild(properties, 'ctrlPr', false)
    : undefined;
  const parsedControlProperties =
    controlProperties === null
      ? null
      : parseMathControlProperties(controlProperties);
  if (
    positionElement === null ||
    parsedControlProperties === null ||
    (positionElement &&
      controlProperties &&
      properties &&
      directChildren(properties)[0] !== positionElement)
  ) {
    return null;
  }
  const sourcePosition = positionElement
    ? mathValueOrDefault(positionElement, 'bot')
    : properties
      ? 'bot'
      : 'top';
  const position = topBottomFromOmml(sourcePosition);
  const parsedBody = parseMathArgument(body, state);
  return position && parsedBody
    ? {
        type: 'bar',
        ...parsedControlProperties,
        position,
        children: parsedBody.children,
        ...(parsedBody.properties
          ? { childrenProperties: parsedBody.properties }
          : {}),
      }
    : null;
}

function parseGroupCharacter(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  if (
    !structuralChildren(element, new Set(['groupChrPr', 'e'])) ||
    !orderedMathChildren(element, ['groupChrPr', 'e'])
  ) {
    return null;
  }
  const properties = uniqueMathChild(element, 'groupChrPr', false);
  const body = uniqueMathChild(element, 'e');
  if (properties === null || !body) return null;

  let character: string | null = DEFAULT_GROUP_CHARACTER;
  let sourcePosition: string | null = 'bot';
  let sourceVerticalJustification: string | null = 'top';
  let parsedControlProperties: ParsedMathControlProperties | null = {};
  if (properties) {
    const propertyOrder = ['chr', 'pos', 'vertJc', 'ctrlPr'];
    if (
      !structuralChildren(properties, new Set(propertyOrder)) ||
      !orderedMathChildren(properties, propertyOrder)
    ) {
      return null;
    }
    const characterElement = uniqueMathChild(properties, 'chr', false);
    const positionElement = uniqueMathChild(properties, 'pos', false);
    const verticalJustificationElement = uniqueMathChild(
      properties,
      'vertJc',
      false,
    );
    const controlProperties = uniqueMathChild(properties, 'ctrlPr', false);
    parsedControlProperties =
      controlProperties === null
        ? null
        : parseMathControlProperties(controlProperties);
    if (
      characterElement === null ||
      positionElement === null ||
      verticalJustificationElement === null ||
      parsedControlProperties === null
    ) {
      return null;
    }
    character = characterElement
      ? mathValueOrDefault(characterElement, '')
      : DEFAULT_GROUP_CHARACTER;
    sourcePosition = positionElement
      ? mathValueOrDefault(positionElement, 'bot')
      : 'bot';
    sourceVerticalJustification = verticalJustificationElement
      ? mathValueOrDefault(verticalJustificationElement, 'bot')
      : 'top';
  }

  const position = topBottomFromOmml(sourcePosition);
  const verticalJustification = topBottomFromOmml(sourceVerticalJustification);
  const parsedBody = parseMathArgument(body, state);
  return character !== null &&
    (character === '' || mathCharacter(character)) &&
    position &&
    verticalJustification &&
    parsedBody
    ? {
        type: 'groupCharacter',
        ...parsedControlProperties,
        character,
        position,
        verticalJustification,
        children: parsedBody.children,
        ...(parsedBody.properties
          ? { childrenProperties: parsedBody.properties }
          : {}),
      }
    : null;
}

function parsePhantom(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  if (
    !structuralChildren(element, new Set(['phantPr', 'e'])) ||
    !orderedMathChildren(element, ['phantPr', 'e'])
  ) {
    return null;
  }
  const properties = uniqueMathChild(element, 'phantPr', false);
  const body = uniqueMathChild(element, 'e');
  if (properties === null || !body) return null;

  let show: boolean | null = true;
  let zeroWidth: boolean | null = false;
  let zeroAscent: boolean | null = false;
  let zeroDescent: boolean | null = false;
  let transparent: boolean | null = false;
  let parsedControlProperties: ParsedMathControlProperties | null = {};
  if (properties) {
    const propertyOrder = [
      'show',
      'zeroWid',
      'zeroAsc',
      'zeroDesc',
      'transp',
      'ctrlPr',
    ];
    if (
      !structuralChildren(properties, new Set(propertyOrder)) ||
      !orderedMathChildren(properties, propertyOrder)
    ) {
      return null;
    }
    const showElement = uniqueMathChild(properties, 'show', false);
    const zeroWidthElement = uniqueMathChild(properties, 'zeroWid', false);
    const zeroAscentElement = uniqueMathChild(properties, 'zeroAsc', false);
    const zeroDescentElement = uniqueMathChild(properties, 'zeroDesc', false);
    const transparentElement = uniqueMathChild(properties, 'transp', false);
    const controlProperties = uniqueMathChild(properties, 'ctrlPr', false);
    parsedControlProperties =
      controlProperties === null
        ? null
        : parseMathControlProperties(controlProperties);
    if (
      showElement === null ||
      zeroWidthElement === null ||
      zeroAscentElement === null ||
      zeroDescentElement === null ||
      transparentElement === null ||
      parsedControlProperties === null
    ) {
      return null;
    }
    show = showElement ? mathOnOff(showElement) : true;
    zeroWidth = zeroWidthElement ? mathOnOff(zeroWidthElement) : false;
    zeroAscent = zeroAscentElement ? mathOnOff(zeroAscentElement) : false;
    zeroDescent = zeroDescentElement ? mathOnOff(zeroDescentElement) : false;
    transparent = transparentElement ? mathOnOff(transparentElement) : false;
  }

  const parsedBody = parseMathArgument(body, state);
  return show !== null &&
    zeroWidth !== null &&
    zeroAscent !== null &&
    zeroDescent !== null &&
    transparent !== null &&
    parsedBody
    ? {
        type: 'phantom',
        ...parsedControlProperties,
        show,
        zeroWidth,
        zeroAscent,
        zeroDescent,
        transparent,
        children: parsedBody.children,
        ...(parsedBody.properties
          ? { childrenProperties: parsedBody.properties }
          : {}),
      }
    : null;
}

function parseBorderBox(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  if (!structuralChildren(element, new Set(['borderBoxPr', 'e']))) {
    return null;
  }
  const properties = uniqueMathChild(element, 'borderBoxPr', false);
  const body = uniqueMathChild(element, 'e');
  const propertyNames = [
    'hideTop',
    'hideBot',
    'hideLeft',
    'hideRight',
    'strikeH',
    'strikeV',
    'strikeBLTR',
    'strikeTLBR',
    'ctrlPr',
  ] as const;
  if (
    properties === null ||
    !body ||
    (properties && directChildren(element)[0] !== properties) ||
    (properties &&
      (!structuralChildren(properties, new Set(propertyNames)) ||
        !orderedMathChildren(properties, propertyNames)))
  ) {
    return null;
  }
  const hideTop = mathOnOffProperty(properties, 'hideTop');
  const hideBottom = mathOnOffProperty(properties, 'hideBot');
  const hideLeft = mathOnOffProperty(properties, 'hideLeft');
  const hideRight = mathOnOffProperty(properties, 'hideRight');
  const strikeHorizontal = mathOnOffProperty(properties, 'strikeH');
  const strikeVertical = mathOnOffProperty(properties, 'strikeV');
  const strikeBottomLeftToTopRight = mathOnOffProperty(
    properties,
    'strikeBLTR',
  );
  const strikeTopLeftToBottomRight = mathOnOffProperty(
    properties,
    'strikeTLBR',
  );
  const controlProperties = properties
    ? uniqueMathChild(properties, 'ctrlPr', false)
    : undefined;
  const parsedControlProperties =
    controlProperties === null
      ? null
      : parseMathControlProperties(controlProperties);
  const parsedBody = parseMathArgument(body, state);
  return hideTop !== null &&
    hideBottom !== null &&
    hideLeft !== null &&
    hideRight !== null &&
    strikeHorizontal !== null &&
    strikeVertical !== null &&
    strikeBottomLeftToTopRight !== null &&
    strikeTopLeftToBottomRight !== null &&
    parsedControlProperties !== null &&
    parsedBody
    ? {
        type: 'borderBox',
        ...parsedControlProperties,
        hideTop,
        hideBottom,
        hideLeft,
        hideRight,
        strikeHorizontal,
        strikeVertical,
        strikeBottomLeftToTopRight,
        strikeTopLeftToBottomRight,
        children: parsedBody.children,
        ...(parsedBody.properties
          ? { childrenProperties: parsedBody.properties }
          : {}),
      }
    : null;
}

function parseBox(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  if (!structuralChildren(element, new Set(['boxPr', 'e']))) return null;
  const properties = uniqueMathChild(element, 'boxPr', false);
  const body = uniqueMathChild(element, 'e');
  const propertyNames = [
    'opEmu',
    'noBreak',
    'diff',
    'brk',
    'aln',
    'ctrlPr',
  ] as const;
  if (
    properties === null ||
    !body ||
    (properties && directChildren(element)[0] !== properties) ||
    (properties &&
      (!structuralChildren(properties, new Set(propertyNames)) ||
        !orderedMathChildren(properties, propertyNames)))
  ) {
    return null;
  }
  const operatorEmulator = mathOnOffProperty(properties, 'opEmu');
  const noBreak = mathOnOffProperty(properties, 'noBreak');
  const differential = mathOnOffProperty(properties, 'diff');
  const alignment = mathOnOffProperty(properties, 'aln');
  const breakElement = properties
    ? uniqueMathChild(properties, 'brk', false)
    : undefined;
  const manualBreak = breakElement
    ? parseManualBreak(breakElement)
    : breakElement;
  const controlProperties = properties
    ? uniqueMathChild(properties, 'ctrlPr', false)
    : undefined;
  const parsedControlProperties =
    controlProperties === null
      ? null
      : parseMathControlProperties(controlProperties);
  const parsedBody = parseMathArgument(body, state);
  return operatorEmulator !== null &&
    noBreak !== null &&
    differential !== null &&
    alignment !== null &&
    breakElement !== null &&
    manualBreak !== null &&
    parsedControlProperties !== null &&
    parsedBody
    ? {
        type: 'box',
        ...parsedControlProperties,
        operatorEmulator,
        noBreak,
        differential,
        alignment,
        ...(manualBreak ? { manualBreak } : {}),
        children: parsedBody.children,
        ...(parsedBody.properties
          ? { childrenProperties: parsedBody.properties }
          : {}),
      }
    : null;
}

function parseFraction(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  const childOrder = ['fPr', 'num', 'den'];
  if (
    !structuralChildren(element, new Set(childOrder)) ||
    !orderedMathChildren(element, childOrder)
  ) {
    return null;
  }
  const properties = uniqueMathChild(element, 'fPr', false);
  const numerator = uniqueMathChild(element, 'num');
  const denominator = uniqueMathChild(element, 'den');
  if (
    properties === null ||
    !numerator ||
    !denominator ||
    (properties && !fractionProperties(properties))
  ) {
    return null;
  }
  const typeElement = properties
    ? uniqueMathChild(properties, 'type', false)
    : undefined;
  const controlProperties = properties
    ? uniqueMathChild(properties, 'ctrlPr', false)
    : undefined;
  const parsedControlProperties =
    controlProperties === null
      ? null
      : parseMathControlProperties(controlProperties);
  if (typeElement === null || parsedControlProperties === null) {
    return null;
  }
  const sourceType = typeElement
    ? mathValueOrDefault(typeElement, 'bar')
    : 'bar';
  if (sourceType === null) return null;
  const fractionType = fractionTypeFromOmml(sourceType);
  const parsedNumerator = parseMathArgument(numerator, state);
  const parsedDenominator = parseMathArgument(denominator, state);
  return fractionType && parsedNumerator && parsedDenominator
    ? {
        type: 'fraction',
        ...parsedControlProperties,
        fractionType,
        numerator: parsedNumerator.children,
        ...(parsedNumerator.properties
          ? { numeratorProperties: parsedNumerator.properties }
          : {}),
        denominator: parsedDenominator.children,
        ...(parsedDenominator.properties
          ? { denominatorProperties: parsedDenominator.properties }
          : {}),
      }
    : null;
}

function fractionProperties(properties: Element): boolean {
  const propertyOrder = ['type', 'ctrlPr'];
  return (
    structuralChildren(properties, new Set(propertyOrder)) &&
    orderedMathChildren(properties, propertyOrder)
  );
}

function parseSuperScript(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  const childOrder = ['sSupPr', 'e', 'sup'];
  if (
    !structuralChildren(element, new Set(childOrder)) ||
    !orderedMathChildren(element, childOrder)
  ) {
    return null;
  }
  const properties = uniqueMathChild(element, 'sSupPr', false);
  const base = uniqueMathChild(element, 'e');
  const superScript = uniqueMathChild(element, 'sup');
  const parsedBase = base ? parseMathArgument(base, state) : null;
  const parsedSuperScript = superScript
    ? parseMathArgument(superScript, state)
    : null;
  const parsedProperties =
    properties === null ? null : parseControlOnlyMathProperties(properties);
  return properties !== null &&
    parsedProperties &&
    parsedBase &&
    parsedSuperScript
    ? {
        type: 'superscript',
        ...parsedProperties,
        base: parsedBase.children,
        ...(parsedBase.properties
          ? { baseProperties: parsedBase.properties }
          : {}),
        superScript: parsedSuperScript.children,
        ...(parsedSuperScript.properties
          ? { superScriptProperties: parsedSuperScript.properties }
          : {}),
      }
    : null;
}

function parseSubScript(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  const childOrder = ['sSubPr', 'e', 'sub'];
  if (
    !structuralChildren(element, new Set(childOrder)) ||
    !orderedMathChildren(element, childOrder)
  ) {
    return null;
  }
  const properties = uniqueMathChild(element, 'sSubPr', false);
  const base = uniqueMathChild(element, 'e');
  const subScript = uniqueMathChild(element, 'sub');
  const parsedBase = base ? parseMathArgument(base, state) : null;
  const parsedSubScript = subScript
    ? parseMathArgument(subScript, state)
    : null;
  const parsedProperties =
    properties === null ? null : parseControlOnlyMathProperties(properties);
  return properties !== null &&
    parsedProperties &&
    parsedBase &&
    parsedSubScript
    ? {
        type: 'subscript',
        ...parsedProperties,
        base: parsedBase.children,
        ...(parsedBase.properties
          ? { baseProperties: parsedBase.properties }
          : {}),
        subScript: parsedSubScript.children,
        ...(parsedSubScript.properties
          ? { subScriptProperties: parsedSubScript.properties }
          : {}),
      }
    : null;
}

function parseSubSuperScript(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  const childOrder = ['sSubSupPr', 'e', 'sub', 'sup'];
  if (
    !structuralChildren(element, new Set(childOrder)) ||
    !orderedMathChildren(element, childOrder)
  ) {
    return null;
  }
  const properties = uniqueMathChild(element, 'sSubSupPr', false);
  const base = uniqueMathChild(element, 'e');
  const subScript = uniqueMathChild(element, 'sub');
  const superScript = uniqueMathChild(element, 'sup');
  const parsedProperties =
    properties === null ? null : parseSubSuperScriptProperties(properties);
  const parsedBase = base ? parseMathArgument(base, state) : null;
  const parsedSubScript = subScript
    ? parseMathArgument(subScript, state)
    : null;
  const parsedSuperScript = superScript
    ? parseMathArgument(superScript, state)
    : null;
  return parsedProperties && parsedBase && parsedSubScript && parsedSuperScript
    ? {
        type: 'subSuperScript',
        ...parsedProperties,
        base: parsedBase.children,
        ...(parsedBase.properties
          ? { baseProperties: parsedBase.properties }
          : {}),
        subScript: parsedSubScript.children,
        ...(parsedSubScript.properties
          ? { subScriptProperties: parsedSubScript.properties }
          : {}),
        superScript: parsedSuperScript.children,
        ...(parsedSuperScript.properties
          ? { superScriptProperties: parsedSuperScript.properties }
          : {}),
      }
    : null;
}

function parseSubSuperScriptProperties(
  properties: Element | undefined,
): (ParsedMathControlProperties & { alignScripts?: boolean }) | null {
  if (!properties) return {};
  const propertyOrder = ['alnScr', 'ctrlPr'];
  if (
    !structuralChildren(properties, new Set(propertyOrder)) ||
    !orderedMathChildren(properties, propertyOrder)
  ) {
    return null;
  }
  const alignScripts = mathOnOffProperty(properties, 'alnScr');
  const controlProperties = uniqueMathChild(properties, 'ctrlPr', false);
  const parsedControlProperties =
    controlProperties === null
      ? null
      : parseMathControlProperties(controlProperties);
  return alignScripts !== null && parsedControlProperties
    ? {
        ...(alignScripts ? { alignScripts } : {}),
        ...parsedControlProperties,
      }
    : null;
}

function parseControlOnlyMathProperties(
  properties: Element | undefined,
): ParsedMathControlProperties | null {
  if (!properties) return {};
  if (!structuralChildren(properties, new Set(['ctrlPr']))) return null;
  const controlProperties = uniqueMathChild(properties, 'ctrlPr', false);
  return controlProperties === null
    ? null
    : parseMathControlProperties(controlProperties);
}

function parsePreSubSuperScript(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  const childOrder = ['sPrePr', 'sub', 'sup', 'e'];
  if (
    !structuralChildren(element, new Set(childOrder)) ||
    !orderedMathChildren(element, childOrder)
  ) {
    return null;
  }
  const properties = uniqueMathChild(element, 'sPrePr', false);
  const subScript = uniqueMathChild(element, 'sub');
  const superScript = uniqueMathChild(element, 'sup');
  const base = uniqueMathChild(element, 'e');
  if (properties === null || !subScript || !superScript || !base) return null;
  const parsedProperties = parseControlOnlyMathProperties(properties);
  if (!parsedProperties) return null;
  const parsedSubScript = parseMathArgument(subScript, state);
  const parsedSuperScript = parseMathArgument(superScript, state);
  const parsedBase = parseMathArgument(base, state);
  return parsedSubScript !== null && parsedSuperScript !== null && parsedBase
    ? {
        type: 'preSubSuperScript',
        ...parsedProperties,
        base: parsedBase.children,
        ...(parsedBase.properties
          ? { baseProperties: parsedBase.properties }
          : {}),
        subScript: parsedSubScript.children,
        ...(parsedSubScript.properties
          ? { subScriptProperties: parsedSubScript.properties }
          : {}),
        superScript: parsedSuperScript.children,
        ...(parsedSuperScript.properties
          ? { superScriptProperties: parsedSuperScript.properties }
          : {}),
      }
    : null;
}

function parseLimit(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  const lower = element.localName === 'limLow';
  const propertyName = lower ? 'limLowPr' : 'limUppPr';
  const allowed = [propertyName, 'e', 'lim'];
  if (
    !structuralChildren(element, new Set(allowed)) ||
    !orderedMathChildren(element, allowed)
  ) {
    return null;
  }
  const properties = uniqueMathChild(element, propertyName, false);
  const base = uniqueMathChild(element, 'e');
  const limit = uniqueMathChild(element, 'lim');
  if (properties === null || !base || !limit) return null;
  const parsedProperties = parseControlOnlyMathProperties(properties);
  if (!parsedProperties) return null;
  const parsedBase = parseMathArgument(base, state);
  const parsedLimit = parseMathArgument(limit, state);
  return parsedBase && parsedLimit
    ? {
        type: lower ? 'lowerLimit' : 'upperLimit',
        ...parsedProperties,
        base: parsedBase.children,
        ...(parsedBase.properties
          ? { baseProperties: parsedBase.properties }
          : {}),
        limit: parsedLimit.children,
        ...(parsedLimit.properties
          ? { limitProperties: parsedLimit.properties }
          : {}),
      }
    : null;
}

function parseRadical(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  const childOrder = ['radPr', 'deg', 'e'];
  if (
    !structuralChildren(element, new Set(childOrder)) ||
    !orderedMathChildren(element, childOrder)
  ) {
    return null;
  }
  const properties = uniqueMathChild(element, 'radPr', false);
  const degree = uniqueMathChild(element, 'deg', false);
  const body = uniqueMathChild(element, 'e');
  if (
    properties === null ||
    degree === null ||
    !body ||
    (properties && !radicalProperties(properties))
  ) {
    return null;
  }
  const degreeHiddenElement = properties
    ? uniqueMathChild(properties, 'degHide', false)
    : undefined;
  const controlProperties = properties
    ? uniqueMathChild(properties, 'ctrlPr', false)
    : undefined;
  const degreeHidden = degreeHiddenElement
    ? mathOnOff(degreeHiddenElement)
    : false;
  const parsedControlProperties =
    controlProperties === null
      ? null
      : parseMathControlProperties(controlProperties);
  if (
    degreeHiddenElement === null ||
    parsedControlProperties === null ||
    degreeHidden === null
  ) {
    return null;
  }
  const parsedBody = parseMathArgument(body, state);
  const parsedDegree = degree ? parseMathArgument(degree, state) : undefined;
  if (
    !parsedBody ||
    parsedDegree === null ||
    (degreeHidden && Boolean(parsedDegree?.children.length))
  ) {
    return null;
  }
  return {
    type: 'radical',
    ...parsedControlProperties,
    children: parsedBody.children,
    ...(parsedBody.properties
      ? { childrenProperties: parsedBody.properties }
      : {}),
    ...(parsedDegree?.children.length ? { degree: parsedDegree.children } : {}),
    ...(parsedDegree?.properties
      ? { degreeProperties: parsedDegree.properties }
      : {}),
  };
}

function radicalProperties(properties: Element): boolean {
  const propertyOrder = ['degHide', 'ctrlPr'];
  return (
    structuralChildren(properties, new Set(propertyOrder)) &&
    orderedMathChildren(properties, propertyOrder)
  );
}

function parseFunction(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  const childOrder = ['funcPr', 'fName', 'e'];
  if (
    !structuralChildren(element, new Set(childOrder)) ||
    !orderedMathChildren(element, childOrder)
  ) {
    return null;
  }
  const properties = uniqueMathChild(element, 'funcPr', false);
  const name = uniqueMathChild(element, 'fName');
  const body = uniqueMathChild(element, 'e');
  const parsedName = name ? parseMathArgument(name, state) : null;
  const parsedBody = body ? parseMathArgument(body, state) : null;
  const parsedProperties =
    properties === null ? null : parseControlOnlyMathProperties(properties);
  return properties !== null && parsedProperties && parsedName && parsedBody
    ? {
        type: 'function',
        ...parsedProperties,
        name: parsedName.children,
        ...(parsedName.properties
          ? { nameProperties: parsedName.properties }
          : {}),
        children: parsedBody.children,
        ...(parsedBody.properties
          ? { childrenProperties: parsedBody.properties }
          : {}),
      }
    : null;
}

function parseNary(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  const childOrder = ['naryPr', 'sub', 'sup', 'e'];
  if (
    !structuralChildren(element, new Set(childOrder)) ||
    !orderedMathChildren(element, childOrder)
  ) {
    return null;
  }
  const properties = uniqueMathChild(element, 'naryPr', false);
  const subScript = uniqueMathChild(element, 'sub');
  const superScript = uniqueMathChild(element, 'sup');
  const body = uniqueMathChild(element, 'e');
  if (
    properties === null ||
    !subScript ||
    !superScript ||
    !body ||
    (properties && !naryProperties(properties))
  ) {
    return null;
  }
  const operatorElement = properties
    ? uniqueMathChild(properties, 'chr', false)
    : undefined;
  const locationElement = properties
    ? uniqueMathChild(properties, 'limLoc', false)
    : undefined;
  const growElement = properties
    ? uniqueMathChild(properties, 'grow', false)
    : undefined;
  const subHiddenElement = properties
    ? uniqueMathChild(properties, 'subHide', false)
    : undefined;
  const superHiddenElement = properties
    ? uniqueMathChild(properties, 'supHide', false)
    : undefined;
  const controlProperties = properties
    ? uniqueMathChild(properties, 'ctrlPr', false)
    : undefined;
  const parsedControlProperties =
    controlProperties === null
      ? null
      : parseMathControlProperties(controlProperties);
  if (
    operatorElement === null ||
    locationElement === null ||
    growElement === null ||
    subHiddenElement === null ||
    superHiddenElement === null ||
    parsedControlProperties === null
  ) {
    return null;
  }
  const operatorValue = operatorElement
    ? mathValueOrDefault(operatorElement, '')
    : DEFAULT_NARY_OPERATOR;
  if (!NARY_OPERATORS.has(operatorValue as WorkDocumentEquationNaryOperator)) {
    return null;
  }
  const operator = operatorValue as WorkDocumentEquationNaryOperator;
  const locationValue = locationElement
    ? mathValueOrDefault(locationElement, 'undOvr')
    : null;
  const limitLocation = naryLimitLocation(operator, locationValue);
  const grow = growElement ? mathOnOff(growElement) : false;
  const subHidden = subHiddenElement ? mathOnOff(subHiddenElement) : false;
  const superHidden = superHiddenElement
    ? mathOnOff(superHiddenElement)
    : false;
  const parsedBody = parseMathArgument(body, state);
  const parsedSubScript = parseMathArgument(subScript, state);
  const parsedSuperScript = parseMathArgument(superScript, state);
  if (
    !limitLocation ||
    grow === null ||
    subHidden === null ||
    superHidden === null ||
    !parsedBody ||
    parsedSubScript === null ||
    parsedSuperScript === null
  ) {
    return null;
  }
  if (
    subHidden === Boolean(parsedSubScript.children.length) ||
    superHidden === Boolean(parsedSuperScript.children.length)
  ) {
    return null;
  }
  return {
    type: 'nary',
    ...parsedControlProperties,
    operator,
    limitLocation,
    ...(grow ? { grow: true } : {}),
    children: parsedBody.children,
    ...(parsedBody.properties
      ? { childrenProperties: parsedBody.properties }
      : {}),
    ...(parsedSubScript.children.length
      ? { subScript: parsedSubScript.children }
      : {}),
    ...(parsedSubScript.properties
      ? { subScriptProperties: parsedSubScript.properties }
      : {}),
    ...(parsedSuperScript.children.length
      ? { superScript: parsedSuperScript.children }
      : {}),
    ...(parsedSuperScript.properties
      ? { superScriptProperties: parsedSuperScript.properties }
      : {}),
  };
}

function naryProperties(properties: Element): boolean {
  const propertyOrder = [
    'chr',
    'limLoc',
    'grow',
    'subHide',
    'supHide',
    'ctrlPr',
  ];
  return (
    structuralChildren(properties, new Set(propertyOrder)) &&
    orderedMathChildren(properties, propertyOrder)
  );
}

function parseEquationArray(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  if (!structuralChildren(element, new Set(['eqArrPr', 'e']))) return null;
  const properties = uniqueMathChild(element, 'eqArrPr', false);
  const rowElements = mathDirectChildren(element, 'e');
  if (
    properties === null ||
    !rowElements.length ||
    rowElements.length > MAX_EQUATION_ARRAY_ROWS ||
    (properties && directChildren(element)[0] !== properties)
  ) {
    return null;
  }
  const parsedProperties = parseEquationArrayProperties(properties);
  if (!parsedProperties) return null;
  const parsedRows = rowElements.map((row) => parseMathArgument(row, state));
  if (!parsedRows.every((row): row is ParsedMathArgument => row !== null)) {
    return null;
  }
  const rowProperties = parsedRows.map((row) => row.properties ?? null);
  return {
    type: 'equationArray',
    ...parsedProperties,
    rows: parsedRows.map((row) => row.children),
    ...(rowProperties.some(Boolean) ? { rowProperties } : {}),
  };
}

function parseEquationArrayProperties(properties: Element | undefined): {
  baseAlignment: WorkDocumentEquationMatrixBaseAlignment;
  maximumDistribution: boolean;
  objectDistribution: boolean;
  rowSpacingRule: WorkDocumentEquationRowSpacingRule;
  rowSpacing: number;
  controlProperties?: WorkDocumentEquationWordRunProperties;
  controlRevision?: WorkDocumentEquationControlRevision;
} | null {
  if (!properties) {
    return {
      baseAlignment: 'center',
      maximumDistribution: false,
      objectDistribution: false,
      rowSpacingRule: 'single',
      rowSpacing: 0,
    };
  }
  const propertyNames = [
    'baseJc',
    'maxDist',
    'objDist',
    'rSpRule',
    'rSp',
    'ctrlPr',
  ] as const;
  if (
    !structuralChildren(properties, new Set(propertyNames)) ||
    !orderedMathChildren(properties, propertyNames)
  ) {
    return null;
  }
  const baseElement = uniqueMathChild(properties, 'baseJc', false);
  const rowSpacingRuleElement = uniqueMathChild(properties, 'rSpRule', false);
  const rowSpacingElement = uniqueMathChild(properties, 'rSp', false);
  const controlProperties = uniqueMathChild(properties, 'ctrlPr', false);
  const parsedControlProperties =
    controlProperties === null
      ? null
      : parseMathControlProperties(controlProperties);
  if (
    baseElement === null ||
    rowSpacingRuleElement === null ||
    rowSpacingElement === null ||
    parsedControlProperties === null
  ) {
    return null;
  }
  const baseValue = baseElement
    ? mathValueOrDefault(baseElement, 'center')
    : 'center';
  const baseAlignment =
    baseValue === 'top'
      ? 'top'
      : baseValue === 'center'
        ? 'center'
        : baseValue === 'bot'
          ? 'bottom'
          : null;
  const maximumDistribution = mathOnOffProperty(properties, 'maxDist');
  const objectDistribution = mathOnOffProperty(properties, 'objDist');
  const rowSpacingRule = mathSpacingRule(rowSpacingRuleElement);
  const rowSpacing = unsignedMathInteger(
    rowSpacingElement ? mathValueOrDefault(rowSpacingElement, '0') : '0',
    65_535,
  );
  return baseAlignment &&
    maximumDistribution !== null &&
    objectDistribution !== null &&
    rowSpacingRule &&
    rowSpacing !== null
    ? {
        ...parsedControlProperties,
        baseAlignment,
        maximumDistribution,
        objectDistribution,
        rowSpacingRule,
        rowSpacing,
      }
    : null;
}

function parseMatrix(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  if (!structuralChildren(element, new Set(['mPr', 'mr']))) return null;
  const properties = uniqueMathChild(element, 'mPr', false);
  const rowElements = mathDirectChildren(element, 'mr');
  if (
    properties === null ||
    !rowElements.length ||
    rowElements.length > MAX_MATRIX_ROWS ||
    (properties && directChildren(element)[0] !== properties)
  ) {
    return null;
  }
  const rows: WorkDocumentEquationExpression[][][] = [];
  const cellProperties: Array<
    Array<WorkDocumentEquationArgumentProperties | null>
  > = [];
  let columnCount = 0;
  for (const rowElement of rowElements) {
    if (!structuralChildren(rowElement, new Set(['e']))) return null;
    const cellElements = mathDirectChildren(rowElement, 'e');
    if (
      !cellElements.length ||
      cellElements.length > MAX_MATRIX_COLUMNS ||
      (columnCount > 0 && cellElements.length !== columnCount)
    ) {
      return null;
    }
    columnCount ||= cellElements.length;
    if (rowElements.length * columnCount > MAX_MATRIX_CELLS) return null;
    const parsedRow = cellElements.map((cell) =>
      parseMathArgument(cell, state),
    );
    if (!parsedRow.every((cell): cell is ParsedMathArgument => cell !== null)) {
      return null;
    }
    rows.push(parsedRow.map((cell) => cell.children));
    cellProperties.push(parsedRow.map((cell) => cell.properties ?? null));
  }
  const parsedProperties = parseMatrixProperties(properties, columnCount);
  return parsedProperties
    ? {
        type: 'matrix',
        ...parsedProperties,
        rows,
        ...(cellProperties.some((row) => row.some(Boolean))
          ? { cellProperties }
          : {}),
      }
    : null;
}

function parseMatrixProperties(
  properties: Element | undefined,
  columnCount: number,
): {
  baseAlignment: WorkDocumentEquationMatrixBaseAlignment;
  placeholdersHidden: boolean;
  columnAlignments: WorkDocumentEquationMatrixAlignment[];
  spacing?: WorkDocumentEquationMatrixSpacing;
  controlProperties?: WorkDocumentEquationWordRunProperties;
  controlRevision?: WorkDocumentEquationControlRevision;
} | null {
  if (!properties) {
    return {
      baseAlignment: 'center',
      placeholdersHidden: false,
      columnAlignments: Array.from({ length: columnCount }, () => 'center'),
    };
  }
  if (
    !structuralChildren(
      properties,
      new Set([
        'baseJc',
        'plcHide',
        'rSpRule',
        'cGpRule',
        'rSp',
        'cSp',
        'cGp',
        'mcs',
        'ctrlPr',
      ]),
    )
  ) {
    return null;
  }
  const order = new Map([
    ['baseJc', 0],
    ['plcHide', 1],
    ['rSpRule', 2],
    ['cGpRule', 3],
    ['rSp', 4],
    ['cSp', 5],
    ['cGp', 6],
    ['mcs', 7],
    ['ctrlPr', 8],
  ]);
  let previous = -1;
  for (const child of directChildren(properties)) {
    const position = order.get(child.localName);
    if (position === undefined || position < previous) return null;
    previous = position;
  }
  const baseElement = uniqueMathChild(properties, 'baseJc', false);
  const placeholderElement = uniqueMathChild(properties, 'plcHide', false);
  const rowSpacingRuleElement = uniqueMathChild(properties, 'rSpRule', false);
  const columnGapRuleElement = uniqueMathChild(properties, 'cGpRule', false);
  const rowSpacingElement = uniqueMathChild(properties, 'rSp', false);
  const minimumColumnWidthElement = uniqueMathChild(properties, 'cSp', false);
  const columnGapElement = uniqueMathChild(properties, 'cGp', false);
  const columnsElement = uniqueMathChild(properties, 'mcs', false);
  const controlProperties = uniqueMathChild(properties, 'ctrlPr', false);
  const parsedControlProperties =
    controlProperties === null
      ? null
      : parseMathControlProperties(controlProperties);
  if (
    baseElement === null ||
    placeholderElement === null ||
    rowSpacingRuleElement === null ||
    columnGapRuleElement === null ||
    rowSpacingElement === null ||
    minimumColumnWidthElement === null ||
    columnGapElement === null ||
    columnsElement === null ||
    parsedControlProperties === null
  ) {
    return null;
  }
  const baseValue = baseElement
    ? mathValueOrDefault(baseElement, 'center')
    : 'center';
  const baseAlignment =
    baseValue === 'top'
      ? 'top'
      : baseValue === 'center'
        ? 'center'
        : baseValue === 'bot'
          ? 'bottom'
          : null;
  const placeholdersHidden = placeholderElement
    ? mathOnOff(placeholderElement)
    : false;
  const hasSpacing = Boolean(
    rowSpacingRuleElement ||
      columnGapRuleElement ||
      rowSpacingElement ||
      minimumColumnWidthElement ||
      columnGapElement,
  );
  const rowSpacingRule = mathSpacingRule(rowSpacingRuleElement);
  const columnGapRule = mathSpacingRule(columnGapRuleElement);
  const rowSpacing = unsignedMathInteger(
    rowSpacingElement ? mathValueOrDefault(rowSpacingElement, '0') : '0',
    MAX_MATRIX_SPACING,
  );
  const minimumColumnWidthTwips = unsignedMathInteger(
    minimumColumnWidthElement
      ? mathValueOrDefault(minimumColumnWidthElement, '0')
      : '0',
    MAX_MATRIX_MINIMUM_COLUMN_WIDTH,
  );
  const columnGap = unsignedMathInteger(
    columnGapElement ? mathValueOrDefault(columnGapElement, '0') : '0',
    MAX_MATRIX_SPACING,
  );
  const columnAlignments = columnsElement
    ? parseMatrixColumns(columnsElement, columnCount)
    : Array.from(
        { length: columnCount },
        (): WorkDocumentEquationMatrixAlignment => 'center',
      );
  return baseAlignment &&
    placeholdersHidden !== null &&
    rowSpacingRule &&
    columnGapRule &&
    rowSpacing !== null &&
    minimumColumnWidthTwips !== null &&
    columnGap !== null &&
    columnAlignments
    ? {
        ...parsedControlProperties,
        baseAlignment,
        placeholdersHidden,
        columnAlignments,
        ...(hasSpacing
          ? {
              spacing: {
                rowSpacingRule,
                rowSpacing,
                columnGapRule,
                columnGap,
                minimumColumnWidthTwips,
              },
            }
          : {}),
      }
    : null;
}

function parseMatrixColumns(
  columns: Element,
  columnCount: number,
): WorkDocumentEquationMatrixAlignment[] | null {
  if (!structuralChildren(columns, new Set(['mc']))) return null;
  const columnGroups = mathDirectChildren(columns, 'mc');
  if (!columnGroups.length || columnGroups.length > MAX_MATRIX_COLUMNS) {
    return null;
  }
  const alignments: WorkDocumentEquationMatrixAlignment[] = [];
  for (const columnGroup of columnGroups) {
    if (!structuralChildren(columnGroup, new Set(['mcPr']))) return null;
    const properties = uniqueMathChild(columnGroup, 'mcPr');
    if (
      !properties ||
      !structuralChildren(properties, new Set(['count', 'mcJc']))
    ) {
      return null;
    }
    const propertyChildren = directChildren(properties);
    if (
      propertyChildren.length === 2 &&
      (propertyChildren[0]?.localName !== 'count' ||
        propertyChildren[1]?.localName !== 'mcJc')
    ) {
      return null;
    }
    const countElement = uniqueMathChild(properties, 'count', false);
    const alignmentElement = uniqueMathChild(properties, 'mcJc', false);
    if (countElement === null || alignmentElement === null) return null;
    const countValue = countElement
      ? mathValueOrDefault(countElement, '1')
      : '1';
    const count = positiveMathInteger(countValue, MAX_MATRIX_COLUMNS);
    const alignmentValue = alignmentElement
      ? mathValueOrDefault(alignmentElement, 'center')
      : 'center';
    const alignment =
      alignmentValue === 'left' ||
      alignmentValue === 'center' ||
      alignmentValue === 'right'
        ? alignmentValue
        : null;
    if (!count || !alignment || alignments.length + count > columnCount) {
      return null;
    }
    alignments.push(...Array.from({ length: count }, () => alignment));
  }
  return alignments.length === columnCount ? alignments : null;
}

function parseDelimiter(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  const childOrder = ['dPr', 'e'];
  if (
    !structuralChildren(element, new Set(childOrder)) ||
    !orderedMathChildren(element, childOrder)
  ) {
    return null;
  }
  const properties = uniqueMathChild(element, 'dPr', false);
  const arguments_ = mathDirectChildren(element, 'e');
  if (
    properties === null ||
    !arguments_.length ||
    arguments_.length > MAX_DELIMITER_ARGUMENTS ||
    (properties && !delimiterProperties(properties))
  ) {
    return null;
  }
  const growElement = properties
    ? uniqueMathChild(properties, 'grow', false)
    : undefined;
  const shapeElement = properties
    ? uniqueMathChild(properties, 'shp', false)
    : undefined;
  const controlProperties = properties
    ? uniqueMathChild(properties, 'ctrlPr', false)
    : undefined;
  const parsedControlProperties =
    controlProperties === null
      ? null
      : parseMathControlProperties(controlProperties);
  if (
    growElement === null ||
    shapeElement === null ||
    parsedControlProperties === null
  ) {
    return null;
  }
  const grow = growElement ? mathOnOff(growElement) : true;
  const shapeValue = shapeElement
    ? mathValueOrDefault(shapeElement, 'centered')
    : 'centered';
  const shape: WorkDocumentEquationDelimiterShape | null =
    shapeValue === 'centered' || shapeValue === 'match' ? shapeValue : null;
  const opening = delimiterProperty(properties, 'begChr', '(');
  const closing = delimiterProperty(properties, 'endChr', ')');
  const separator = delimiterProperty(
    properties,
    'sepChr',
    DEFAULT_DELIMITER_SEPARATOR,
  );
  const parsedArguments = arguments_.map((argument) =>
    parseMathArgument(argument, state),
  );
  return grow !== null &&
    shape &&
    opening !== null &&
    closing !== null &&
    separator !== null &&
    parsedArguments.every(
      (argument): argument is ParsedMathArgument => argument !== null,
    )
    ? {
        type: 'delimiter',
        ...parsedControlProperties,
        opening,
        closing,
        separator,
        ...(grow ? {} : { grow: false }),
        ...(shape === 'match' ? { shape } : {}),
        arguments: parsedArguments.map((argument) => argument.children),
        ...(parsedArguments.some((argument) => argument.properties)
          ? {
              argumentProperties: parsedArguments.map(
                (argument) => argument.properties ?? null,
              ),
            }
          : {}),
      }
    : null;
}

function delimiterProperties(properties: Element): boolean {
  const propertyOrder = ['begChr', 'sepChr', 'endChr', 'grow', 'shp', 'ctrlPr'];
  return (
    structuralChildren(properties, new Set(propertyOrder)) &&
    orderedMathChildren(properties, propertyOrder)
  );
}

function structuralChildren(
  element: Element,
  allowed: ReadonlySet<string>,
): boolean {
  return (
    meaningfulAttributes(element).length === 0 &&
    !hasMeaningfulDirectText(element) &&
    directChildren(element).every(
      (child) =>
        DOCX_MATH_NAMESPACES.has(child.namespaceURI ?? '') &&
        allowed.has(child.localName),
    )
  );
}

function orderedMathChildren(
  element: Element,
  names: readonly string[],
): boolean {
  const order = new Map(names.map((name, index) => [name, index]));
  let previous = -1;
  for (const child of directChildren(element)) {
    const position = order.get(child.localName);
    if (position === undefined || position < previous) return false;
    previous = position;
  }
  return true;
}

function mathDirectChildren(element: Element, name: string): Element[] {
  return directChildren(element, name).filter((child) =>
    DOCX_MATH_NAMESPACES.has(child.namespaceURI ?? ''),
  );
}

function uniqueMathChild(
  element: Element,
  name: string,
  required = true,
): Element | null | undefined {
  const localMatches = directChildren(element, name);
  const matches = localMatches.filter((child) =>
    DOCX_MATH_NAMESPACES.has(child.namespaceURI ?? ''),
  );
  if (localMatches.length !== matches.length || matches.length > 1) return null;
  if (!matches.length) return required ? null : undefined;
  return matches[0];
}

function meaningfulAttributes(element: Element): Attr[] {
  return Array.from(element.attributes).filter(
    (attribute) =>
      xmlAttributeNamespace(element, attribute) !== XMLNS_NAMESPACE &&
      attribute.name !== 'xmlns' &&
      !attribute.name.startsWith('xmlns:'),
  );
}

function hasMeaningfulDirectText(element: Element): boolean {
  return Array.from(element.childNodes).some(
    (node) =>
      (node.nodeType === Node.TEXT_NODE ||
        node.nodeType === Node.CDATA_SECTION_NODE) &&
      Boolean(node.textContent?.trim()),
  );
}

function safeMathTextAttributes(element: Element): boolean {
  return meaningfulAttributes(element).every(
    (attribute) =>
      xmlAttributeNamespace(element, attribute) === XML_NAMESPACE &&
      xmlAttributeLocalName(attribute) === 'space' &&
      (attribute.value === 'default' || attribute.value === 'preserve'),
  );
}

function mathValue(element: Element): string | null {
  if (
    !DOCX_MATH_NAMESPACES.has(element.namespaceURI ?? '') ||
    directChildren(element).length ||
    hasMeaningfulDirectText(element)
  ) {
    return null;
  }
  const attributes = meaningfulAttributes(element);
  if (
    attributes.length !== 1 ||
    xmlAttributeLocalName(attributes[0]) !== 'val' ||
    !DOCX_MATH_NAMESPACES.has(
      xmlAttributeNamespace(element, attributes[0]) ?? '',
    )
  ) {
    return null;
  }
  return attributes[0]?.value ?? null;
}

function mathValueOrDefault(element: Element, fallback: string): string | null {
  if (
    !DOCX_MATH_NAMESPACES.has(element.namespaceURI ?? '') ||
    directChildren(element).length ||
    hasMeaningfulDirectText(element)
  ) {
    return null;
  }
  return meaningfulAttributes(element).length ? mathValue(element) : fallback;
}

function positiveMathInteger(value: string | null, maximum: number): number {
  if (!value || !/^[1-9]\d*$/u.test(value)) return 0;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed <= maximum ? parsed : 0;
}

function unsignedMathInteger(
  value: string | null,
  maximum: number,
): number | null {
  const source = value?.trim() ?? '';
  if (!/^\+?\d+$/u.test(source)) return null;
  const parsed = Number(source);
  return Number.isSafeInteger(parsed) && parsed <= maximum ? parsed : null;
}

function mathSpacingRule(
  element: Element | undefined,
): WorkDocumentEquationSpacingRule | null {
  const value = unsignedMathInteger(
    element ? mathValueOrDefault(element, '0') : '0',
    4,
  );
  return value === null
    ? null
    : ((['single', 'oneAndHalf', 'double', 'exact', 'multiple'] as const)[
        value
      ] ?? null);
}

function mathOnOff(element: Element): boolean | null {
  const attributes = meaningfulAttributes(element);
  if (!attributes.length) {
    return directChildren(element).length || hasMeaningfulDirectText(element)
      ? null
      : true;
  }
  const value = mathValue(element)?.trim().toLowerCase();
  if (value === '1' || value === 'on' || value === 'true') return true;
  if (value === '0' || value === 'off' || value === 'false') return false;
  return null;
}

function mathOnOffProperty(
  properties: Element | undefined,
  name: string,
): boolean | null {
  if (!properties) return false;
  const element = uniqueMathChild(properties, name, false);
  if (element === null) return null;
  return element ? mathOnOff(element) : false;
}

function parseManualBreak(element: Element): { alignmentAt?: number } | null {
  if (
    !DOCX_MATH_NAMESPACES.has(element.namespaceURI ?? '') ||
    directChildren(element).length ||
    hasMeaningfulDirectText(element)
  ) {
    return null;
  }
  const attributes = meaningfulAttributes(element);
  if (!attributes.length) return {};
  if (
    attributes.length !== 1 ||
    xmlAttributeLocalName(attributes[0]) !== 'alnAt' ||
    !DOCX_MATH_NAMESPACES.has(
      xmlAttributeNamespace(element, attributes[0]) ?? '',
    )
  ) {
    return null;
  }
  const source = attributes[0]?.value.trim() ?? '';
  if (!/^\+?\d+$/u.test(source)) return null;
  const alignmentAt = Number(source);
  return Number.isSafeInteger(alignmentAt) &&
    alignmentAt >= 1 &&
    alignmentAt <= 255
    ? { alignmentAt }
    : null;
}

function delimiterProperty(
  properties: Element | undefined,
  name: string,
  fallback: string,
): string | null {
  if (!properties) return fallback;
  const element = uniqueMathChild(properties, name, false);
  if (element === null) return null;
  const value = element ? mathValueOrDefault(element, '') : fallback;
  return value !== null && mathCharacter(value)
    ? value
    : value === ''
      ? ''
      : null;
}

function mathCharacter(value: string): boolean {
  return Array.from(value).length === 1 && !/[\p{Cc}\p{Cs}]/u.test(value);
}

function accentCharacter(value: string): boolean {
  if (Array.from(value).length !== 1) return false;
  const codePoint = value.codePointAt(0);
  return (
    codePoint !== undefined &&
    ((codePoint >= 0x0300 && codePoint <= 0x036f) ||
      (codePoint >= 0x20d0 && codePoint <= 0x20ef))
  );
}

function fractionTypeFromOmml(
  value: string | null,
): WorkDocumentEquationFractionType | null {
  if (value === 'bar') return 'bar';
  if (value === 'noBar') return 'noBar';
  if (value === 'skw') return 'skewed';
  if (value === 'lin') return 'linear';
  return null;
}

function topBottomFromOmml(value: string | null): 'top' | 'bottom' | null {
  if (value === 'top') return 'top';
  if (value === 'bot') return 'bottom';
  return null;
}

function mathJustificationFromOmml(
  value: string | null,
): WorkDocumentEquationJustification | null {
  if (
    value === 'left' ||
    value === 'right' ||
    value === 'center' ||
    value === 'centerGroup'
  ) {
    return value;
  }
  return null;
}

function runScriptFromOmml(
  value: string | null,
): WorkDocumentEquationRunScript | null {
  if (value === 'roman') return 'roman';
  if (value === 'sans-serif') return 'sansSerif';
  if (value === 'monospace') return 'monospace';
  if (value === 'fraktur') return 'fraktur';
  if (value === 'double-struck') return 'doubleStruck';
  if (value === 'script') return 'script';
  return null;
}

function runStyleFromOmml(
  value: string | null,
): WorkDocumentEquationRunStyle | null {
  if (value === 'p') return 'plain';
  if (value === 'i') return 'italic';
  if (value === 'b') return 'bold';
  if (value === 'bi') return 'boldItalic';
  return null;
}

function naryLimitLocation(
  operator: WorkDocumentEquationNaryOperator,
  value: string | null,
): WorkDocumentEquationLimitLocation | null {
  if (value === 'undOvr') return 'underOver';
  if (value === 'subSup') return 'subSup';
  if (value !== null) return null;
  return NARY_INTEGRALS.has(operator) ? 'subSup' : 'underOver';
}

function nextEquationMarker(state: EquationMarkerState): string {
  let marker = '';
  do {
    marker = `__A3S_WORK_EQUATION_${state.nextMarker}__`;
    state.nextMarker += 1;
  } while (state.occupiedText.includes(marker));
  state.occupiedText += marker;
  return marker;
}
