import {
  createDocumentEquationElement,
  documentEquationText,
  normalizeDocumentEquation,
  type WorkDocumentEquation,
  type WorkDocumentEquationExpression,
  type WorkDocumentEquationFractionType,
  type WorkDocumentEquationLimitLocation,
  type WorkDocumentEquationNaryOperator,
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

const TRANSITIONAL_MATH_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/math';
const STRICT_MATH_NAMESPACE = 'http://purl.oclc.org/ooxml/officeDocument/math';
const DOCX_MATH_NAMESPACES = new Set([
  TRANSITIONAL_MATH_NAMESPACE,
  STRICT_MATH_NAMESPACE,
]);

const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const MAX_IMPORTED_EQUATIONS = 4_096;
const MAX_MATH_DEPTH = 32;
const MAX_MATH_NODES = 4_096;
const MAX_MATH_TEXT_LENGTH = 65_536;
const MAX_DELIMITER_ARGUMENTS = 32;
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
const ARGUMENT_PROPERTY_NAMES = new Set(['argPr', 'ctrlPr']);
const STRUCTURAL_MATH_NAMES = new Set([
  'acc',
  'bar',
  'borderBox',
  'box',
  'd',
  'deg',
  'den',
  'e',
  'eqArr',
  'f',
  'fName',
  'func',
  'groupChr',
  'lim',
  'limLow',
  'limUpp',
  'm',
  'mr',
  'nary',
  'num',
  'oMath',
  'oMathPara',
  'phant',
  'r',
  'rad',
  'sPre',
  'sSub',
  'sSubSup',
  'sSup',
  'sub',
  'sup',
]);
const UNSAFE_PROPERTY_NAMES = new Set([
  'altChunk',
  'drawing',
  'fldChar',
  'instrText',
  'object',
  'oleObject',
  'pict',
]);
const RELATIONSHIP_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
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
  const texts = descendants(element, 't').filter((candidate) =>
    DOCX_MATH_NAMESPACES.has(candidate.namespaceURI ?? ''),
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
    if (!structuralChildren(element, new Set(['oMathParaPr', 'oMath']))) {
      return null;
    }
    const properties = mathDirectChildren(element, 'oMathParaPr');
    const equations = mathDirectChildren(element, 'oMath');
    if (
      properties.length > 1 ||
      equations.length !== 1 ||
      properties.some((property) => !emptyMathProperty(property))
    ) {
      return null;
    }
    const children = parseExpressionContainer(equations[0], state);
    return normalizedEquation('block', children);
  }
  const children = parseExpressionContainer(element, state);
  return normalizedEquation('inline', children);
}

function normalizedEquation(
  display: WorkDocumentEquation['display'],
  children: WorkDocumentEquationExpression[] | null,
): WorkDocumentEquation | null {
  return children
    ? normalizeDocumentEquation({ version: 1, display, children })
    : null;
}

function parseExpressionContainer(
  container: Element,
  state: EquationParseState,
  optional = false,
): WorkDocumentEquationExpression[] | null {
  if (
    !DOCX_MATH_NAMESPACES.has(container.namespaceURI ?? '') ||
    meaningfulAttributes(container).length ||
    hasMeaningfulDirectText(container)
  ) {
    return null;
  }
  const expressions: WorkDocumentEquationExpression[] = [];
  const properties = new Set<string>();
  for (const child of directChildren(container)) {
    if (!DOCX_MATH_NAMESPACES.has(child.namespaceURI ?? '')) return null;
    if (
      container.localName !== 'oMath' &&
      ARGUMENT_PROPERTY_NAMES.has(child.localName)
    ) {
      if (
        expressions.length ||
        properties.has(child.localName) ||
        !emptyMathProperty(child)
      ) {
        return null;
      }
      properties.add(child.localName);
      continue;
    }
    const expression = parseExpression(child, state);
    if (!expression) return null;
    expressions.push(expression);
  }
  return expressions.length || optional ? expressions : null;
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
    if (element.localName === 'f') return parseFraction(element, state);
    if (element.localName === 'sSup') return parseSuperScript(element, state);
    if (element.localName === 'sSub') return parseSubScript(element, state);
    if (element.localName === 'sSubSup') {
      return parseSubSuperScript(element, state);
    }
    if (element.localName === 'rad') return parseRadical(element, state);
    if (element.localName === 'func') return parseFunction(element, state);
    if (element.localName === 'nary') return parseNary(element, state);
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
  const textElements: Element[] = [];
  let sawProperties = false;
  for (const child of directChildren(element)) {
    if (child.localName === 'rPr') {
      if (sawProperties || textElements.length || !emptyMathProperty(child)) {
        return null;
      }
      sawProperties = true;
      continue;
    }
    if (
      child.localName !== 't' ||
      !DOCX_MATH_NAMESPACES.has(child.namespaceURI ?? '') ||
      directChildren(child).length ||
      !safeMathTextAttributes(child)
    ) {
      return null;
    }
    textElements.push(child);
  }
  if (textElements.length !== 1) return null;
  const text = textElements.map((child) => child.textContent ?? '').join('');
  if (!text || text.length > MAX_MATH_TEXT_LENGTH) return null;
  state.textLength += text.length;
  return state.textLength <= MAX_MATH_TEXT_LENGTH
    ? { type: 'run', text }
    : null;
}

function parseFraction(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  if (!structuralChildren(element, new Set(['fPr', 'num', 'den']))) {
    return null;
  }
  const properties = uniqueMathChild(element, 'fPr', false);
  const numerator = uniqueMathChild(element, 'num');
  const denominator = uniqueMathChild(element, 'den');
  if (
    properties === null ||
    !numerator ||
    !denominator ||
    (properties && !structuralChildren(properties, new Set(['type', 'ctrlPr'])))
  ) {
    return null;
  }
  const typeElement = properties
    ? uniqueMathChild(properties, 'type', false)
    : undefined;
  const controlProperties = properties
    ? uniqueMathChild(properties, 'ctrlPr', false)
    : undefined;
  if (
    typeElement === null ||
    controlProperties === null ||
    (controlProperties && !emptyMathProperty(controlProperties))
  ) {
    return null;
  }
  const sourceType = typeElement ? mathValue(typeElement) : 'bar';
  if (sourceType === null) return null;
  const fractionType = fractionTypeFromOmml(sourceType);
  const parsedNumerator = parseExpressionContainer(numerator, state);
  const parsedDenominator = parseExpressionContainer(denominator, state);
  return fractionType && parsedNumerator && parsedDenominator
    ? {
        type: 'fraction',
        fractionType,
        numerator: parsedNumerator,
        denominator: parsedDenominator,
      }
    : null;
}

function parseSuperScript(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  if (!structuralChildren(element, new Set(['sSupPr', 'e', 'sup']))) {
    return null;
  }
  const properties = uniqueMathChild(element, 'sSupPr', false);
  const base = uniqueMathChild(element, 'e');
  const superScript = uniqueMathChild(element, 'sup');
  const parsedBase = base ? parseExpressionContainer(base, state) : null;
  const parsedSuperScript = superScript
    ? parseExpressionContainer(superScript, state)
    : null;
  return properties !== null &&
    (!properties || emptyMathProperty(properties)) &&
    parsedBase &&
    parsedSuperScript
    ? { type: 'superscript', base: parsedBase, superScript: parsedSuperScript }
    : null;
}

function parseSubScript(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  if (!structuralChildren(element, new Set(['sSubPr', 'e', 'sub']))) {
    return null;
  }
  const properties = uniqueMathChild(element, 'sSubPr', false);
  const base = uniqueMathChild(element, 'e');
  const subScript = uniqueMathChild(element, 'sub');
  const parsedBase = base ? parseExpressionContainer(base, state) : null;
  const parsedSubScript = subScript
    ? parseExpressionContainer(subScript, state)
    : null;
  return properties !== null &&
    (!properties || emptyMathProperty(properties)) &&
    parsedBase &&
    parsedSubScript
    ? { type: 'subscript', base: parsedBase, subScript: parsedSubScript }
    : null;
}

function parseSubSuperScript(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  if (!structuralChildren(element, new Set(['sSubSupPr', 'e', 'sub', 'sup']))) {
    return null;
  }
  const properties = uniqueMathChild(element, 'sSubSupPr', false);
  const base = uniqueMathChild(element, 'e');
  const subScript = uniqueMathChild(element, 'sub');
  const superScript = uniqueMathChild(element, 'sup');
  const parsedBase = base ? parseExpressionContainer(base, state) : null;
  const parsedSubScript = subScript
    ? parseExpressionContainer(subScript, state)
    : null;
  const parsedSuperScript = superScript
    ? parseExpressionContainer(superScript, state)
    : null;
  return properties !== null &&
    (!properties || emptyMathProperty(properties)) &&
    parsedBase &&
    parsedSubScript &&
    parsedSuperScript
    ? {
        type: 'subSuperScript',
        base: parsedBase,
        subScript: parsedSubScript,
        superScript: parsedSuperScript,
      }
    : null;
}

function parseRadical(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  if (!structuralChildren(element, new Set(['radPr', 'deg', 'e']))) {
    return null;
  }
  const properties = uniqueMathChild(element, 'radPr', false);
  const degree = uniqueMathChild(element, 'deg', false);
  const body = uniqueMathChild(element, 'e');
  if (
    properties === null ||
    degree === null ||
    !body ||
    (properties &&
      !structuralChildren(properties, new Set(['degHide', 'ctrlPr'])))
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
  if (
    degreeHiddenElement === null ||
    controlProperties === null ||
    degreeHidden === null ||
    (controlProperties && !emptyMathProperty(controlProperties))
  ) {
    return null;
  }
  const parsedBody = parseExpressionContainer(body, state);
  const parsedDegree = degree
    ? parseExpressionContainer(degree, state, true)
    : undefined;
  if (
    !parsedBody ||
    parsedDegree === null ||
    degreeHidden === Boolean(parsedDegree?.length)
  ) {
    return null;
  }
  return {
    type: 'radical',
    children: parsedBody,
    ...(parsedDegree?.length ? { degree: parsedDegree } : {}),
  };
}

function parseFunction(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  if (!structuralChildren(element, new Set(['funcPr', 'fName', 'e']))) {
    return null;
  }
  const properties = uniqueMathChild(element, 'funcPr', false);
  const name = uniqueMathChild(element, 'fName');
  const body = uniqueMathChild(element, 'e');
  const parsedName = name ? parseExpressionContainer(name, state) : null;
  const parsedBody = body ? parseExpressionContainer(body, state) : null;
  return properties !== null &&
    (!properties || emptyMathProperty(properties)) &&
    parsedName &&
    parsedBody
    ? { type: 'function', name: parsedName, children: parsedBody }
    : null;
}

function parseNary(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  if (!structuralChildren(element, new Set(['naryPr', 'sub', 'sup', 'e']))) {
    return null;
  }
  const properties = uniqueMathChild(element, 'naryPr', false);
  const subScript = uniqueMathChild(element, 'sub', false);
  const superScript = uniqueMathChild(element, 'sup', false);
  const body = uniqueMathChild(element, 'e');
  if (
    properties === null ||
    subScript === null ||
    superScript === null ||
    !body ||
    (properties &&
      !structuralChildren(
        properties,
        new Set(['chr', 'limLoc', 'subHide', 'supHide', 'ctrlPr']),
      ))
  ) {
    return null;
  }
  const operatorElement = properties
    ? uniqueMathChild(properties, 'chr', false)
    : undefined;
  const locationElement = properties
    ? uniqueMathChild(properties, 'limLoc', false)
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
  if (
    operatorElement === null ||
    locationElement === null ||
    subHiddenElement === null ||
    superHiddenElement === null ||
    controlProperties === null ||
    (controlProperties && !emptyMathProperty(controlProperties))
  ) {
    return null;
  }
  const operatorValue = operatorElement ? mathValue(operatorElement) : null;
  if (!NARY_OPERATORS.has(operatorValue as WorkDocumentEquationNaryOperator)) {
    return null;
  }
  const operator = operatorValue as WorkDocumentEquationNaryOperator;
  const locationValue = locationElement ? mathValue(locationElement) : null;
  if (locationElement && locationValue === null) return null;
  const limitLocation = naryLimitLocation(operator, locationValue);
  const subHidden = subHiddenElement ? mathOnOff(subHiddenElement) : false;
  const superHidden = superHiddenElement
    ? mathOnOff(superHiddenElement)
    : false;
  const parsedBody = parseExpressionContainer(body, state);
  const parsedSubScript = subScript
    ? parseExpressionContainer(subScript, state, true)
    : undefined;
  const parsedSuperScript = superScript
    ? parseExpressionContainer(superScript, state, true)
    : undefined;
  if (
    !limitLocation ||
    subHidden === null ||
    superHidden === null ||
    !parsedBody ||
    parsedSubScript === null ||
    parsedSuperScript === null ||
    subHidden === Boolean(parsedSubScript?.length) ||
    superHidden === Boolean(parsedSuperScript?.length)
  ) {
    return null;
  }
  return {
    type: 'nary',
    operator,
    limitLocation,
    children: parsedBody,
    ...(parsedSubScript?.length ? { subScript: parsedSubScript } : {}),
    ...(parsedSuperScript?.length ? { superScript: parsedSuperScript } : {}),
  };
}

function parseDelimiter(
  element: Element,
  state: EquationParseState,
): WorkDocumentEquationExpression | null {
  if (!structuralChildren(element, new Set(['dPr', 'e']))) return null;
  const properties = uniqueMathChild(element, 'dPr', false);
  const arguments_ = mathDirectChildren(element, 'e');
  if (
    properties === null ||
    !arguments_.length ||
    arguments_.length > MAX_DELIMITER_ARGUMENTS ||
    (properties &&
      !structuralChildren(
        properties,
        new Set(['begChr', 'endChr', 'sepChr', 'ctrlPr']),
      ))
  ) {
    return null;
  }
  const controlProperties = properties
    ? uniqueMathChild(properties, 'ctrlPr', false)
    : undefined;
  if (
    controlProperties === null ||
    (controlProperties && !emptyMathProperty(controlProperties))
  ) {
    return null;
  }
  const opening = delimiterProperty(properties, 'begChr', '(');
  const closing = delimiterProperty(properties, 'endChr', ')');
  const separator = delimiterProperty(properties, 'sepChr', '|');
  const parsedArguments = arguments_.map((argument) =>
    parseExpressionContainer(argument, state),
  );
  return opening !== null &&
    closing !== null &&
    separator !== null &&
    parsedArguments.every(
      (argument): argument is WorkDocumentEquationExpression[] =>
        Boolean(argument),
    )
    ? {
        type: 'delimiter',
        opening,
        closing,
        separator,
        arguments: parsedArguments,
      }
    : null;
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

function safePassiveProperty(element: Element): boolean {
  if (!DOCX_MATH_NAMESPACES.has(element.namespaceURI ?? '')) return false;
  return [element, ...Array.from(element.querySelectorAll('*'))].every(
    (candidate) =>
      !UNSAFE_PROPERTY_NAMES.has(candidate.localName) &&
      (candidate === element ||
        !STRUCTURAL_MATH_NAMES.has(candidate.localName)) &&
      (DOCX_MATH_NAMESPACES.has(candidate.namespaceURI ?? '') ||
        DOCX_WORDPROCESSING_NAMESPACES.has(candidate.namespaceURI ?? '')) &&
      !hasUnsafeAttributes(candidate),
  );
}

function emptyMathProperty(element: Element): boolean {
  return (
    meaningfulAttributes(element).length === 0 &&
    directChildren(element).length === 0 &&
    !hasMeaningfulDirectText(element) &&
    safePassiveProperty(element)
  );
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

function hasUnsafeAttributes(element: Element): boolean {
  return Array.from(element.attributes).some((attribute) => {
    const localName = xmlAttributeLocalName(attribute);
    const namespace = xmlAttributeNamespace(element, attribute);
    return (
      RELATIONSHIP_NAMESPACES.has(namespace ?? '') ||
      (/^(?:embed|href|id|link)$/i.test(localName) &&
        Boolean(attribute.value.trim()))
    );
  });
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

function delimiterProperty(
  properties: Element | undefined,
  name: string,
  fallback: string,
): string | null {
  if (!properties) return fallback;
  const element = uniqueMathChild(properties, name, false);
  if (element === null) return null;
  const value = element ? mathValue(element) : fallback;
  return value !== null && mathCharacter(value)
    ? value
    : value === ''
      ? ''
      : null;
}

function mathCharacter(value: string): boolean {
  return Array.from(value).length === 1 && !/[\p{Cc}\p{Cs}]/u.test(value);
}

function fractionTypeFromOmml(
  value: string | null,
): WorkDocumentEquationFractionType | null {
  if (value === 'bar' || value === null) return 'bar';
  if (value === 'noBar') return 'noBar';
  if (value === 'skw') return 'skewed';
  if (value === 'lin') return 'linear';
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
