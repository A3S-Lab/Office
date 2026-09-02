import {
  normalizeDocumentTextBoxProperties,
  textBoxCss,
  textBoxDomAttributes,
  type WorkDocumentTextBoxProperties,
} from './work-document-text-box';
import {
  attribute,
  descendants,
  directChild,
  directChildren,
} from './work-ooxml-package';

export interface ImportedDocxTextBoxMarker {
  marker: string;
  properties: WorkDocumentTextBoxProperties;
}

export interface ImportedDocxTextBoxMarkers {
  textBoxes: ImportedDocxTextBoxMarker[];
}

export interface DocxTextBoxInspection {
  supported: number;
  unsupported: number;
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const WORDPROCESSING_DRAWING_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  'http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing',
]);
const DRAWING_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/drawingml/2006/main',
  'http://purl.oclc.org/ooxml/drawingml/main',
]);
const WORDPROCESSING_SHAPE_NAMESPACES = new Set([
  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
]);
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const WORDPROCESSING_SHAPE_URI =
  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const EMUS_PER_MILLIMETER = 36_000;
const MAX_IMPORTED_TEXT_BOXES = 1_024;
const TEXT_BOX_MARKER_PATTERN = /__A3S_WORK_TEXT_BOX_\d+__/g;
const TEXT_BOX_RUN_CONTENT_NAMES = new Set([
  'br',
  'cr',
  'noBreakHyphen',
  'softHyphen',
  'sym',
  'tab',
  't',
]);

/**
 * Rewrites supported WPS text boxes into ordinary Word runs for Mammoth.
 * Mammoth intentionally ignores `wps:txbxContent`; the marker is consumed
 * after conversion and replaced with an editable HTML block.
 */
export function markDocxTextBoxes(
  document: Document,
): ImportedDocxTextBoxMarkers {
  const textBoxes: ImportedDocxTextBoxMarker[] = [];
  const drawings = docxDrawings(document);
  let nextMarker = 1;
  for (const drawing of drawings) {
    const shape = supportedTextBoxShape(drawing);
    if (!shape) continue;
    const paragraph = closestAncestor(drawing, 'p');
    const run = closestAncestor(drawing, 'r');
    if (!paragraph || !run || !isSoleDrawingParagraph(paragraph, run)) {
      continue;
    }
    if (textBoxes.length >= MAX_IMPORTED_TEXT_BOXES) {
      throw new Error('Imported DOCX exceeds the text-box limit.');
    }
    const marker = nextTextBoxMarker(document, nextMarker);
    nextMarker += 1;
    const properties = textBoxProperties(drawing, shape, textBoxes.length + 1);
    replaceParagraphWithTextBoxRuns(document, paragraph, shape.content, marker);
    textBoxes.push({ marker, properties });
  }
  return { textBoxes };
}

/**
 * Inspects WPS text-box shapes without mutating the source package. The
 * importer deliberately supports only a shape that is the sole drawing in a
 * paragraph; mixed paragraphs remain on Mammoth's normal compatibility path.
 */
export function inspectDocxTextBoxes(
  document: Document,
): DocxTextBoxInspection {
  let supported = 0;
  let unsupported = 0;
  for (const drawing of docxDrawings(document)) {
    // Count every explicitly declared text-box shape, including malformed
    // bodies, so diagnostics can explain why the native path was skipped.
    if (!textBoxShapeCandidate(drawing)) continue;
    const shape = supportedTextBoxShape(drawing);
    if (!shape) {
      unsupported += 1;
      continue;
    }
    const paragraph = closestAncestor(drawing, 'p');
    const run = closestAncestor(drawing, 'r');
    if (paragraph && run && isSoleDrawingParagraph(paragraph, run)) {
      supported += 1;
    } else {
      unsupported += 1;
    }
  }
  return { supported, unsupported };
}

export function applyImportedDocxTextBoxMarkers(
  document: Document,
  markers: ImportedDocxTextBoxMarkers,
): void {
  if (!markers.textBoxes.length) return;
  const byMarker = new Map(
    markers.textBoxes.map((textBox) => [textBox.marker, textBox]),
  );
  const nodes = textNodes(document.body);
  for (const node of nodes) {
    if (!node.data.includes('__A3S_WORK_TEXT_BOX_')) continue;
    const match = node.data
      .match(TEXT_BOX_MARKER_PATTERN)
      ?.find((value) => byMarker.has(value));
    if (!match) continue;
    const textBox = byMarker.get(match);
    if (!textBox) continue;
    node.data = node.data.replace(match, '');
    const block = closestHtmlBlock(node.parentElement);
    if (!block) continue;
    const wrapper = document.createElement('div');
    for (const attribute of Array.from(block.attributes)) {
      // Keep paragraph identities and safe marker attributes. Mammoth class
      // names and style projections are not shape semantics.
      if (
        attribute.name.startsWith('data-office-') ||
        attribute.name.startsWith('data-document-')
      ) {
        wrapper.setAttribute(attribute.name, attribute.value);
      }
    }
    for (const [name, value] of Object.entries(
      textBoxDomAttributes(textBox.properties),
    )) {
      if (value !== undefined) wrapper.setAttribute(name, value);
    }
    wrapper.className = 'work-document-text-box';
    wrapper.setAttribute('role', 'textbox');
    wrapper.setAttribute('aria-label', '文本框');
    wrapper.setAttribute('style', textBoxCss(textBox.properties));
    while (block.firstChild) wrapper.append(block.firstChild);
    block.replaceWith(wrapper);
  }
  document.body.normalize();
}

export function hasImportedDocxTextBoxMarkers(
  markers: ImportedDocxTextBoxMarkers,
): boolean {
  return markers.textBoxes.length > 0;
}

interface TextBoxShape {
  shape: Element;
  content: Element;
}

/**
 * Recognize the explicit WPS text-box declaration independently from its
 * content. This lets the compatibility layer report malformed or missing
 * `txbxContent` instead of silently classifying the drawing as an image.
 */
function textBoxShapeCandidate(drawing: Element): Element | null {
  if (!WORDPROCESSING_DRAWING_NAMESPACES.has(drawing.namespaceURI ?? '')) {
    return null;
  }
  const graphicData = descendants(drawing, 'graphicData').find(
    (element) =>
      DRAWING_NAMESPACES.has(element.namespaceURI ?? '') &&
      attribute(element, 'uri') === WORDPROCESSING_SHAPE_URI,
  );
  if (!graphicData) return null;
  const shape = directChild(graphicData, 'wsp');
  if (!shape || !WORDPROCESSING_SHAPE_NAMESPACES.has(shape.namespaceURI ?? ''))
    return null;
  const nonVisual = directChild(shape, 'cNvSpPr');
  if (
    nonVisual &&
    !WORDPROCESSING_SHAPE_NAMESPACES.has(nonVisual.namespaceURI ?? '')
  ) {
    return null;
  }
  const txBox = attribute(nonVisual ?? shape, 'txBox')
    ?.trim()
    .toLowerCase();
  if (txBox !== '1' && txBox !== 'true') return null;
  return shape;
}

function supportedTextBoxShape(drawing: Element): TextBoxShape | null {
  const shape = textBoxShapeCandidate(drawing);
  if (!shape) return null;
  const txbx = directChild(shape, 'txbx');
  if (txbx && !WORDPROCESSING_SHAPE_NAMESPACES.has(txbx.namespaceURI ?? '')) {
    return null;
  }
  const content = txbx ? directChild(txbx, 'txbxContent') : undefined;
  return content &&
    isWordElement(content, 'txbxContent') &&
    isSupportedTextBoxContent(content)
    ? { shape, content }
    : null;
}

/**
 * The marker rewrite intentionally accepts only paragraph/run inline content.
 * Other children (tables, fields, nested drawings, hyperlinks, and revision
 * containers) would otherwise be dropped when the inner body is projected to
 * Mammoth runs, so they remain on the explicit compatibility path.
 */
function isSupportedTextBoxContent(content: Element): boolean {
  const paragraphs = directChildren(content);
  if (!paragraphs.length) return false;
  for (const paragraph of paragraphs) {
    if (!isWordElement(paragraph, 'p')) return false;
    const children = directChildren(paragraph).filter(
      (element) => element.localName !== 'pPr',
    );
    for (const child of children) {
      if (!isWordElement(child, 'r')) return false;
      const runChildren = directChildren(child).filter(
        (element) => element.localName !== 'rPr',
      );
      if (
        runChildren.some(
          (element) =>
            !isWordNamespace(element) ||
            !TEXT_BOX_RUN_CONTENT_NAMES.has(element.localName),
        )
      ) {
        return false;
      }
    }
  }
  return true;
}

function isWordElement(element: Element, localName: string): boolean {
  return element.localName === localName && isWordNamespace(element);
}

function isWordNamespace(element: Element): boolean {
  return (
    element.namespaceURI === WORD_NAMESPACE ||
    element.namespaceURI === STRICT_WORD_NAMESPACE
  );
}

function docxDrawings(document: Document): Element[] {
  return [
    ...descendants(document, 'anchor'),
    ...descendants(document, 'inline'),
  ].filter((element) =>
    WORDPROCESSING_DRAWING_NAMESPACES.has(element.namespaceURI ?? ''),
  );
}

function isSoleDrawingParagraph(paragraph: Element, run: Element): boolean {
  const meaningful = directChildren(paragraph).filter(
    (element) => element.localName !== 'pPr',
  );
  if (meaningful.length !== 1 || meaningful[0] !== run) return false;
  const runContent = directChildren(run).filter(
    (element) => element.localName !== 'rPr',
  );
  return runContent.length === 1 && runContent[0]?.localName === 'drawing';
}

function replaceParagraphWithTextBoxRuns(
  document: Document,
  paragraph: Element,
  content: Element,
  marker: string,
): void {
  const namespace = paragraph.namespaceURI || WORD_NAMESPACE;
  const prefix = paragraph.prefix ? `${paragraph.prefix}:` : 'w:';
  const paragraphProperties = directChild(paragraph, 'pPr');
  const runs: Element[] = [];
  const sourceParagraphs = directChildren(content, 'p');
  for (const [paragraphIndex, sourceParagraph] of sourceParagraphs.entries()) {
    if (paragraphIndex > 0)
      runs.push(createBreakRun(document, namespace, prefix));
    const sourceRuns = directChildren(sourceParagraph, 'r');
    for (const sourceRun of sourceRuns) {
      runs.push(document.importNode(sourceRun, true) as Element);
    }
  }
  const markerRun = createTextRun(document, namespace, prefix, marker);
  paragraph.replaceChildren(
    ...(paragraphProperties ? [paragraphProperties] : []),
    markerRun,
    ...runs,
  );
}

function createTextRun(
  document: Document,
  namespace: string,
  prefix: string,
  textValue: string,
): Element {
  const run = document.createElementNS(namespace, `${prefix}r`);
  const text = document.createElementNS(namespace, `${prefix}t`);
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = textValue;
  run.append(text);
  return run;
}

function createBreakRun(
  document: Document,
  namespace: string,
  prefix: string,
): Element {
  const run = document.createElementNS(namespace, `${prefix}r`);
  run.append(document.createElementNS(namespace, `${prefix}br`));
  return run;
}

function textBoxProperties(
  drawing: Element,
  shape: TextBoxShape,
  index: number,
): WorkDocumentTextBoxProperties {
  const extent = directChild(drawing, 'extent');
  const shapeProperties = directChild(shape.shape, 'spPr');
  const transform = shapeProperties
    ? directChild(shapeProperties, 'xfrm')
    : undefined;
  const extension = directChild(
    transform ?? shapeProperties ?? shape.shape,
    'ext',
  );
  const width = emuMillimeters(
    attribute(extent ?? extension ?? shape.shape, 'cx') ??
      attribute(extent ?? extension ?? shape.shape, 'x'),
    120,
  );
  const height = emuMillimeters(
    attribute(extent ?? extension ?? shape.shape, 'cy') ??
      attribute(extent ?? extension ?? shape.shape, 'y'),
    45,
  );
  const docProperties = directChild(drawing, 'docPr');
  const docPropertiesId = integerAttribute(docProperties, 'id');
  const sourceName = attributeOptional(docProperties, 'name')?.trim() ?? '';
  const id = sourceName.startsWith('A3S Text Box ')
    ? sourceName.slice('A3S Text Box '.length).trim()
    : docPropertiesId === null
      ? `docx-text-box-${index}`
      : `docx-text-box-${docPropertiesId}`;
  const bodyProperties = directChild(shape.shape, 'bodyPr');
  const horizontal = directChild(drawing, 'positionH');
  const vertical = directChild(drawing, 'positionV');
  const layout = drawing.localName === 'anchor' ? 'floating' : 'inline';
  const fill = shapeFill(shapeProperties);
  const border = shapeBorder(shapeProperties);
  return normalizeDocumentTextBoxProperties({
    id,
    width,
    height,
    layout,
    horizontalOffset: positionOffset(horizontal),
    verticalOffset: positionOffset(vertical),
    horizontalReference: positionReference(horizontal, 'column'),
    verticalReference: positionReference(vertical, 'paragraph'),
    fill,
    borderColor: border.color,
    borderWidth: border.width,
    padding: bodyPadding(bodyProperties),
    verticalAlign: bodyVerticalAlign(bodyProperties),
    docPropertiesId,
  });
}

function shapeFill(shapeProperties: Element | undefined): string {
  if (!shapeProperties || directChild(shapeProperties, 'noFill')) {
    return 'transparent';
  }
  const fill = directChild(shapeProperties, 'solidFill');
  const rgb = fill ? directChild(fill, 'srgbClr') : undefined;
  const value = attributeOptional(rgb, 'val')?.trim();
  return value && /^[0-9a-f]{6}$/i.test(value) ? `#${value}` : 'transparent';
}

function shapeBorder(shapeProperties: Element | undefined): {
  color: string;
  width: number;
} {
  const line = shapeProperties ? directChild(shapeProperties, 'ln') : undefined;
  if (!line || directChild(line, 'noFill')) {
    return { color: 'none', width: 0 };
  }
  const fill = directChild(line, 'solidFill');
  const rgb = fill ? directChild(fill, 'srgbClr') : undefined;
  const value = attributeOptional(rgb, 'val')?.trim();
  return {
    color: value && /^[0-9a-f]{6}$/i.test(value) ? `#${value}` : '#4472c4',
    width: emuMillimeters(attribute(line, 'w'), 0.35),
  };
}

function bodyPadding(bodyProperties: Element | undefined): number {
  if (!bodyProperties) return 3;
  const values = ['lIns', 'rIns', 'tIns', 'bIns'].map((name) =>
    emuMillimeters(attributeOptional(bodyProperties, name), 3),
  );
  return Math.min(...values);
}

function bodyVerticalAlign(
  bodyProperties: Element | undefined,
): 'top' | 'center' | 'bottom' {
  const value = attributeOptional(bodyProperties, 'anchor');
  if (value === 'ctr') return 'center';
  if (value === 'b') return 'bottom';
  return 'top';
}

function positionOffset(element: Element | undefined): number | null {
  const offset = element ? directChild(element, 'posOffset') : undefined;
  if (!offset?.textContent?.trim()) return null;
  const value = Number(offset.textContent);
  return Number.isFinite(value) ? value / EMUS_PER_MILLIMETER : null;
}

function positionReference(
  element: Element | undefined,
  fallback: 'column' | 'margin' | 'page' | 'paragraph',
): 'column' | 'margin' | 'page' | 'paragraph' {
  const value = attributeOptional(element, 'relativeFrom');
  if (value === 'margin' || value === 'page' || value === 'column')
    return value;
  if (value === 'paragraph') return value;
  return fallback;
}

function emuMillimeters(value: string | null, fallback: number): number {
  const number = value === null ? Number.NaN : Number(value);
  return Number.isFinite(number) && number >= 0
    ? number / EMUS_PER_MILLIMETER
    : fallback;
}

function integerAttribute(
  element: Element | undefined,
  name: string,
): number | null {
  if (!element) return null;
  const value = Number(attributeOptional(element, name));
  return Number.isSafeInteger(value) && value >= 0 ? value : null;
}

function attributeOptional(
  element: Element | undefined,
  name: string,
): string | null {
  return element ? attribute(element, name) : null;
}

function nextTextBoxMarker(document: Document, start: number): string {
  let index = start;
  while (
    document.documentElement.textContent?.includes(
      `__A3S_WORK_TEXT_BOX_${index}__`,
    )
  ) {
    index += 1;
  }
  return `__A3S_WORK_TEXT_BOX_${index}__`;
}

function closestAncestor(element: Element, localName: string): Element | null {
  let current: Element | null = element;
  while (current) {
    if (current.localName === localName) return current;
    current = current.parentElement;
  }
  return null;
}

function closestHtmlBlock(element: Element | null): HTMLElement | null {
  if (!element) return null;
  const block = element.closest(
    'p, h1, h2, h3, h4, h5, h6, li, blockquote, td, th',
  );
  return block instanceof HTMLElement ? block : null;
}

function textNodes(root: ParentNode): Text[] {
  const walker = root.ownerDocument?.createTreeWalker(
    root,
    NodeFilter.SHOW_TEXT,
  );
  const nodes: Text[] = [];
  if (!walker) return nodes;
  while (walker.nextNode()) nodes.push(walker.currentNode as Text);
  return nodes;
}
