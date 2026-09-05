import {
  connectorCss,
  connectorDomAttributes,
  normalizeDocumentConnectorProperties,
  type WorkDocumentConnectorArrow,
  type WorkDocumentConnectorLineStyle,
  type WorkDocumentConnectorProperties,
} from './work-document-connector';
import {
  attribute,
  descendants,
  directChild,
  directChildren,
} from './work-ooxml-package';

export interface ImportedDocxConnectorMarker {
  marker: string;
  properties: WorkDocumentConnectorProperties;
}

export interface ImportedDocxConnectorMarkers {
  connectors: ImportedDocxConnectorMarker[];
}

export interface DocxConnectorInspection {
  supported: number;
  unsupported: number;
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const VML_NAMESPACE = 'urn:schemas-microsoft-com:vml';
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
const WORDPROCESSING_SHAPE_URI =
  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const EMUS_PER_MILLIMETER = 36_000;
const POINTS_PER_MILLIMETER = 72 / 25.4;
const MAX_IMPORTED_CONNECTORS = 1_024;
const CONNECTOR_MARKER_PATTERN = /__A3S_WORK_CONNECTOR_\d+__/g;

/**
 * Rewrites the supported straight connector subset to a Mammoth-safe marker.
 * VML connectors are the native WPS 12 shape emitted by Shapes.AddConnector;
 * the DrawingML branch is the representation emitted by A3S on export.
 */
export function markDocxConnectors(
  document: Document,
): ImportedDocxConnectorMarkers {
  const connectors: ImportedDocxConnectorMarker[] = [];
  const claimedRuns = new Set<Element>();
  let nextMarker = 1;

  for (const candidate of connectorCandidates(document)) {
    const paragraph = closestAncestor(candidate.element, 'p');
    const run = closestAncestor(candidate.element, 'r');
    if (!paragraph || !run || claimedRuns.has(run)) continue;
    if (!isSoleDrawingParagraph(paragraph, run)) continue;
    if (connectors.length >= MAX_IMPORTED_CONNECTORS) {
      throw new Error('Imported DOCX exceeds the connector limit.');
    }
    const marker = nextConnectorMarker(document, nextMarker);
    nextMarker += 1;
    replaceDrawingRunWithConnectorMarker(document, run, marker);
    connectors.push({
      marker,
      properties: connectorProperties(
        candidate.element,
        candidate.kind,
        connectors.length + 1,
      ),
    });
    claimedRuns.add(run);
  }
  return { connectors };
}

/** Inspects connector declarations without mutating the source package. */
export function inspectDocxConnectorShapes(
  document: Document,
): DocxConnectorInspection {
  let supported = 0;
  let unsupported = 0;
  for (const candidate of connectorCandidates(document)) {
    const paragraph = closestAncestor(candidate.element, 'p');
    const run = closestAncestor(candidate.element, 'r');
    if (
      isSoleDrawingParagraph(paragraph, run) &&
      connectorProperties(candidate.element, candidate.kind, supported + 1)
    ) {
      supported += 1;
    } else {
      unsupported += 1;
    }
  }
  return { supported, unsupported };
}

export function applyImportedDocxConnectorMarkers(
  document: Document,
  markers: ImportedDocxConnectorMarkers,
): void {
  if (!markers.connectors.length) return;
  const byMarker = new Map(
    markers.connectors.map((connector) => [connector.marker, connector]),
  );
  const nodes = textNodes(document.body);
  for (const node of nodes) {
    if (!node.data.includes('__A3S_WORK_CONNECTOR_')) continue;
    const match = node.data
      .match(CONNECTOR_MARKER_PATTERN)
      ?.find((value) => byMarker.has(value));
    if (!match) continue;
    const connector = byMarker.get(match);
    if (!connector) continue;
    node.data = node.data.replace(match, '');
    const block = closestHtmlBlock(node.parentElement);
    if (!block) continue;
    const wrapper = document.createElement('div');
    for (const attribute of Array.from(block.attributes)) {
      if (
        attribute.name.startsWith('data-office-') ||
        attribute.name.startsWith('data-document-')
      ) {
        wrapper.setAttribute(attribute.name, attribute.value);
      }
    }
    for (const [name, value] of Object.entries(
      connectorDomAttributes(connector.properties),
    )) {
      if (value !== undefined) wrapper.setAttribute(name, value);
    }
    wrapper.className = 'work-document-connector';
    wrapper.setAttribute('role', 'img');
    wrapper.setAttribute('aria-label', '直线连接符');
    wrapper.setAttribute('style', connectorCss(connector.properties));
    block.replaceWith(wrapper);
  }
  document.body.normalize();
}

export function hasImportedDocxConnectorMarkers(
  markers: ImportedDocxConnectorMarkers,
): boolean {
  return markers.connectors.length > 0;
}

type ConnectorCandidateKind = 'vml' | 'drawingml';
interface ConnectorCandidate {
  element: Element;
  kind: ConnectorCandidateKind;
}

function connectorCandidates(document: Document): ConnectorCandidate[] {
  const drawingCandidates = descendants(document, 'anchor')
    .concat(descendants(document, 'inline'))
    .filter((element) =>
      WORDPROCESSING_DRAWING_NAMESPACES.has(element.namespaceURI ?? ''),
    )
    .filter((element) => isDrawingMlConnector(element))
    .map((element) => ({ element, kind: 'drawingml' as const }));
  const drawingRuns = new Set(
    drawingCandidates
      .map(({ element }) => closestAncestor(element, 'r'))
      .filter((value): value is Element => Boolean(value)),
  );
  const vmlCandidates = descendants(document, 'pict')
    .filter((element) => isWordNamespace(element))
    .flatMap((pict) => {
      const shape = descendants(pict, 'shape').find(isVmlConnector);
      const line = descendants(pict, 'line').find(isVmlConnector);
      const element = shape ?? line;
      if (!element || drawingRuns.has(closestAncestor(pict, 'r') as Element)) {
        return [];
      }
      return [{ element, kind: 'vml' as const }];
    });
  return [...drawingCandidates, ...vmlCandidates];
}

function isDrawingMlConnector(drawing: Element): boolean {
  const graphicData = descendants(drawing, 'graphicData').find(
    (element) =>
      DRAWING_NAMESPACES.has(element.namespaceURI ?? '') &&
      attribute(element, 'uri') === WORDPROCESSING_SHAPE_URI,
  );
  const shape = graphicData && directChild(graphicData, 'wsp');
  if (
    !shape ||
    !WORDPROCESSING_SHAPE_NAMESPACES.has(shape.namespaceURI ?? '')
  ) {
    return false;
  }
  const nonVisual = shape ? directChild(shape, 'cNvSpPr') : undefined;
  if (attribute(nonVisual ?? shape, 'txBox')) return false;
  const properties = shape ? directChild(shape, 'spPr') : undefined;
  if (!properties) return false;
  const preset = properties ? directChild(properties, 'prstGeom') : undefined;
  if (attribute(preset ?? shape, 'prst') === 'line') return true;
  return Boolean(properties && directChild(properties, 'custGeom'));
}

function isVmlConnector(element: Element): boolean {
  if (element.namespaceURI !== VML_NAMESPACE) return false;
  if (element.localName === 'line') return true;
  if (element.localName !== 'shape') return false;
  const shapeType = attribute(element, 'spt')?.trim();
  const type = attribute(element, 'type')?.trim().toLowerCase();
  return shapeType === '32' || type === '#_x0000_t32';
}

function connectorProperties(
  element: Element,
  kind: ConnectorCandidateKind,
  index: number,
): WorkDocumentConnectorProperties {
  return kind === 'vml'
    ? vmlConnectorProperties(element, index)
    : drawingMlConnectorProperties(element, index);
}

function vmlConnectorProperties(
  element: Element,
  index: number,
): WorkDocumentConnectorProperties {
  const style = parseVmlStyle(attribute(element, 'style'));
  const width = vmlLength(style.width, 50.8);
  const height = vmlLength(style.height, width > 0 ? 0.35 : 25);
  const path = directChild(element, 'path');
  const line = directChild(element, 'line');
  const coordinateSize = parseCoordinateSize(attribute(element, 'coordsize'));
  const endpoints = parseVmlEndpoints(
    line ?? path,
    coordinateSize,
    width,
    height,
  );
  const stroke = directChild(element, 'stroke');
  const color = normalizeConnectorColor(
    (stroke ? attribute(stroke, 'color') : null) ?? style.strokecolor,
  );
  const lineWidth = vmlLength(
    stroke ? attribute(stroke, 'weight') : null,
    0.35,
  );
  return normalizeDocumentConnectorProperties({
    id: attribute(element, 'id')?.trim() || `docx-connector-${index}`,
    width,
    height,
    layout: style.position === 'static' ? 'inline' : 'floating',
    horizontalOffset: vmlOffset(style.left, style['margin-left']),
    verticalOffset: vmlOffset(style.top, style['margin-top']),
    horizontalReference: 'page',
    verticalReference: 'page',
    ...endpoints,
    lineColor: color,
    lineWidth,
    lineStyle: vmlLineStyle(stroke ? attribute(stroke, 'dashstyle') : null),
    startArrow: vmlArrow(stroke ? attribute(stroke, 'startarrow') : null),
    endArrow: vmlArrow(stroke ? attribute(stroke, 'endarrow') : null),
  });
}

function drawingMlConnectorProperties(
  drawing: Element,
  index: number,
): WorkDocumentConnectorProperties {
  const shape = descendants(drawing, 'wsp')[0];
  const properties = shape ? directChild(shape, 'spPr') : undefined;
  const extent = directChild(drawing, 'extent');
  const transform = properties ? directChild(properties, 'xfrm') : undefined;
  const extension = directChild(
    transform ?? properties ?? shape ?? drawing,
    'ext',
  );
  const width = emuMillimeters(
    attribute(extent ?? extension ?? shape ?? drawing, 'cx'),
    120,
  );
  const height = emuMillimeters(
    attribute(extent ?? extension ?? shape ?? drawing, 'cy'),
    25,
  );
  const horizontal = directChild(drawing, 'positionH');
  const vertical = directChild(drawing, 'positionV');
  const line = properties ? directChild(properties, 'ln') : undefined;
  const color = drawingMlColor(line);
  const customGeometry = properties
    ? directChild(properties, 'custGeom')
    : null;
  const endpoints = drawingMlEndpoints(customGeometry);
  const docProperties = directChild(drawing, 'docPr');
  return normalizeDocumentConnectorProperties({
    id:
      (docProperties ? attribute(docProperties, 'name') : null)?.trim() ||
      `docx-connector-${index}`,
    width,
    height,
    layout: drawing.localName === 'anchor' ? 'floating' : 'inline',
    horizontalOffset: positionOffset(horizontal),
    verticalOffset: positionOffset(vertical),
    horizontalReference: positionReference(horizontal, 'column'),
    verticalReference: positionReference(vertical, 'paragraph'),
    ...endpoints,
    lineColor: color,
    lineWidth: emuMillimeters(line ? attribute(line, 'w') : null, 0.35),
    lineStyle: drawingMlLineStyle(
      line ? directChild(line, 'prstDash') : undefined,
    ),
    startArrow: drawingMlArrow(line ? directChild(line, 'headEnd') : undefined),
    endArrow: drawingMlArrow(line ? directChild(line, 'tailEnd') : undefined),
    docPropertiesId: integerAttribute(docProperties, 'id'),
  });
}

function drawingMlColor(line: Element | undefined): string {
  const fill = line ? directChild(line, 'solidFill') : undefined;
  const rgb = fill ? directChild(fill, 'srgbClr') : undefined;
  return normalizeConnectorColor(rgb ? attribute(rgb, 'val') : null);
}

function drawingMlEndpoints(
  geometry: Element | null | undefined,
): Pick<
  WorkDocumentConnectorProperties,
  'startX' | 'startY' | 'endX' | 'endY'
> {
  const path = geometry ? descendants(geometry, 'path')[0] : undefined;
  const move = path ? descendants(path, 'moveTo')[0] : undefined;
  const end = path ? descendants(path, 'lnTo')[0] : undefined;
  const first = move ? directChild(move, 'pt') : undefined;
  const second = end ? directChild(end, 'pt') : undefined;
  return {
    startX: percentage(first ? attribute(first, 'x') : null, 0),
    startY: percentage(first ? attribute(first, 'y') : null, 50),
    endX: percentage(second ? attribute(second, 'x') : null, 100),
    endY: percentage(second ? attribute(second, 'y') : null, 50),
  };
}

function parseVmlEndpoints(
  element: Element | undefined,
  coordinateSize: { width: number; height: number },
  width: number,
  height: number,
): Pick<
  WorkDocumentConnectorProperties,
  'startX' | 'startY' | 'endX' | 'endY'
> {
  const from = parsePoint(element ? attribute(element, 'from') : null);
  const to = parsePoint(element ? attribute(element, 'to') : null);
  if (from && to) {
    return {
      startX: percentage(from.x, coordinateSize.width),
      startY: percentage(from.y, coordinateSize.height),
      endX: percentage(to.x, coordinateSize.width),
      endY: percentage(to.y, coordinateSize.height),
    };
  }
  const horizontal = height <= Math.max(1, width * 0.05);
  return {
    startX: 0,
    startY: horizontal ? 50 : 0,
    endX: 100,
    endY: horizontal ? 50 : 100,
  };
}

function parseVmlStyle(value: string | null): Record<string, string> {
  const result: Record<string, string> = {};
  for (const item of value?.split(';') ?? []) {
    const separator = item.indexOf(':');
    if (separator < 0) continue;
    const key = item.slice(0, separator).trim().toLowerCase();
    const entry = item
      .slice(separator + 1)
      .trim()
      .toLowerCase();
    if (key) result[key] = entry;
  }
  return result;
}

function parseCoordinateSize(value: string | null): {
  width: number;
  height: number;
} {
  const [width, height] = value?.split(',').map(Number) ?? [];
  return {
    width: Number.isFinite(width) && width > 0 ? width : 21_600,
    height: Number.isFinite(height) && height > 0 ? height : 21_600,
  };
}

function parsePoint(value: string | null): { x: number; y: number } | null {
  if (!value) return null;
  const [x, y] = value.split(',').map(Number);
  return Number.isFinite(x) && Number.isFinite(y) ? { x, y } : null;
}

function percentage(
  value: string | number | null,
  scale: string | number,
): number {
  const number = typeof value === 'number' ? value : Number(value);
  const size = typeof scale === 'number' ? scale : Number(scale);
  if (!Number.isFinite(number) || !Number.isFinite(size) || size === 0)
    return 0;
  return Math.min(100, Math.max(0, (number / size) * 100));
}

function vmlArrow(value: string | null): WorkDocumentConnectorArrow {
  const normalized = value?.trim().toLowerCase();
  return normalized && normalized !== 'none' ? 'triangle' : 'none';
}

function drawingMlArrow(
  element: Element | undefined,
): WorkDocumentConnectorArrow {
  const value = element
    ? attribute(element, 'type')?.trim().toLowerCase()
    : null;
  return value && value !== 'none' ? 'triangle' : 'none';
}

function vmlLineStyle(value: string | null): WorkDocumentConnectorLineStyle {
  const normalized = value?.trim().toLowerCase();
  if (normalized === 'dash') return 'dash';
  if (normalized === 'dot') return 'dot';
  if (normalized === 'dashdot' || normalized === 'dash-dot') return 'dashDot';
  return 'solid';
}

function drawingMlLineStyle(
  element: Element | undefined,
): WorkDocumentConnectorLineStyle {
  const value = element ? attribute(element, 'val')?.trim() : undefined;
  if (value === 'dash') return 'dash';
  if (value === 'dot') return 'dot';
  if (value === 'dashDot') return 'dashDot';
  return 'solid';
}

function vmlLength(value: string | null | undefined, fallback: number): number {
  if (!value) return fallback;
  const match = value.trim().match(/^(-?\d+(?:\.\d+)?)\s*(pt|mm|cm|in|px)?$/i);
  if (!match) return fallback;
  const number = Number(match[1]);
  if (!Number.isFinite(number)) return fallback;
  const unit = match[2]?.toLowerCase() ?? 'pt';
  if (unit === 'mm') return Math.abs(number);
  if (unit === 'cm') return Math.abs(number * 10);
  if (unit === 'in') return Math.abs(number * 25.4);
  if (unit === 'px') return Math.abs((number * 25.4) / 96);
  return Math.abs(number / POINTS_PER_MILLIMETER);
}

function vmlOffset(
  left: string | undefined,
  margin: string | undefined,
): number | null {
  if (!left && !margin) return null;
  return (
    vmlLength(left, 0) * (left?.trim().startsWith('-') ? -1 : 1) +
    vmlLength(margin, 0) * (margin?.trim().startsWith('-') ? -1 : 1)
  );
}

function emuMillimeters(value: string | null, fallback: number): number {
  const number = value === null ? Number.NaN : Number(value);
  return Number.isFinite(number) && number >= 0
    ? number / EMUS_PER_MILLIMETER
    : fallback;
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
  const value = element ? attribute(element, 'relativeFrom') : null;
  if (value === 'margin' || value === 'page' || value === 'column')
    return value;
  if (value === 'paragraph') return value;
  return fallback;
}

function integerAttribute(
  element: Element | undefined,
  name: string,
): number | null {
  if (!element) return null;
  const value = Number(attribute(element, name));
  return Number.isSafeInteger(value) && value >= 0 ? value : null;
}

function isSoleDrawingParagraph(
  paragraph: Element | null,
  run: Element | null,
): boolean {
  if (!paragraph || !run) return false;
  const meaningful = directChildren(paragraph).filter(
    (element) =>
      element.localName !== 'pPr' &&
      !['bookmarkStart', 'bookmarkEnd', 'proofErr'].includes(element.localName),
  );
  if (meaningful.length !== 1 || meaningful[0] !== run) return false;
  const runContent = directChildren(run).filter(
    (element) => element.localName !== 'rPr',
  );
  return (
    runContent.length === 1 &&
    (runContent[0]?.localName === 'pict' ||
      runContent[0]?.localName === 'drawing' ||
      runContent[0]?.localName === 'AlternateContent')
  );
}

function replaceDrawingRunWithConnectorMarker(
  document: Document,
  drawingRun: Element,
  marker: string,
): void {
  const namespace = drawingRun.namespaceURI || WORD_NAMESPACE;
  const prefix = drawingRun.prefix ? `${drawingRun.prefix}:` : 'w:';
  const run = document.createElementNS(namespace, `${prefix}r`);
  const text = document.createElementNS(namespace, `${prefix}t`);
  text.setAttributeNS(
    'http://www.w3.org/XML/1998/namespace',
    'xml:space',
    'preserve',
  );
  text.textContent = marker;
  run.append(text);
  drawingRun.replaceWith(run);
}

function nextConnectorMarker(document: Document, start: number): string {
  let index = start;
  while (
    document.documentElement.textContent?.includes(
      `__A3S_WORK_CONNECTOR_${index}__`,
    )
  ) {
    index += 1;
  }
  return `__A3S_WORK_CONNECTOR_${index}__`;
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

function isWordNamespace(element: Element): boolean {
  return (
    element.namespaceURI === WORD_NAMESPACE ||
    element.namespaceURI === STRICT_WORD_NAMESPACE
  );
}

function normalizeConnectorColor(value: string | null): string {
  if (!value) return '#c00000';
  const normalized = value
    .trim()
    .replace(/^#/, '')
    .split(' ')[0]
    ?.toLowerCase();
  return normalized && /^[0-9a-f]{6}$/.test(normalized)
    ? `#${normalized}`
    : '#c00000';
}
