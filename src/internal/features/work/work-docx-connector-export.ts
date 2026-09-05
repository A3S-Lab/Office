import JSZip from 'jszip';
import type { WorkDocumentConnectorProperties } from './work-document-connector';
import { descendants, directChild, parseXml } from './work-ooxml-package';

interface DocxConnectorPatch {
  marker: string;
  preferredId: number | null;
  properties: WorkDocumentConnectorProperties;
}

export interface DocxConnectorRegistration {
  marker: string;
  docPropertiesId: number | null;
}

const DOC_PROPERTIES_ID_MAX = 0xffff_ffff;
const CONNECTOR_ID_MARKER_PATTERN = /__A3S_CONNECTOR_ID_\d+__/g;
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const GEOMETRY_SCALE = 100_000;
const EMUS_PER_MILLIMETER = 36_000;

/** Allocates stable drawing-property markers before package-wide ID repair. */
export class DocxConnectorPatchCollector {
  readonly patches: DocxConnectorPatch[] = [];
  private nextMarker = 1;

  register(
    properties: WorkDocumentConnectorProperties,
  ): DocxConnectorRegistration {
    const marker = `__A3S_CONNECTOR_ID_${this.nextMarker}__`;
    this.nextMarker += 1;
    this.patches.push({
      marker,
      preferredId: properties.docPropertiesId,
      properties,
    });
    return { marker, docPropertiesId: properties.docPropertiesId };
  }
}

export async function patchDocxConnectors(
  buffer: ArrayBuffer,
  patches: readonly DocxConnectorPatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  const archive = await JSZip.loadAsync(buffer);
  const byMarker = new Map(patches.map((patch) => [patch.marker, patch]));
  const entries = Object.values(archive.files).filter(
    (entry) =>
      !entry.dir &&
      /^word\/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/.test(
        entry.name,
      ),
  );
  const documents = await Promise.all(
    entries.map(async (entry) => ({
      entry,
      document: parseXml(await entry.async('text'), entry.name),
    })),
  );
  const properties = documents.flatMap(({ document }) =>
    descendants(document, 'docPr'),
  );
  const marked = properties
    .map((property) => ({
      property,
      patch: connectorPatch(property, byMarker),
    }))
    .filter(
      (
        value,
      ): value is {
        property: Element;
        patch: DocxConnectorPatch;
      } => Boolean(value.patch),
    );
  if (!marked.length) return buffer;

  const used = new Set<number>();
  for (const property of properties) {
    if (marked.some((value) => value.property === property)) continue;
    const id = parseDocPropertiesId(property.getAttribute('id'));
    if (id !== null) used.add(id);
  }
  let nextId = 1;
  const changedDocuments = new Set<Document>();
  for (const { property, patch } of marked) {
    removeConnectorMarker(property, patch.marker);
    const preferred =
      patch.preferredId !== null && !used.has(patch.preferredId)
        ? patch.preferredId
        : nextAvailableDocPropertiesId(used, nextId);
    used.add(preferred);
    nextId = preferred + 1;
    property.setAttribute('id', String(preferred));
    patchConnectorShape(property, patch.properties);
    changedDocuments.add(property.ownerDocument);
  }

  for (const { entry, document } of documents) {
    if (changedDocuments.has(document)) {
      archive.file(entry.name, new XMLSerializer().serializeToString(document));
    }
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function connectorPatch(
  property: Element,
  byMarker: ReadonlyMap<string, DocxConnectorPatch>,
): DocxConnectorPatch | null {
  for (const attribute of Array.from(property.attributes)) {
    const marker = attribute.value
      .match(CONNECTOR_ID_MARKER_PATTERN)
      ?.find((value) => byMarker.has(value));
    if (marker) return byMarker.get(marker) ?? null;
  }
  return null;
}

function removeConnectorMarker(property: Element, marker: string): void {
  for (const attribute of Array.from(property.attributes)) {
    if (!attribute.value.includes(marker)) continue;
    attribute.value = attribute.value
      .replace(marker, '')
      .replace(/\s{2,}/g, ' ')
      .trim();
  }
}

function patchConnectorShape(
  property: Element,
  properties: WorkDocumentConnectorProperties,
): void {
  let ancestor = property.parentElement;
  while (ancestor) {
    const shape = descendants(ancestor, 'wsp').find(
      (candidate) => candidate.localName === 'wsp',
    );
    if (shape) {
      patchShapeContents(shape, properties);
      return;
    }
    ancestor = ancestor.parentElement;
  }
}

function patchShapeContents(
  shape: Element,
  properties: WorkDocumentConnectorProperties,
): void {
  const document = shape.ownerDocument;
  const nonVisual = directChild(shape, 'cNvSpPr');
  nonVisual?.removeAttribute('txBox');
  directChild(shape, 'txbx')?.remove();
  directChild(shape, 'bodyPr')?.remove();
  const shapeProperties = directChild(shape, 'spPr');
  if (!shapeProperties) return;
  directChild(shapeProperties, 'prstGeom')?.remove();
  directChild(shapeProperties, 'custGeom')?.remove();
  const geometry = connectorGeometry(document, properties);
  const firstFill =
    directChild(shapeProperties, 'solidFill') ??
    directChild(shapeProperties, 'noFill');
  if (firstFill) shapeProperties.insertBefore(geometry, firstFill);
  else shapeProperties.append(geometry);
  directChild(shapeProperties, 'ln')?.remove();
  shapeProperties.append(connectorLine(document, properties));
}

function connectorGeometry(
  document: Document,
  properties: WorkDocumentConnectorProperties,
): Element {
  const geometry = document.createElementNS(DRAWING_NAMESPACE, 'a:custGeom');
  appendEmpty(document, geometry, 'a:avLst');
  appendEmpty(document, geometry, 'a:gdLst');
  appendEmpty(document, geometry, 'a:ahLst');
  appendEmpty(document, geometry, 'a:cxnLst');
  const rect = document.createElementNS(DRAWING_NAMESPACE, 'a:rect');
  rect.setAttribute('l', '0');
  rect.setAttribute('t', '0');
  rect.setAttribute('r', String(GEOMETRY_SCALE));
  rect.setAttribute('b', String(GEOMETRY_SCALE));
  geometry.append(rect);
  const pathList = document.createElementNS(DRAWING_NAMESPACE, 'a:pathLst');
  const path = document.createElementNS(DRAWING_NAMESPACE, 'a:path');
  path.setAttribute('w', String(GEOMETRY_SCALE));
  path.setAttribute('h', String(GEOMETRY_SCALE));
  path.append(
    connectorPoint(document, 'a:moveTo', properties.startX, properties.startY),
    connectorPoint(document, 'a:lnTo', properties.endX, properties.endY),
  );
  pathList.append(path);
  geometry.append(pathList);
  return geometry;
}

function connectorPoint(
  document: Document,
  elementName: string,
  x: number,
  y: number,
): Element {
  const wrapper = document.createElementNS(DRAWING_NAMESPACE, elementName);
  const point = document.createElementNS(DRAWING_NAMESPACE, 'a:pt');
  point.setAttribute('x', String(Math.round((x / 100) * GEOMETRY_SCALE)));
  point.setAttribute('y', String(Math.round((y / 100) * GEOMETRY_SCALE)));
  wrapper.append(point);
  return wrapper;
}

function connectorLine(
  document: Document,
  properties: WorkDocumentConnectorProperties,
): Element {
  const line = document.createElementNS(DRAWING_NAMESPACE, 'a:ln');
  line.setAttribute(
    'w',
    String(Math.round(properties.lineWidth * EMUS_PER_MILLIMETER)),
  );
  const fill = document.createElementNS(DRAWING_NAMESPACE, 'a:solidFill');
  const color = document.createElementNS(DRAWING_NAMESPACE, 'a:srgbClr');
  color.setAttribute('val', properties.lineColor.slice(1).toUpperCase());
  fill.append(color);
  line.append(fill);
  if (properties.lineStyle !== 'solid')
    line.append(lineDash(document, properties.lineStyle));
  if (properties.startArrow !== 'none')
    line.append(arrowEnd(document, 'a:headEnd'));
  if (properties.endArrow !== 'none')
    line.append(arrowEnd(document, 'a:tailEnd'));
  return line;
}

function lineDash(
  document: Document,
  style: WorkDocumentConnectorProperties['lineStyle'],
): Element {
  const dash = document.createElementNS(DRAWING_NAMESPACE, 'a:prstDash');
  dash.setAttribute('val', style);
  return dash;
}

function arrowEnd(document: Document, name: string): Element {
  const arrow = document.createElementNS(DRAWING_NAMESPACE, name);
  arrow.setAttribute('type', 'triangle');
  arrow.setAttribute('w', 'med');
  arrow.setAttribute('len', 'med');
  return arrow;
}

function appendEmpty(document: Document, parent: Element, name: string): void {
  parent.append(document.createElementNS(DRAWING_NAMESPACE, name));
}

function parseDocPropertiesId(value: string | null): number | null {
  if (value === null || !/^\d+$/.test(value)) return null;
  const id = Number(value);
  return Number.isSafeInteger(id) && id <= DOC_PROPERTIES_ID_MAX ? id : null;
}

function nextAvailableDocPropertiesId(
  used: ReadonlySet<number>,
  start: number,
): number {
  let candidate = Math.max(1, start);
  while (candidate <= DOC_PROPERTIES_ID_MAX && used.has(candidate)) {
    candidate += 1;
  }
  if (candidate > DOC_PROPERTIES_ID_MAX) {
    throw new Error('DOCX drawing-property IDs are exhausted.');
  }
  return candidate;
}
