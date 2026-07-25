import JSZip from 'jszip';
import {
  attribute,
  childPath,
  descendants,
  directChild,
  directChildren,
  firstDescendant,
  parseXml,
} from './work-ooxml-package';
import { presentationGroupPath } from './work-presentation-groups';
import type { WorkSlideElement } from './work-types';

const EXPORT_MARKER_PREFIX = '__A3S_OFFICE_EXPORT__';
const PRESENTATIONML_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWINGML_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const GROUPABLE_PART =
  /^ppt\/(?:slides|slideLayouts|slideMasters)\/[^/]+\.xml$/;

export type PptxExportObjectRole =
  | 'chart'
  | 'image'
  | 'line'
  | 'shape'
  | 'table'
  | 'text';

interface PptxExportBinding {
  displayName: string;
  groupPath: string[];
  groupScope: string;
  marker: string;
}

interface PptxSceneNode {
  binding?: PptxExportBinding;
  node: Element;
  order: number;
}

interface PptxGroupBucket {
  children: Map<string, PptxGroupBucket>;
  directNodes: PptxSceneNode[];
  order: number;
}

interface PptxBounds {
  bottom: number;
  left: number;
  right: number;
  top: number;
}

export class PptxGroupExportRegistry {
  private readonly bindings = new Map<string, PptxExportBinding>();
  private readonly roleCounts = new Map<PptxExportObjectRole, number>();
  private markerSequence = 0;

  get size(): number {
    return this.bindings.size;
  }

  objectName(
    scope: string,
    element: WorkSlideElement,
    role: PptxExportObjectRole,
  ): string | undefined {
    const groupPath = presentationGroupPath(element);
    if (!groupPath.length) return undefined;
    const roleCount = (this.roleCounts.get(role) ?? 0) + 1;
    this.roleCounts.set(role, roleCount);
    this.markerSequence += 1;
    const marker = `${EXPORT_MARKER_PREFIX}${this.markerSequence}`;
    this.bindings.set(marker, {
      displayName: `${pptxRoleName(role)} ${roleCount}`,
      groupPath,
      groupScope: scope,
      marker,
    });
    return marker;
  }

  binding(marker: string): PptxExportBinding | undefined {
    return this.bindings.get(marker);
  }

  values(): IterableIterator<PptxExportBinding> {
    return this.bindings.values();
  }
}

export async function patchPptxNativeGroups(
  buffer: ArrayBuffer,
  registry: PptxGroupExportRegistry,
): Promise<ArrayBuffer> {
  if (!registry.size) return buffer;
  const archive = await JSZip.loadAsync(buffer);
  const found = new Map<string, number>();
  const partPaths = Object.keys(archive.files)
    .filter((path) => GROUPABLE_PART.test(path))
    .sort();

  for (const partPath of partPaths) {
    const entry = archive.file(partPath);
    if (!entry) continue;
    const source = await entry.async('text');
    if (!source.includes(EXPORT_MARKER_PREFIX)) continue;
    const document = parseXml(source, partPath);
    let changed = false;
    for (const shapeTree of descendants(document, 'spTree')) {
      changed = patchPptxShapeTree(shapeTree, registry, found) || changed;
    }
    if (changed) {
      archive.file(partPath, new XMLSerializer().serializeToString(document));
    }
  }

  for (const binding of registry.values()) {
    const count = found.get(binding.marker) ?? 0;
    if (count !== 1) {
      throw new Error(
        `PPTX group export expected one generated object for ${binding.displayName}, but found ${count}.`,
      );
    }
  }
  return archive.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
  });
}

function patchPptxShapeTree(
  shapeTree: Element,
  registry: PptxGroupExportRegistry,
  found: Map<string, number>,
): boolean {
  const sceneNodes = directChildren(shapeTree)
    .filter(isPptxSceneNode)
    .map<PptxSceneNode>((node, order) => {
      const properties = firstDescendant(node, 'cNvPr');
      const objectName = properties ? attribute(properties, 'name') : null;
      if (!objectName?.startsWith(EXPORT_MARKER_PREFIX)) {
        return { node, order };
      }
      const binding = registry.binding(objectName);
      if (!binding) {
        throw new Error(
          `PPTX group export found an unknown object marker in the generated package: ${objectName}.`,
        );
      }
      found.set(binding.marker, (found.get(binding.marker) ?? 0) + 1);
      properties?.setAttribute('name', binding.displayName);
      return { binding, node, order };
    });
  if (!sceneNodes.some((item) => item.binding)) return false;

  const roots = new Map<string, PptxGroupBucket>();
  const rootNodes = sceneNodes.filter((item) => !item.binding);
  for (const item of sceneNodes) {
    if (!item.binding) continue;
    let siblings = roots;
    let bucket: PptxGroupBucket | undefined;
    for (const groupId of item.binding.groupPath) {
      const key = groupBucketKey(item.binding.groupScope, groupId);
      bucket = siblings.get(key);
      if (!bucket) {
        bucket = {
          children: new Map(),
          directNodes: [],
          order: item.order,
        };
        siblings.set(key, bucket);
      }
      bucket.order = Math.max(bucket.order, item.order);
      siblings = bucket.children;
    }
    bucket?.directNodes.push(item);
  }

  const nextId = pptxNonVisualIdFactory(shapeTree);
  for (const item of sceneNodes) item.node.remove();
  let groupNumber = 0;
  const renderedRoots = [
    ...rootNodes,
    ...Array.from(roots.values()).map((bucket) => ({
      node: renderPptxGroup(bucket, shapeTree.ownerDocument, nextId, () => {
        groupNumber += 1;
        return groupNumber;
      }),
      order: bucket.order,
    })),
  ].sort((left, right) => left.order - right.order);
  const extensionList = directChild(shapeTree, 'extLst') ?? null;
  for (const item of renderedRoots) {
    shapeTree.insertBefore(item.node, extensionList);
  }
  normalizePptxNonVisualIds(shapeTree);
  return true;
}

function renderPptxGroup(
  bucket: PptxGroupBucket,
  document: Document,
  nextId: () => number,
  nextGroupNumber: () => number,
): Element {
  const nonVisualId = nextId();
  const groupNumber = nextGroupNumber();
  const children = [
    ...bucket.directNodes,
    ...Array.from(bucket.children.values()).map((child) => ({
      node: renderPptxGroup(child, document, nextId, nextGroupNumber),
      order: child.order,
    })),
  ]
    .sort((left, right) => left.order - right.order)
    .map((item) => item.node);
  const bounds = unionPptxBounds(children.map(pptxNodeBounds));
  const group = document.createElementNS(PRESENTATIONML_NAMESPACE, 'p:grpSp');
  const nonVisual = document.createElementNS(
    PRESENTATIONML_NAMESPACE,
    'p:nvGrpSpPr',
  );
  const properties = document.createElementNS(
    PRESENTATIONML_NAMESPACE,
    'p:cNvPr',
  );
  properties.setAttribute('id', String(nonVisualId));
  properties.setAttribute('name', `Group ${groupNumber}`);
  nonVisual.append(
    properties,
    document.createElementNS(PRESENTATIONML_NAMESPACE, 'p:cNvGrpSpPr'),
    document.createElementNS(PRESENTATIONML_NAMESPACE, 'p:nvPr'),
  );
  const groupProperties = document.createElementNS(
    PRESENTATIONML_NAMESPACE,
    'p:grpSpPr',
  );
  groupProperties.append(pptxIdentityGroupTransform(document, bounds));
  group.append(nonVisual, groupProperties, ...children);
  return group;
}

function pptxIdentityGroupTransform(
  document: Document,
  bounds: PptxBounds,
): Element {
  const transform = document.createElementNS(DRAWINGML_NAMESPACE, 'a:xfrm');
  const x = Math.round(bounds.left);
  const y = Math.round(bounds.top);
  const width = Math.max(1, Math.round(bounds.right - bounds.left));
  const height = Math.max(1, Math.round(bounds.bottom - bounds.top));
  transform.append(
    pptxPositionNode(document, 'a:off', x, y),
    pptxExtentNode(document, 'a:ext', width, height),
    pptxPositionNode(document, 'a:chOff', x, y),
    pptxExtentNode(document, 'a:chExt', width, height),
  );
  return transform;
}

function pptxPositionNode(
  document: Document,
  name: 'a:chOff' | 'a:off',
  x: number,
  y: number,
): Element {
  const node = document.createElementNS(DRAWINGML_NAMESPACE, name);
  node.setAttribute('x', String(x));
  node.setAttribute('y', String(y));
  return node;
}

function pptxExtentNode(
  document: Document,
  name: 'a:chExt' | 'a:ext',
  width: number,
  height: number,
): Element {
  const node = document.createElementNS(DRAWINGML_NAMESPACE, name);
  node.setAttribute('cx', String(width));
  node.setAttribute('cy', String(height));
  return node;
}

function pptxNodeBounds(node: Element): PptxBounds {
  const transform =
    (node.localName === 'grpSp'
      ? childPath(node, 'grpSpPr', 'xfrm')
      : node.localName === 'graphicFrame'
        ? directChild(node, 'xfrm')
        : childPath(node, 'spPr', 'xfrm')) ?? firstDescendant(node, 'xfrm');
  const offset = directChild(transform ?? node, 'off');
  const extent = directChild(transform ?? node, 'ext');
  const x = finitePptxCoordinate(offset, 'x');
  const y = finitePptxCoordinate(offset, 'y');
  const width = finitePptxCoordinate(extent, 'cx');
  const height = finitePptxCoordinate(extent, 'cy');
  if (
    x === undefined ||
    y === undefined ||
    width === undefined ||
    height === undefined
  ) {
    const name = attribute(firstDescendant(node, 'cNvPr') ?? node, 'name');
    throw new Error(
      `PPTX group export could not resolve generated geometry for ${name || node.localName}.`,
    );
  }
  return {
    left: x,
    top: y,
    right: x + width,
    bottom: y + height,
  };
}

function unionPptxBounds(bounds: PptxBounds[]): PptxBounds {
  if (!bounds.length) {
    throw new Error('PPTX group export cannot create an empty native group.');
  }
  return {
    left: Math.min(...bounds.map((item) => item.left)),
    top: Math.min(...bounds.map((item) => item.top)),
    right: Math.max(...bounds.map((item) => item.right)),
    bottom: Math.max(...bounds.map((item) => item.bottom)),
  };
}

function finitePptxCoordinate(
  node: Element | undefined,
  name: string,
): number | undefined {
  if (!node) return undefined;
  const raw = attribute(node, name);
  if (raw === null || raw.trim() === '') return undefined;
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

function pptxNonVisualIdFactory(shapeTree: Element): () => number {
  const used = new Set(
    descendants(shapeTree, 'cNvPr')
      .map((node) => Number(attribute(node, 'id')))
      .filter((value) => Number.isInteger(value) && value > 0),
  );
  let candidate = 1;
  return () => {
    while (used.has(candidate)) candidate += 1;
    const value = candidate;
    used.add(value);
    candidate += 1;
    return value;
  };
}

function normalizePptxNonVisualIds(shapeTree: Element): void {
  const used = new Set<number>();
  let candidate = 1;
  const nextId = () => {
    while (used.has(candidate)) candidate += 1;
    const value = candidate;
    used.add(value);
    candidate += 1;
    return value;
  };
  for (const properties of descendants(shapeTree, 'cNvPr')) {
    const current = Number(attribute(properties, 'id'));
    if (Number.isInteger(current) && current > 0 && !used.has(current)) {
      used.add(current);
      continue;
    }
    properties.setAttribute('id', String(nextId()));
  }
}

function groupBucketKey(scope: string, groupId: string): string {
  return `${scope}\u0000${groupId}`;
}

function isPptxSceneNode(node: Element): boolean {
  return [
    'contentPart',
    'cxnSp',
    'graphicFrame',
    'grpSp',
    'pic',
    'sp',
  ].includes(node.localName);
}

function pptxRoleName(role: PptxExportObjectRole): string {
  if (role === 'image') return 'Picture';
  return `${role[0]?.toUpperCase() ?? ''}${role.slice(1)}`;
}
