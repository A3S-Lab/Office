import JSZip from 'jszip';
import { decodeXmlBytes } from './work-ooxml-xml';

const RELATIONSHIPS_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export interface OoxmlRelationship {
  id: string;
  target: string;
  type: string;
  targetMode?: string;
}

export class OoxmlPackage {
  private readonly textCache = new Map<string, Promise<string>>();

  private constructor(private readonly zip: JSZip) {}

  static async load(buffer: ArrayBuffer): Promise<OoxmlPackage> {
    return new OoxmlPackage(await JSZip.loadAsync(buffer));
  }

  has(partPath: string): boolean {
    return Boolean(this.zip.file(partPath));
  }

  paths(prefix: string): string[] {
    return Object.keys(this.zip.files).filter(
      (path) => path.startsWith(prefix) && !this.zip.files[path]?.dir,
    );
  }

  async text(partPath: string): Promise<string> {
    const cached = this.textCache.get(partPath);
    if (cached) return cached;
    const entry = this.zip.file(partPath);
    if (!entry) throw new Error(`Office package part is missing: ${partPath}`);
    const pending = entry
      .async('uint8array')
      .then((bytes) => decodeXmlBytes(bytes, partPath));
    this.textCache.set(partPath, pending);
    try {
      return await pending;
    } catch (error) {
      this.textCache.delete(partPath);
      throw error;
    }
  }

  async xml(partPath: string): Promise<Document> {
    return parseXml(await this.text(partPath), partPath);
  }

  async bytes(partPath: string): Promise<Uint8Array> {
    const entry = this.zip.file(partPath);
    if (!entry) throw new Error(`Office package part is missing: ${partPath}`);
    return entry.async('uint8array');
  }

  async relationships(
    sourcePart: string,
  ): Promise<Map<string, OoxmlRelationship>> {
    const partPath = relationshipsPartPath(sourcePart);
    if (!this.has(partPath)) return new Map();
    const document = await this.xml(partPath);
    return new Map(
      descendants(document, 'Relationship').map((element) => {
        const relationship: OoxmlRelationship = {
          id: attribute(element, 'Id') ?? '',
          target: resolvePartTarget(
            sourcePart,
            attribute(element, 'Target') ?? '',
          ),
          type: attribute(element, 'Type') ?? '',
          targetMode: attribute(element, 'TargetMode') ?? undefined,
        };
        return [relationship.id, relationship];
      }),
    );
  }
}

/**
 * Parse Office Open XML with the environment DOMParser.
 *
 * Spec-compliant browsers keep both `id` and `r:id` on the same element.
 * happy-dom drops the prefixed attribute when an unprefixed attribute of the
 * same local name appears earlier on the start tag (real PPTX: `id="256"
 * r:id="rId2"`). Neutralize that collision before parse so slide relationship
 * ids survive in tests and incomplete DOM hosts without switching to a second
 * XML stack that lacks Element.remove / children.
 */
export function parseXml(source: string, label = 'Office XML'): Document {
  const document = new DOMParser().parseFromString(
    neutralizePrefixedAttributeCollisions(source),
    'application/xml',
  );
  const error = descendants(document, 'parsererror')[0];
  if (error)
    throw new Error(
      `${label} is not valid XML: ${error.textContent?.trim() || 'parse error'}`,
    );
  return document;
}

/**
 * Move `prefix:local` attributes ahead of a bare `local` sibling on the same
 * start tag so incomplete DOM parsers retain both. Infoset-equivalent for
 * consumers that read by name / namespace.
 */
export function neutralizePrefixedAttributeCollisions(source: string): string {
  return source.replace(
    /<([A-Za-z_][\w:.-]*)(\s[^>]*?)?(\/?)>/g,
    (full, name: string, attrText: string | undefined, close: string) => {
      if (!attrText) return full;
      if (!/\sr:id="/.test(attrText) || !/(?:^|\s)id="/.test(attrText)) {
        return full;
      }
      const relationshipIds: string[] = [];
      const withoutRelationshipIds = attrText.replace(
        /\s+r:id="[^"]*"/g,
        (match) => {
          relationshipIds.push(match);
          return '';
        },
      );
      if (!relationshipIds.length) return full;
      return `<${name}${relationshipIds.join('')}${withoutRelationshipIds}${close}>`;
    },
  );
}

const xmlElementPatterns = new Map<string, RegExp>();

/**
 * Checks raw XML for selected element names before paying the cost of a full
 * DOM tree. False positives are safe because they only select the slower
 * parser path; the tag boundary prevents cell values and longer tag names
 * from selecting it accidentally.
 */
export function xmlContainsAnyElement(
  source: string,
  localNames: readonly string[],
): boolean {
  if (!source || !localNames.length) return false;
  const key = localNames.join('\u0000');
  let pattern = xmlElementPatterns.get(key);
  if (!pattern) {
    const alternatives = localNames
      .map((name) => name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
      .join('|');
    pattern = new RegExp(
      `<(?:[A-Za-z_][\\w.-]*:)?(?:${alternatives})(?=[\\s/>])`,
    );
    xmlElementPatterns.set(key, pattern);
  }
  return pattern.test(source);
}

export function attribute(element: Element, name: string): string | null {
  if (name === 'r:id') {
    const namespaced =
      typeof element.getAttributeNS === 'function'
        ? element.getAttributeNS(RELATIONSHIPS_NS, 'id')
        : null;
    if (namespaced) return namespaced;
  }
  const direct = element.getAttribute(name);
  if (direct !== null) return direct;
  const localName = name.includes(':')
    ? name.slice(name.indexOf(':') + 1)
    : name;
  return (
    listAttributes(element).find((item) => {
      const itemLocalName = item.localName.includes(':')
        ? item.localName.slice(item.localName.indexOf(':') + 1)
        : item.localName;
      return (
        itemLocalName === localName &&
        (!name.includes(':') || item.name === name)
      );
    })?.value ?? null
  );
}

export function xmlNamespacePrefix(
  element: Element,
  namespace: string | null,
): string | null {
  if (!namespace) return element.prefix;
  if (typeof element.lookupPrefix === 'function') {
    const prefix = element.lookupPrefix(namespace);
    if (prefix) return prefix;
  }
  let current: Element | null = element;
  while (current) {
    if (current.namespaceURI === namespace && current.prefix) {
      return current.prefix;
    }
    const declaration = listAttributes(current).find(
      (item) =>
        item.value === namespace &&
        (item.name === 'xmlns' || item.name.startsWith('xmlns:')),
    );
    if (declaration?.name.startsWith('xmlns:')) {
      return declaration.name.slice('xmlns:'.length);
    }
    current = parentElement(current);
  }
  return null;
}

export function directChildren(
  parent: ParentNode,
  localName?: string,
): Element[] {
  return elementChildren(parent).filter(
    (element) => !localName || element.localName === localName,
  );
}

export function directChild(
  parent: ParentNode,
  localName: string,
): Element | undefined {
  return directChildren(parent, localName)[0];
}

export function descendants(parent: ParentNode, localName: string): Element[] {
  return allElements(parent).filter(
    (element) => element.localName === localName,
  );
}

function listAttributes(element: Element): Attr[] {
  const attrs = element.attributes;
  if (!attrs) return [];
  if (typeof (attrs as Iterable<Attr>)[Symbol.iterator] === 'function') {
    return Array.from(attrs as Iterable<Attr>);
  }
  const listed: Attr[] = [];
  for (let index = 0; index < attrs.length; index += 1) {
    const item = attrs.item(index);
    if (item) listed.push(item);
  }
  return listed;
}

function elementChildren(parent: ParentNode): Element[] {
  const children = (parent as ParentNode & { children?: HTMLCollection })
    .children;
  if (children) return Array.from(children);
  return Array.from(parent.childNodes).filter(
    (node): node is Element => node.nodeType === 1,
  );
}

function allElements(parent: ParentNode): Element[] {
  if (typeof parent.querySelectorAll === 'function') {
    return Array.from(parent.querySelectorAll('*'));
  }
  const tagged = parent as ParentNode & {
    getElementsByTagName?: (name: string) => HTMLCollectionOf<Element>;
  };
  if (typeof tagged.getElementsByTagName === 'function') {
    return Array.from(tagged.getElementsByTagName('*'));
  }
  const collected: Element[] = [];
  const visit = (node: Node) => {
    if (node.nodeType === 1) {
      collected.push(node as Element);
      for (const child of Array.from(node.childNodes)) visit(child);
    }
  };
  for (const child of Array.from(parent.childNodes)) visit(child);
  return collected;
}

function parentElement(element: Element): Element | null {
  if (element.parentElement) return element.parentElement;
  const parent = element.parentNode;
  return parent && parent.nodeType === 1 ? (parent as Element) : null;
}

export function firstDescendant(
  parent: ParentNode | null | undefined,
  localName: string,
): Element | undefined {
  if (!parent) return undefined;
  return descendants(parent, localName)[0];
}

export function childPath(
  parent: ParentNode | null | undefined,
  ...localNames: string[]
): Element | undefined {
  let current = parent;
  for (const name of localNames) {
    if (!current) return undefined;
    current = directChild(current, name);
  }
  return current instanceof Element ? current : undefined;
}

export function resolvePartTarget(sourcePart: string, target: string): string {
  if (/^[a-z][a-z0-9+.-]*:/i.test(target)) return target;
  const segments = target.startsWith('/')
    ? []
    : sourcePart.split('/').slice(0, -1);
  for (const segment of target.replace(/^\/+/, '').split('/')) {
    if (!segment || segment === '.') continue;
    if (segment === '..') segments.pop();
    else segments.push(segment);
  }
  return segments.join('/');
}

export function contentTypeForPart(partPath: string): string {
  const extension = partPath.split('.').pop()?.toLowerCase();
  const types: Record<string, string> = {
    apng: 'image/apng',
    bmp: 'image/bmp',
    emf: 'image/emf',
    gif: 'image/gif',
    jpeg: 'image/jpeg',
    jpg: 'image/jpeg',
    png: 'image/png',
    svg: 'image/svg+xml',
    tif: 'image/tiff',
    tiff: 'image/tiff',
    webp: 'image/webp',
    wmf: 'image/wmf',
  };
  return types[extension ?? ''] ?? 'application/octet-stream';
}

export function bytesToDataUrl(bytes: Uint8Array, contentType: string): string {
  let binary = '';
  const chunkSize = 32_768;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(
      ...bytes.subarray(offset, offset + chunkSize),
    );
  }
  return `data:${contentType};base64,${btoa(binary)}`;
}

function relationshipsPartPath(sourcePart: string): string {
  const separator = sourcePart.lastIndexOf('/');
  const directory = separator >= 0 ? sourcePart.slice(0, separator + 1) : '';
  const fileName =
    separator >= 0 ? sourcePart.slice(separator + 1) : sourcePart;
  return `${directory}_rels/${fileName}.rels`;
}
