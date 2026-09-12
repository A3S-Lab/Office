import { MAX_DOCUMENT_NUMBERING_START } from './work-document-lists';
import {
  bulletStyleFromNumberingSuffix,
  isReviewableNumberingFormat,
  isBulletNumberingFormat,
  numberingTypeFromFormat,
  parseNumberingOriginalLevels,
  serializeDocumentNumberingChange,
  serializeNumberingOriginalLevels,
} from './work-document-numbering-changes';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { descendants, directChildren } from './work-ooxml-package';

export interface ImportedDocxNumberingChangeGroup {
  markers: string[];
  id: string;
  author: string;
  date: string;
  start: number;
  level: number;
  format: number;
  suffix: string;
  originalLevels: string;
}

export interface ImportedDocxNumberingChangeMarkers {
  groups: ImportedDocxNumberingChangeGroup[];
}

interface SupportedDocxNumberingChange {
  id: string;
  author: string;
  date: string;
  numberingId: number;
  level: number;
  value: number;
  format: number;
  suffix: string;
  originalLevels: string;
  paragraph: Element;
  properties: Element;
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const MAX_NUMBERING_CHANGES = 65_536;
const NUMBERING_CHANGE_MARKER_PATTERN = /__A3S_WORK_NUMBERING_CHANGE_\d+__/g;

export function markDocxNumberingChanges(
  document: Document,
): ImportedDocxNumberingChangeMarkers {
  const groups: ImportedDocxNumberingChangeGroup[] = [];
  let active:
    | {
        change: SupportedDocxNumberingChange;
        group: ImportedDocxNumberingChangeGroup;
      }
    | undefined;
  let markerIndex = 0;

  for (const paragraph of descendants(document, 'p')) {
    if (!isWordElement(paragraph)) {
      active = undefined;
      continue;
    }
    const change = supportedNumberingChangeInParagraph(paragraph);
    if (!change || markerIndex >= MAX_NUMBERING_CHANGES) {
      active = undefined;
      continue;
    }
    const continues = active && continuesGroup(active.change, change);
    if (!continues) {
      const group: ImportedDocxNumberingChangeGroup = {
        markers: [],
        id: `docx-numbering-change-${change.id}`,
        author: change.author,
        date: change.date,
        start: change.value,
        level: change.level,
        format: change.format,
        suffix: change.suffix,
        originalLevels: change.originalLevels,
      };
      groups.push(group);
      active = { change, group };
    }
    markerIndex += 1;
    const marker = `__A3S_WORK_NUMBERING_CHANGE_${markerIndex}__`;
    insertParagraphMarker(
      document,
      change.paragraph,
      change.properties,
      marker,
    );
    active?.group.markers.push(marker);
    if (active) active.change = change;
  }
  return { groups };
}

export function applyImportedDocxNumberingChangeMarkers(
  document: Document,
  markers: ImportedDocxNumberingChangeMarkers,
): void {
  const nodesByMarker = numberingMarkerNodes(document.body);
  const groupsByList = new Map<
    HTMLOListElement | HTMLUListElement,
    ImportedDocxNumberingChangeGroup[]
  >();
  for (const group of markers.groups) {
    const list = numberingChangeList(group, nodesByMarker);
    if (!list) continue;
    const groups = groupsByList.get(list) ?? [];
    groups.push(group);
    groupsByList.set(list, groups);
  }
  for (const [list, groups] of groupsByList) {
    if (groups.length === 1 && groups[0]) {
      applyNumberingChange(list, groups[0]);
    }
  }
  for (const nodes of nodesByMarker.values()) {
    for (const node of nodes) {
      node.data = node.data.replace(NUMBERING_CHANGE_MARKER_PATTERN, '');
    }
  }
  document.body.normalize();
}

export function hasImportedDocxNumberingChangeMarkers(
  markers: ImportedDocxNumberingChangeMarkers,
): boolean {
  return markers.groups.length > 0;
}

export function isSupportedDocxNumberingChange(change: Element): boolean {
  return supportedNumberingChange(change) !== null;
}

function supportedNumberingChangeInParagraph(
  paragraph: Element,
): SupportedDocxNumberingChange | null {
  const propertyNodes = directWordChildren(paragraph, 'pPr');
  if (propertyNodes.length !== 1) return null;
  const properties = propertyNodes[0];
  if (!properties) return null;
  const numberingNodes = directWordChildren(properties, 'numPr');
  if (numberingNodes.length !== 1) return null;
  const changes = directChildren(numberingNodes[0], 'numberingChange');
  if (changes.length !== 1) return null;
  return supportedNumberingChange(changes[0]);
}

function supportedNumberingChange(
  change: Element | undefined,
): SupportedDocxNumberingChange | null {
  if (
    !change ||
    change.localName !== 'numberingChange' ||
    !isWordElement(change)
  ) {
    return null;
  }
  const namespace = change.namespaceURI ?? '';
  const numbering = change.parentElement;
  const properties = numbering?.parentElement;
  const paragraph = properties?.parentElement;
  if (
    numbering?.localName !== 'numPr' ||
    numbering.namespaceURI !== namespace ||
    properties?.localName !== 'pPr' ||
    properties.namespaceURI !== namespace ||
    paragraph?.localName !== 'p' ||
    paragraph.namespaceURI !== namespace
  ) {
    return null;
  }
  const numberingChildren = directChildren(numbering);
  if (
    numberingChildren.some((child) => child.namespaceURI !== namespace) ||
    directChildren(numbering, 'numberingChange').length !== 1 ||
    change.children.length !== 0 ||
    hasUnsupportedWordAttribute(change)
  ) {
    return null;
  }
  const levels = directWordChildren(numbering, 'ilvl');
  const numberingIds = directWordChildren(numbering, 'numId');
  // OOXML CT_NumPr: w:ilvl is optional and defaults to 0. Require exactly one
  // w:numId; admit either {ilvl,numId,numberingChange} or {numId,numberingChange}.
  if (numberingIds.length !== 1 || levels.length > 1) return null;
  if (numberingChildren.length !== (levels.length === 1 ? 3 : 2)) return null;
  const level =
    levels.length === 1 ? boundedWordInteger(levels[0], 'val', 0, 8) : 0;
  const numberingId = boundedWordInteger(
    numberingIds[0],
    'val',
    0,
    MAX_DOCUMENT_NUMBERING_START,
  );
  const id = wordAttribute(change, 'id')?.trim() ?? '';
  const author = wordAttribute(change, 'author')?.trim() ?? '';
  const rawDate = wordAttribute(change, 'date');
  const original = wordAttribute(change, 'original') ?? '';
  if (
    level === null ||
    numberingId === null ||
    !/^\+?\d{1,10}$/.test(id) ||
    !author ||
    author.length > 255 ||
    /[\u0000-\u001f\u007f]/.test(author) ||
    (rawDate !== null && !Number.isFinite(Date.parse(rawDate)))
  ) {
    return null;
  }
  const definitions = parseNumberingOriginalLevels(original);
  const definition = definitions?.get(level + 1);
  // Current w:ilvl must stay reviewable (nfc 0–4 or bullet 23); sibling segments
  // may carry other ST_NumberFormat values and round-trip as opaque prior text.
  // Picture and other non-reviewable current-level changes stay fail-closed.
  if (
    !definition ||
    !definitions ||
    !isReviewableNumberingFormat(definition.format)
  ) {
    return null;
  }
  const originalLevels =
    definitions.size >= 2 ? serializeNumberingOriginalLevels(definitions) : '';
  return {
    id,
    author,
    date: normalizeRevisionDate(rawDate),
    numberingId,
    level,
    ...definition,
    originalLevels,
    paragraph,
    properties,
  };
}

function continuesGroup(
  previous: SupportedDocxNumberingChange,
  current: SupportedDocxNumberingChange,
): boolean {
  return (
    current.numberingId === previous.numberingId &&
    current.level === previous.level &&
    current.author === previous.author &&
    current.date === previous.date &&
    current.format === previous.format &&
    current.suffix === previous.suffix &&
    current.value === previous.value + 1 &&
    sameOriginalSiblingLevels(previous, current)
  );
}

function sameOriginalSiblingLevels(
  previous: SupportedDocxNumberingChange,
  current: SupportedDocxNumberingChange,
): boolean {
  if (!previous.originalLevels && !current.originalLevels) return true;
  const left = parseNumberingOriginalLevels(previous.originalLevels);
  const right = parseNumberingOriginalLevels(current.originalLevels);
  if (!left || !right || left.size !== right.size) return false;
  const currentLevel = previous.level + 1;
  for (const [level, definition] of left) {
    const other = right.get(level);
    if (!other) return false;
    if (level === currentLevel) continue;
    if (
      other.value !== definition.value ||
      other.format !== definition.format ||
      other.suffix !== definition.suffix
    ) {
      return false;
    }
  }
  return true;
}

function applyNumberingChange(
  list: HTMLOListElement | HTMLUListElement,
  change: ImportedDocxNumberingChangeGroup,
): void {
  if (isBulletNumberingFormat(change.format)) {
    if (!(list instanceof HTMLUListElement)) return;
    applyBulletNumberingChange(list, change);
    return;
  }
  if (!(list instanceof HTMLOListElement)) return;
  const type = numberingTypeFromFormat(change.format);
  if (type === undefined) return;
  const attributes = {
    start: change.start,
    type,
    officeNumberingId: optionalDataset(list, 'officeNumberingId'),
    officeAbstractNumberingId: optionalDataset(
      list,
      'officeAbstractNumberingId',
    ),
    officeNumberingLevel:
      optionalDataset(list, 'officeNumberingLevel') ?? String(change.level),
    officeNumberingFormat: numberingFormatName(change.format),
    officeNumberingText: `%${change.level + 1}${change.suffix}`,
    officeNumberingSuffix: optionalDataset(list, 'officeNumberingSuffix'),
    officeNumberingAlignment: optionalDataset(list, 'officeNumberingAlignment'),
    officeNumberingIndentLeft: optionalDataset(
      list,
      'officeNumberingIndentLeft',
    ),
    officeNumberingIndentRight: optionalDataset(
      list,
      'officeNumberingIndentRight',
    ),
    officeNumberingIndentStart: optionalDataset(
      list,
      'officeNumberingIndentStart',
    ),
    officeNumberingIndentEnd: optionalDataset(list, 'officeNumberingIndentEnd'),
    officeNumberingIndentHanging: optionalDataset(
      list,
      'officeNumberingIndentHanging',
    ),
    officeNumberingIndentFirstLine: optionalDataset(
      list,
      'officeNumberingIndentFirstLine',
    ),
    officeNumberingRestartAfterLevel: optionalDataset(
      list,
      'officeNumberingRestartAfterLevel',
    ),
    level: change.level,
    originalFormat: change.format,
    originalSuffix: change.suffix,
    originalLevels: change.originalLevels,
  };
  list.dataset.documentChange = 'true';
  list.dataset.changeKind = 'numbering';
  list.dataset.changeId = change.id;
  list.dataset.changeAuthor = change.author;
  list.dataset.changeDate = change.date;
  list.dataset.changeBefore = serializeDocumentNumberingChange(attributes);
}

function applyBulletNumberingChange(
  list: HTMLUListElement,
  change: ImportedDocxNumberingChangeGroup,
): void {
  const attributes = {
    start: change.start,
    type: null,
    bulletStyle: bulletStyleFromNumberingSuffix(change.suffix),
    officeNumberingId: optionalDataset(list, 'officeNumberingId'),
    officeAbstractNumberingId: optionalDataset(
      list,
      'officeAbstractNumberingId',
    ),
    officeNumberingLevel:
      optionalDataset(list, 'officeNumberingLevel') ?? String(change.level),
    officeNumberingFormat: 'bullet',
    officeNumberingText: `%${change.level + 1}${change.suffix}`,
    officeNumberingSuffix: optionalDataset(list, 'officeNumberingSuffix'),
    officeNumberingAlignment: optionalDataset(list, 'officeNumberingAlignment'),
    officeNumberingIndentLeft: optionalDataset(
      list,
      'officeNumberingIndentLeft',
    ),
    officeNumberingIndentRight: optionalDataset(
      list,
      'officeNumberingIndentRight',
    ),
    officeNumberingIndentStart: optionalDataset(
      list,
      'officeNumberingIndentStart',
    ),
    officeNumberingIndentEnd: optionalDataset(list, 'officeNumberingIndentEnd'),
    officeNumberingIndentHanging: optionalDataset(
      list,
      'officeNumberingIndentHanging',
    ),
    officeNumberingIndentFirstLine: optionalDataset(
      list,
      'officeNumberingIndentFirstLine',
    ),
    officeNumberingRestartAfterLevel: optionalDataset(
      list,
      'officeNumberingRestartAfterLevel',
    ),
    level: change.level,
    originalFormat: change.format,
    originalSuffix: change.suffix,
    originalLevels: change.originalLevels,
  };
  list.dataset.documentChange = 'true';
  list.dataset.changeKind = 'numbering';
  list.dataset.changeId = change.id;
  list.dataset.changeAuthor = change.author;
  list.dataset.changeDate = change.date;
  list.dataset.changeBefore = serializeDocumentNumberingChange(attributes);
}

function numberingChangeList(
  group: ImportedDocxNumberingChangeGroup,
  nodesByMarker: ReadonlyMap<string, Text[]>,
): HTMLOListElement | HTMLUListElement | null {
  if (!group.markers.length) return null;
  const occurrences = group.markers.map((marker) => nodesByMarker.get(marker));
  if (occurrences.some((nodes) => nodes?.length !== 1)) return null;
  const nodes = occurrences.map((nodes) => nodes?.[0]).filter(isTextNode);
  if (nodes.length !== group.markers.length) return null;
  const items = nodes.map((node) => node.parentElement?.closest('li'));
  const lists = items.map((item) => item?.parentElement);
  const list = lists[0];
  if (
    !list ||
    items.some((item) => !(item instanceof HTMLLIElement)) ||
    lists.some((candidate) => candidate !== list)
  ) {
    return null;
  }
  const expectsBullet = isBulletNumberingFormat(group.format);
  if (expectsBullet) {
    if (!(list instanceof HTMLUListElement)) return null;
    return contiguousListItems(list, items) ? list : null;
  }
  if (!(list instanceof HTMLOListElement)) return null;
  return contiguousListItems(list, items) ? list : null;
}

function contiguousListItems(
  list: HTMLOListElement | HTMLUListElement,
  items: Array<Element | null | undefined>,
): boolean {
  const directItems = Array.from(list.children).filter(
    (child): child is HTMLLIElement => child instanceof HTMLLIElement,
  );
  const indexes = items.map((item) =>
    directItems.indexOf(item as HTMLLIElement),
  );
  const first = indexes[0] ?? -1;
  return (
    first >= 0 && indexes.every((index, offset) => index === first + offset)
  );
}

function numberingMarkerNodes(root: ParentNode): Map<string, Text[]> {
  const result = new Map<string, Text[]>();
  for (const node of textNodes(root)) {
    for (const marker of node.data.match(NUMBERING_CHANGE_MARKER_PATTERN) ??
      []) {
      const nodes = result.get(marker) ?? [];
      nodes.push(node);
      result.set(marker, nodes);
    }
  }
  return result;
}

function insertParagraphMarker(
  document: Document,
  paragraph: Element,
  properties: Element,
  marker: string,
): void {
  const namespace = paragraph.namespaceURI ?? WORD_NAMESPACE;
  const prefix = paragraph.prefix || 'w';
  const run = document.createElementNS(namespace, `${prefix}:r`);
  const text = document.createElementNS(namespace, `${prefix}:t`);
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  run.append(text);
  paragraph.insertBefore(run, properties.nextSibling);
}

function directWordChildren(parent: ParentNode, localName: string): Element[] {
  const namespace = (parent as Element).namespaceURI;
  return directChildren(parent, localName).filter(
    (element) => element.namespaceURI === namespace && isWordElement(element),
  );
}

function boundedWordInteger(
  element: Element | undefined,
  name: string,
  minimum: number,
  maximum: number,
): number | null {
  if (!element) return null;
  const raw = wordAttribute(element, name);
  if (!raw || !/^\d{1,10}$/.test(raw)) return null;
  const value = Number(raw);
  return Number.isSafeInteger(value) && value >= minimum && value <= maximum
    ? value
    : null;
}

function wordAttribute(element: Element, localName: string): string | null {
  const namespace = element.namespaceURI;
  if (!namespace) return null;
  const attributes = Array.from(element.attributes).filter(
    (attribute) =>
      xmlAttributeLocalName(attribute) === localName &&
      xmlAttributeNamespace(element, attribute) === namespace,
  );
  return attributes.length === 1 ? (attributes[0]?.value ?? null) : null;
}

function hasUnsupportedWordAttribute(element: Element): boolean {
  const namespace = element.namespaceURI;
  const supported = new Set(['id', 'author', 'date', 'original']);
  return Array.from(element.attributes).some(
    (attribute) =>
      xmlAttributeNamespace(element, attribute) === namespace &&
      !supported.has(xmlAttributeLocalName(attribute)),
  );
}

function optionalDataset(
  element: HTMLElement,
  key: keyof DOMStringMap,
): string | null {
  const value = element.dataset[key];
  return value ? value : null;
}

function numberingFormatName(format: number): string {
  if (format === 1) return 'upperRoman';
  if (format === 2) return 'lowerRoman';
  if (format === 3) return 'upperLetter';
  if (format === 4) return 'lowerLetter';
  return 'decimal';
}

function normalizeRevisionDate(value: string | null): string {
  if (!value) return '';
  const time = Date.parse(value);
  return Number.isFinite(time) ? new Date(time).toISOString() : '';
}

function isWordElement(element: Element): boolean {
  return DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '');
}

function isTextNode(value: Text | undefined): value is Text {
  return value instanceof Text;
}

function textNodes(root: ParentNode): Text[] {
  const document = root.ownerDocument;
  const walker = document?.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes: Text[] = [];
  if (!walker) return nodes;
  while (walker.nextNode()) nodes.push(walker.currentNode as Text);
  return nodes;
}
