import { attribute, descendants, directChild } from './work-ooxml-package';

export interface ImportedDocxListMarker {
  marker: string;
  start: number;
  type?: ImportedDocxOrderedListType;
}

export interface ImportedDocxListMarkers {
  lists: ImportedDocxListMarker[];
}

export type ImportedDocxOrderedListType = 'a' | 'A' | 'i' | 'I';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const LIST_START_MARKER_PATTERN = /__A3S_WORK_LIST_START_\d+__/g;

export function markDocxLists(
  document: Document,
  numbering: Document | null,
): ImportedDocxListMarkers {
  if (!numbering) return { lists: [] };
  const abstractNumbering = new Map(
    descendants(numbering, 'abstractNum').flatMap((element) => {
      const id = integerAttribute(element, 'abstractNumId');
      return id === null ? [] : [[id, element] as const];
    }),
  );
  const concreteNumbering = new Map(
    descendants(numbering, 'num').flatMap((element) => {
      const id = integerAttribute(element, 'numId');
      return id === null ? [] : [[id, element] as const];
    }),
  );
  const seen = new Set<string>();
  const lists: ImportedDocxListMarker[] = [];
  for (const paragraph of descendants(document, 'p')) {
    const properties = directChild(paragraph, 'pPr');
    const numberingProperties = properties
      ? directChild(properties, 'numPr')
      : undefined;
    const numId = integerAttribute(
      numberingProperties ? directChild(numberingProperties, 'numId') : null,
      'val',
    );
    if (numId === null) continue;
    const level =
      integerAttribute(
        numberingProperties ? directChild(numberingProperties, 'ilvl') : null,
        'val',
      ) ?? 0;
    const key = `${numId}:${level}`;
    if (seen.has(key)) continue;
    seen.add(key);
    const concrete = concreteNumbering.get(numId);
    const abstractId = integerAttribute(
      concrete ? directChild(concrete, 'abstractNumId') : null,
      'val',
    );
    const abstract =
      abstractId === null ? undefined : abstractNumbering.get(abstractId);
    const start = effectiveListStart(concrete, abstract, level);
    const type = effectiveOrderedListType(concrete, abstract, level);
    if (start <= 1 && !type) continue;
    const marker = `__A3S_WORK_LIST_START_${lists.length + 1}__`;
    insertParagraphMarker(document, paragraph, marker);
    lists.push({ marker, start, ...(type ? { type } : {}) });
  }
  return { lists };
}

export function applyImportedDocxListMarkers(
  document: Document,
  markers: ImportedDocxListMarkers,
): void {
  const listByMarker = new Map(
    markers.lists.map((list) => [list.marker, list]),
  );
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_LIST_START_')) continue;
    const item = node.parentElement?.closest('li');
    const list = item?.parentElement;
    node.data = node.data.replace(LIST_START_MARKER_PATTERN, (marker) => {
      const imported = listByMarker.get(marker);
      if (
        imported &&
        list instanceof HTMLElement &&
        list.tagName.toLowerCase() === 'ol'
      ) {
        if (imported.start > 1) {
          list.setAttribute('start', String(imported.start));
        }
        if (imported.type) {
          list.setAttribute('type', imported.type);
        }
      }
      return '';
    });
  }
  document.body.normalize();
}

export function hasImportedDocxListMarkers(
  markers: ImportedDocxListMarkers,
): boolean {
  return markers.lists.length > 0;
}

function effectiveListStart(
  concrete: Element | undefined,
  abstract: Element | undefined,
  level: number,
): number {
  const override = concrete
    ? descendants(concrete, 'lvlOverride').find(
        (candidate) => integerAttribute(candidate, 'ilvl') === level,
      )
    : undefined;
  const overrideStart = integerAttribute(
    override ? directChild(override, 'startOverride') : null,
    'val',
  );
  if (overrideStart !== null) return overrideStart;
  const overrideLevel = override ? directChild(override, 'lvl') : undefined;
  const abstractLevel = abstract
    ? descendants(abstract, 'lvl').find(
        (candidate) => integerAttribute(candidate, 'ilvl') === level,
      )
    : undefined;
  return (
    integerAttribute(
      overrideLevel ? directChild(overrideLevel, 'start') : null,
      'val',
    ) ??
    integerAttribute(
      abstractLevel ? directChild(abstractLevel, 'start') : null,
      'val',
    ) ??
    1
  );
}

function effectiveOrderedListType(
  concrete: Element | undefined,
  abstract: Element | undefined,
  level: number,
): ImportedDocxOrderedListType | undefined {
  const override = concrete
    ? descendants(concrete, 'lvlOverride').find(
        (candidate) => integerAttribute(candidate, 'ilvl') === level,
      )
    : undefined;
  const overrideLevel = override ? directChild(override, 'lvl') : undefined;
  const abstractLevel = abstract
    ? descendants(abstract, 'lvl').find(
        (candidate) => integerAttribute(candidate, 'ilvl') === level,
      )
    : undefined;
  const format =
    stringAttribute(
      overrideLevel ? directChild(overrideLevel, 'numFmt') : null,
      'val',
    ) ??
    stringAttribute(
      abstractLevel ? directChild(abstractLevel, 'numFmt') : null,
      'val',
    );
  if (format === 'lowerLetter') return 'a';
  if (format === 'upperLetter') return 'A';
  if (format === 'lowerRoman') return 'i';
  if (format === 'upperRoman') return 'I';
  return undefined;
}

function insertParagraphMarker(
  document: Document,
  paragraph: Element,
  marker: string,
): void {
  const run = document.createElementNS(WORD_NAMESPACE, 'w:r');
  const text = document.createElementNS(WORD_NAMESPACE, 'w:t');
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  run.append(text);
  const properties = directChild(paragraph, 'pPr');
  paragraph.insertBefore(run, properties?.nextSibling ?? paragraph.firstChild);
}

function integerAttribute(
  element: Element | null | undefined,
  name: string,
): number | null {
  if (!element) return null;
  const source =
    attribute(element, name) ??
    (name.includes(':') ? null : attribute(element, `w:${name}`));
  if (source === null) return null;
  const value = Number(source);
  return Number.isSafeInteger(value) && value >= 0 ? value : null;
}

function stringAttribute(
  element: Element | null | undefined,
  name: string,
): string | null {
  if (!element) return null;
  const value =
    attribute(element, name) ??
    (name.includes(':') ? null : attribute(element, `w:${name}`));
  return value?.trim() || null;
}

function textNodes(root: ParentNode): Text[] {
  const document = root.ownerDocument;
  const walker = document?.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes: Text[] = [];
  if (!walker) return nodes;
  while (walker.nextNode()) nodes.push(walker.currentNode as Text);
  return nodes;
}
