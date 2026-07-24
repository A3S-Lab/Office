import {
  type DocumentTabAlignment,
  type DocumentTabLeader,
  type DocumentTabStop,
  normalizeDocumentTabStops,
  serializeDocumentTabStops,
} from './work-document-tab-stops';
import {
  type DocxParagraphStyleSource,
  docxParagraphPropertySources,
  resolveDocxParagraphStyleResolver,
} from './work-docx-paragraph-styles';
import {
  attribute,
  descendants,
  directChild,
  directChildren,
} from './work-ooxml-package';

export interface ImportedDocxParagraphTabStopMarker {
  marker: string;
  tabStops: DocumentTabStop[];
}

export interface ImportedDocxParagraphTabStopMarkers {
  paragraphs: ImportedDocxParagraphTabStopMarker[];
  inlineTabs: string[];
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const TWIPS_PER_CSS_PIXEL = 15;
const PARAGRAPH_TAB_STOP_MARKER_PATTERN =
  /__A3S_WORK_PARAGRAPH_TAB_STOPS_\d+__/g;
const INLINE_TAB_MARKER_PATTERN = /__A3S_WORK_INLINE_TAB_\d+__/g;

export function markDocxParagraphTabStops(
  document: Document,
  styleSource?: DocxParagraphStyleSource,
): ImportedDocxParagraphTabStopMarkers {
  const paragraphs: ImportedDocxParagraphTabStopMarker[] = [];
  const inlineTabs: string[] = [];
  const styles = resolveDocxParagraphStyleResolver(styleSource);

  for (const paragraph of descendants(document, 'p')) {
    let properties = directChild(paragraph, 'pPr');
    const tabStops = resolvedParagraphTabStops(
      docxParagraphPropertySources(properties, styles),
    );
    if (tabStops.length) {
      properties ??= insertParagraphProperties(document, paragraph);
      const marker = `__A3S_WORK_PARAGRAPH_TAB_STOPS_${paragraphs.length + 1}__`;
      insertParagraphMarker(document, paragraph, properties, marker);
      paragraphs.push({ marker, tabStops });
    }

    for (const tab of inlineRunTabs(paragraph)) {
      const marker = `__A3S_WORK_INLINE_TAB_${inlineTabs.length + 1}__`;
      const text = document.createElementNS(WORD_NAMESPACE, 'w:t');
      text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
      text.textContent = marker;
      tab.replaceWith(text);
      inlineTabs.push(marker);
    }
  }
  return { paragraphs, inlineTabs };
}

export function applyImportedDocxParagraphTabStopMarkers(
  document: Document,
  markers: ImportedDocxParagraphTabStopMarkers,
): void {
  const tabStopsByMarker = new Map(
    markers.paragraphs.map(({ marker, tabStops }) => [marker, tabStops]),
  );
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_PARAGRAPH_TAB_STOPS_')) continue;
    const block = closestParagraphBlock(node.parentElement);
    node.data = node.data.replace(
      PARAGRAPH_TAB_STOP_MARKER_PATTERN,
      (marker) => {
        const tabStops = tabStopsByMarker.get(marker);
        if (block && tabStops?.length) {
          block.dataset.officeTabStops = serializeDocumentTabStops(tabStops);
        }
        return '';
      },
    );
  }

  const inlineMarkers = new Set(markers.inlineTabs);
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_INLINE_TAB_')) continue;
    replaceInlineTabMarkers(node, inlineMarkers);
  }
  document.body.normalize();
}

export function hasImportedDocxParagraphTabStopMarkers(
  markers: ImportedDocxParagraphTabStopMarkers,
): boolean {
  return markers.paragraphs.length > 0 || markers.inlineTabs.length > 0;
}

function resolvedParagraphTabStops(
  sources: readonly Element[],
): DocumentTabStop[] {
  const byPosition = new Map<number, DocumentTabStop>();
  for (const properties of sources) {
    const tabs = directChild(properties, 'tabs');
    if (!tabs) continue;
    for (const tab of directChildren(tabs, 'tab')) {
      const positionTwips = numericAttribute(tab, 'pos');
      if (positionTwips === undefined || positionTwips <= 0) continue;
      const position = positionTwips / TWIPS_PER_CSS_PIXEL;
      const value = wordAttribute(tab, 'val')?.trim();
      if (value === 'clear' || value === 'bar') {
        byPosition.delete(position);
        continue;
      }
      byPosition.set(position, {
        position,
        alignment: docxTabAlignment(value),
        leader: docxTabLeader(wordAttribute(tab, 'leader')),
      });
    }
  }
  return normalizeDocumentTabStops(Array.from(byPosition.values()));
}

function inlineRunTabs(paragraph: Element): Element[] {
  return descendants(paragraph, 'tab').filter((tab) => {
    let parent = tab.parentElement;
    while (parent && parent !== paragraph) {
      if (parent.localName === 'tabs') return false;
      parent = parent.parentElement;
    }
    return parent === paragraph;
  });
}

function replaceInlineTabMarkers(node: Text, markers: ReadonlySet<string>) {
  const matches = Array.from(node.data.matchAll(INLINE_TAB_MARKER_PATTERN));
  if (!matches.length) return;
  const fragment = node.ownerDocument.createDocumentFragment();
  let offset = 0;
  for (const match of matches) {
    const marker = match[0];
    const index = match.index ?? 0;
    if (index > offset) fragment.append(node.data.slice(offset, index));
    if (markers.has(marker)) {
      const tab = node.ownerDocument.createElement('span');
      tab.dataset.documentTab = 'true';
      tab.className = 'work-document-tab';
      tab.setAttribute('role', 'separator');
      tab.setAttribute('aria-label', '制表符');
      fragment.append(tab);
    } else {
      fragment.append(marker);
    }
    offset = index + marker.length;
  }
  if (offset < node.data.length) fragment.append(node.data.slice(offset));
  node.replaceWith(fragment);
}

function insertParagraphMarker(
  document: Document,
  paragraph: Element,
  properties: Element,
  marker: string,
): void {
  const run = document.createElementNS(WORD_NAMESPACE, 'w:r');
  const text = document.createElementNS(WORD_NAMESPACE, 'w:t');
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  run.append(text);
  paragraph.insertBefore(run, properties.nextSibling);
}

function insertParagraphProperties(
  document: Document,
  paragraph: Element,
): Element {
  const properties = document.createElementNS(WORD_NAMESPACE, 'w:pPr');
  paragraph.insertBefore(properties, paragraph.firstChild);
  return properties;
}

function closestParagraphBlock(element: Element | null): HTMLElement | null {
  const block = element?.closest('p, h1, h2, h3, h4, h5, h6, blockquote');
  return block instanceof HTMLElement ? block : null;
}

function docxTabAlignment(
  value: string | null | undefined,
): DocumentTabAlignment {
  if (value === 'center') return 'center';
  if (value === 'right' || value === 'end') return 'right';
  if (value === 'decimal' || value === 'num') return 'decimal';
  return 'left';
}

function docxTabLeader(value: string | null): DocumentTabLeader {
  if (value === 'dot') return 'dot';
  if (value === 'hyphen') return 'hyphen';
  if (value === 'underscore' || value === 'heavy') return 'underscore';
  if (value === 'middleDot') return 'middleDot';
  return 'none';
}

function numericAttribute(element: Element, name: string): number | undefined {
  const raw = wordAttribute(element, name);
  if (raw === null || !raw.trim()) return undefined;
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

function wordAttribute(element: Element, name: string): string | null {
  return attribute(element, name) ?? attribute(element, `w:${name}`);
}

function textNodes(root: ParentNode): Text[] {
  const document = root.ownerDocument;
  const walker = document?.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes: Text[] = [];
  if (!walker) return nodes;
  while (walker.nextNode()) nodes.push(walker.currentNode as Text);
  return nodes;
}
