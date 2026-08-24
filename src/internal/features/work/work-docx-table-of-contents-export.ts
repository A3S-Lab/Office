import JSZip from 'jszip';
import {
  documentTableOfContentsValueFromElement,
  type WorkDocumentTableOfContentsEntry,
  type WorkDocumentTableOfContentsOptions,
} from './work-document-table-of-contents';
import {
  attribute,
  descendants,
  directChild,
  directChildren,
  parseXml,
  xmlNamespacePrefix,
} from './work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';

export interface DocxTableOfContentsPatch {
  options: WorkDocumentTableOfContentsOptions;
  entries: WorkDocumentTableOfContentsEntry[];
}

export class DocxTableOfContentsPatchCollector {
  readonly patches: DocxTableOfContentsPatch[] = [];

  add(element: HTMLElement): DocxTableOfContentsPatch {
    const value = documentTableOfContentsValueFromElement(element);
    const patch = { options: value.options, entries: value.entries };
    this.patches.push(patch);
    return patch;
  }
}

export function docxTableOfContents(
  element: HTMLElement,
  docx: typeof import('docx'),
  collector: DocxTableOfContentsPatchCollector,
): InstanceType<typeof docx.TableOfContents> {
  const patch = collector.add(element);
  const options = patch.options;
  return new docx.TableOfContents('Table of Contents', {
    hyperlink: options.hyperlinks,
    headingStyleRange: `${options.minLevel}-${options.maxLevel}`,
    pageNumbersEntryLevelsRange: options.showPageNumbers
      ? undefined
      : `${options.minLevel}-${options.maxLevel}`,
    entryAndPageNumberSeparator:
      options.showPageNumbers && !options.rightAlignPageNumbers
        ? ' '
        : undefined,
    useAppliedParagraphOutlineLevel: true,
    hideTabAndPageNumbersInWebView: true,
    cachedEntries: patch.entries.map((entry) => ({
      title: entry.title,
      level: entry.level,
      page: options.showPageNumbers ? entry.pageNumber : undefined,
      href: options.hyperlinks ? entry.targetId : undefined,
    })),
  });
}

export async function patchDocxTableOfContents(
  buffer: ArrayBuffer,
  patches: readonly DocxTableOfContentsPatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  const archive = await JSZip.loadAsync(buffer);
  const entry = archive.file('word/document.xml');
  if (!entry) return buffer;
  const document = parseXml(
    await entry.async('text'),
    'generated DOCX word/document.xml',
  );
  const tables = descendants(document, 'sdt').filter((element) =>
    descendants(element, 'instrText').some((instruction) =>
      /^\s*TOC\b/i.test(instruction.textContent ?? ''),
    ),
  );
  if (tables.length !== patches.length) {
    throw new Error(
      `Generated DOCX emitted ${tables.length} table(s) of contents for ${patches.length} source block(s).`,
    );
  }
  for (const [index, table] of tables.entries()) {
    const patch = patches[index];
    if (patch) patchTableOfContents(document, table, patch.options);
  }
  addTableOfContentsBookmarks(document, patches);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

function patchTableOfContents(
  document: Document,
  table: Element,
  options: WorkDocumentTableOfContentsOptions,
): void {
  for (const tabStop of descendants(table, 'tab')) {
    if (tabStop.parentElement?.localName !== 'tabs') continue;
    const value = attribute(tabStop, 'val');
    if (value !== 'right') continue;
    if (!options.showPageNumbers || !options.rightAlignPageNumbers) {
      const parent = tabStop.parentElement;
      tabStop.remove();
      if (parent && !directChildren(parent, 'tab').length) parent.remove();
      continue;
    }
    const leader = docxTableOfContentsLeader(options.leader);
    if (leader) setWordAttribute(document, tabStop, 'leader', leader);
    else removeWordAttribute(tabStop, 'leader');
  }
  if (options.showPageNumbers && options.rightAlignPageNumbers) return;
  for (const tab of descendants(table, 'tab')) {
    if (tab.parentElement?.localName === 'tabs') continue;
    if (!options.showPageNumbers) {
      tab.remove();
      continue;
    }
    const text = document.createElementNS(
      WORD_NAMESPACE,
      wordQualifiedName(document.documentElement, 't'),
    );
    text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
    text.textContent = ' ';
    tab.replaceWith(text);
  }
}

function addTableOfContentsBookmarks(
  document: Document,
  patches: readonly DocxTableOfContentsPatch[],
): void {
  const names = new Set(
    descendants(document, 'bookmarkStart').flatMap((bookmark) => {
      const name = attribute(bookmark, 'name')?.trim();
      return name ? [name] : [];
    }),
  );
  let nextId =
    Math.max(
      0,
      ...descendants(document, 'bookmarkStart').map((bookmark) =>
        Number(attribute(bookmark, 'id') ?? 0),
      ),
    ) + 1;
  const targets = new Set(
    patches.flatMap((patch) =>
      patch.options.hyperlinks
        ? patch.entries.map((candidate) => candidate.targetId)
        : [],
    ),
  );
  for (const target of targets) {
    if (names.has(target)) continue;
    const paragraphId = /^heading-([0-9a-f]{8})$/i.exec(target)?.[1];
    if (!paragraphId) continue;
    const paragraph = descendants(document, 'p').find(
      (candidate) =>
        attribute(candidate, 'paraId')?.toLowerCase() ===
        paragraphId.toLowerCase(),
    );
    if (!paragraph) continue;
    const start = document.createElementNS(
      WORD_NAMESPACE,
      wordQualifiedName(document.documentElement, 'bookmarkStart'),
    );
    const end = document.createElementNS(
      WORD_NAMESPACE,
      wordQualifiedName(document.documentElement, 'bookmarkEnd'),
    );
    setWordAttribute(document, start, 'id', String(nextId));
    setWordAttribute(document, start, 'name', target);
    setWordAttribute(document, end, 'id', String(nextId));
    nextId += 1;
    const properties = directChild(paragraph, 'pPr');
    paragraph.insertBefore(
      start,
      properties?.nextSibling ?? paragraph.firstChild,
    );
    paragraph.append(end);
    names.add(target);
  }
}

function docxTableOfContentsLeader(
  value: WorkDocumentTableOfContentsOptions['leader'],
): string | null {
  if (value === 'dash') return 'hyphen';
  if (value === 'underline') return 'underscore';
  if (value === 'dot') return 'dot';
  return null;
}

function setWordAttribute(
  document: Document,
  element: Element,
  localName: string,
  value: string,
): void {
  removeWordAttribute(element, localName);
  element.setAttributeNS(
    WORD_NAMESPACE,
    wordQualifiedName(document.documentElement, localName),
    value,
  );
}

function removeWordAttribute(element: Element, localName: string): void {
  for (const candidate of Array.from(element.attributes)) {
    if (
      candidate.localName === localName &&
      (candidate.namespaceURI === WORD_NAMESPACE || !candidate.namespaceURI)
    ) {
      element.removeAttributeNode(candidate);
    }
  }
}

function wordQualifiedName(root: Element, localName: string): string {
  return `${xmlNamespacePrefix(root, WORD_NAMESPACE) ?? 'w'}:${localName}`;
}
