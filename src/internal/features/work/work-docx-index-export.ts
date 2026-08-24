import type { FileChild, ParagraphChild } from 'docx';
import JSZip from 'jszip';
import {
  documentIndexEntryFromElement,
  documentIndexValueFromElement,
  type WorkDocumentIndexGeneratedEntry,
  type WorkDocumentIndexOptions,
} from './work-document-index';
import {
  documentIndexEntryInstruction,
  documentIndexInstruction,
} from './work-document-index-fields';
import {
  descendants,
  parseXml,
  xmlNamespacePrefix,
} from './work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';

export interface DocxIndexPatch {
  marker: string;
  options: WorkDocumentIndexOptions;
  entries: WorkDocumentIndexGeneratedEntry[];
}

export class DocxIndexPatchCollector {
  readonly patches: DocxIndexPatch[] = [];

  add(element: HTMLElement): DocxIndexPatch {
    const value = documentIndexValueFromElement(element);
    const patch: DocxIndexPatch = {
      marker: `__A3S_WORK_DOCUMENT_INDEX_${this.patches.length + 1}__`,
      options: value.options,
      entries: value.entries,
    };
    this.patches.push(patch);
    return patch;
  }
}

export function docxDocumentIndex(
  element: HTMLElement,
  docx: typeof import('docx'),
  collector: DocxIndexPatchCollector,
): FileChild[] {
  const patch = collector.add(element);
  return [
    new docx.Paragraph({
      children: [new docx.TextRun(patch.marker)],
    }),
  ];
}

export function docxIndexEntryRun(
  element: HTMLElement,
  docx: typeof import('docx'),
): ParagraphChild | null {
  const value = documentIndexEntryFromElement(element);
  return value
    ? new docx.SimpleField(documentIndexEntryInstruction(value))
    : null;
}

export async function patchDocxIndexes(
  buffer: ArrayBuffer,
  patches: readonly DocxIndexPatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  const archive = await JSZip.loadAsync(buffer);
  const entry = archive.file('word/document.xml');
  if (!entry) return buffer;
  const document = parseXml(
    await entry.async('text'),
    'generated DOCX word/document.xml',
  );
  const paragraphs = descendants(document, 'p');
  for (const patch of patches) {
    const paragraph = paragraphs.find((candidate) =>
      descendants(candidate, 't').some((text) =>
        (text.textContent ?? '').includes(patch.marker),
      ),
    );
    if (!paragraph) {
      throw new Error(
        `Generated DOCX did not emit the marker for document index ${patch.marker}.`,
      );
    }
    paragraph.replaceWith(createIndexContentControl(document, patch));
  }
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

function createIndexContentControl(
  document: Document,
  patch: DocxIndexPatch,
): Element {
  const control = wordElement(document, 'sdt');
  const properties = wordElement(document, 'sdtPr');
  const documentPart = wordElement(document, 'docPartObj');
  const gallery = wordElement(document, 'docPartGallery');
  setWordAttribute(document, gallery, 'val', 'Indexes');
  documentPart.append(gallery, wordElement(document, 'docPartUnique'));
  properties.append(documentPart);
  const content = wordElement(document, 'sdtContent');
  const fieldStart = indexFieldStartParagraph(document, patch.options);
  content.append(fieldStart);

  const cached = cachedIndexParagraphs(document, patch);
  if (!cached.length) cached.push(emptyIndexParagraph(document));
  const last = cached.at(-1);
  last?.append(fieldCharacterRun(document, 'end'));
  content.append(...cached);
  control.append(properties, content);
  return control;
}

function indexFieldStartParagraph(
  document: Document,
  options: WorkDocumentIndexOptions,
): Element {
  const paragraph = wordElement(document, 'p');
  paragraph.append(fieldCharacterRun(document, 'begin'));
  const instructionRun = wordElement(document, 'r');
  const instruction = wordElement(document, 'instrText');
  instruction.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  instruction.textContent = ` ${documentIndexInstruction(options)} `;
  instructionRun.append(instruction);
  paragraph.append(instructionRun, fieldCharacterRun(document, 'separate'));
  return paragraph;
}

function cachedIndexParagraphs(
  document: Document,
  patch: DocxIndexPatch,
): Element[] {
  const paragraphs: Element[] = [];
  let currentMain = '';
  for (const entry of patch.entries) {
    if (entry.subEntry && entry.mainEntry !== currentMain) {
      paragraphs.push(
        cachedIndexParagraph(document, patch.options, {
          style: 'Index1',
          term: entry.mainEntry,
          crossReference: '',
          pages: [],
        }),
      );
    }
    paragraphs.push(
      cachedIndexParagraph(document, patch.options, {
        style: entry.subEntry ? 'Index2' : 'Index1',
        term: entry.subEntry || entry.mainEntry,
        crossReference: entry.crossReference,
        pages: entry.pages,
      }),
    );
    currentMain = entry.mainEntry;
  }
  return paragraphs;
}

function cachedIndexParagraph(
  document: Document,
  options: WorkDocumentIndexOptions,
  source: {
    style: 'Index1' | 'Index2';
    term: string;
    crossReference: string;
    pages: WorkDocumentIndexGeneratedEntry['pages'];
  },
): Element {
  const paragraph = wordElement(document, 'p');
  const properties = wordElement(document, 'pPr');
  const style = wordElement(document, 'pStyle');
  setWordAttribute(document, style, 'val', source.style);
  properties.append(style);
  if (source.pages.length && options.rightAlignPageNumbers) {
    const tabs = wordElement(document, 'tabs');
    const tab = wordElement(document, 'tab');
    setWordAttribute(document, tab, 'val', 'right');
    const leader = docxIndexLeader(options.leader);
    if (leader) setWordAttribute(document, tab, 'leader', leader);
    setWordAttribute(document, tab, 'pos', '9360');
    tabs.append(tab);
    properties.append(tabs);
  }
  paragraph.append(properties, textRun(document, source.term));
  if (source.crossReference) {
    paragraph.append(textRun(document, `, See ${source.crossReference}`));
    return paragraph;
  }
  if (!source.pages.length) return paragraph;
  paragraph.append(
    options.rightAlignPageNumbers ? tabRun(document) : textRun(document, ', '),
  );
  for (const [index, page] of source.pages.entries()) {
    if (index > 0) paragraph.append(textRun(document, ', '));
    paragraph.append(
      textRun(document, String(page.pageNumber), {
        bold: page.pageBold,
        italic: page.pageItalic,
      }),
    );
  }
  return paragraph;
}

function emptyIndexParagraph(document: Document): Element {
  const paragraph = wordElement(document, 'p');
  const properties = wordElement(document, 'pPr');
  const style = wordElement(document, 'pStyle');
  setWordAttribute(document, style, 'val', 'Index1');
  properties.append(style);
  paragraph.append(properties);
  return paragraph;
}

function textRun(
  document: Document,
  value: string,
  style: { bold?: boolean; italic?: boolean } = {},
): Element {
  const run = wordElement(document, 'r');
  if (style.bold || style.italic) {
    const properties = wordElement(document, 'rPr');
    if (style.bold) properties.append(wordElement(document, 'b'));
    if (style.italic) properties.append(wordElement(document, 'i'));
    run.append(properties);
  }
  const text = wordElement(document, 't');
  if (/^\s|\s$/.test(value)) {
    text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  }
  text.textContent = value;
  run.append(text);
  return run;
}

function tabRun(document: Document): Element {
  const run = wordElement(document, 'r');
  run.append(wordElement(document, 'tab'));
  return run;
}

function fieldCharacterRun(document: Document, type: string): Element {
  const run = wordElement(document, 'r');
  const character = wordElement(document, 'fldChar');
  setWordAttribute(document, character, 'fldCharType', type);
  run.append(character);
  return run;
}

function docxIndexLeader(
  value: WorkDocumentIndexOptions['leader'],
): string | null {
  if (value === 'dash') return 'hyphen';
  if (value === 'underline') return 'underscore';
  if (value === 'dot') return 'dot';
  return null;
}

function wordElement(document: Document, localName: string): Element {
  return document.createElementNS(
    WORD_NAMESPACE,
    wordQualifiedName(document.documentElement, localName),
  );
}

function setWordAttribute(
  document: Document,
  element: Element,
  localName: string,
  value: string,
): void {
  for (const candidate of Array.from(element.attributes)) {
    if (
      candidate.localName === localName &&
      (candidate.namespaceURI === WORD_NAMESPACE || !candidate.namespaceURI)
    ) {
      element.removeAttributeNode(candidate);
    }
  }
  element.setAttributeNS(
    WORD_NAMESPACE,
    wordQualifiedName(document.documentElement, localName),
    value,
  );
}

function wordQualifiedName(root: Element, localName: string): string {
  return `${xmlNamespacePrefix(root, WORD_NAMESPACE) ?? 'w'}:${localName}`;
}
