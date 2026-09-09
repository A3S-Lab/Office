import JSZip from 'jszip';
import { parseDocumentSectionFormatting } from './work-document-section-format-changes';
import { descendants, directChild, parseXml } from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';
import type { WorkDocumentSectionLayout } from './work-types';

interface DocxSectionFormattingChangePatch {
  id: number;
  author: string;
  date: string;
  before: string;
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MAX_SECTION_FORMATTING_CHANGE_PATCHES = 512;

export class DocxSectionFormattingChangePatchCollector {
  readonly patches: Array<DocxSectionFormattingChangePatch | null> = [];

  record(layout: WorkDocumentSectionLayout, id: number): void {
    const change = layout.formattingChange;
    if (!change || change.kind !== 'section-formatting') {
      this.patches.push(null);
      return;
    }
    const author = change.author.trim();
    const date = normalizedRevisionDate(change.date);
    if (!author || author.length > 255 || !parseDocumentSectionFormatting(change.before)) {
      throw new Error('Document contains an invalid section-formatting revision.');
    }
    if (
      this.patches.filter(Boolean).length >= MAX_SECTION_FORMATTING_CHANGE_PATCHES
    ) {
      throw new Error('Document exceeds the section-formatting revision limit.');
    }
    this.patches.push({ id, author, date, before: change.before });
  }
}

export async function patchDocxSectionFormattingChanges(
  buffer: ArrayBuffer,
  patches: readonly (DocxSectionFormattingChangePatch | null)[],
): Promise<ArrayBuffer> {
  if (!patches.some(Boolean)) return buffer;
  if (patches.filter(Boolean).length > MAX_SECTION_FORMATTING_CHANGE_PATCHES) {
    throw new Error('Document exceeds the section-formatting revision limit.');
  }
  const archive = await JSZip.loadAsync(buffer);
  const entry = archive.file('word/document.xml');
  if (!entry) return buffer;
  const document = parseXml(
    decodeXmlBytes(
      await entry.async('uint8array'),
      'generated DOCX word/document.xml',
    ),
    'generated DOCX word/document.xml',
  );
  const sections = descendants(document, 'sectPr').filter(
    (element) =>
      element.namespaceURI === WORD_NAMESPACE &&
      element.parentElement?.localName !== 'sectPrChange',
  );
  let changed = false;
  let index = 0;
  for (const section of sections) {
    const patch = patches[index++] ?? null;
    if (!patch) continue;
    setSectionFormattingChange(document, section, patch);
    changed = true;
  }
  if (index !== patches.length) {
    throw new Error(
      `DOCX section-formatting revision patch count mismatch (${patches.length} patches, ${index} sections).`,
    );
  }
  if (changed) {
    archive.file('word/document.xml', serializeUtf8Xml(document));
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function setSectionFormattingChange(
  document: Document,
  section: Element,
  patch: DocxSectionFormattingChangePatch,
): void {
  const formatting = parseDocumentSectionFormatting(patch.before);
  if (!formatting) {
    throw new Error('Document contains an invalid section-formatting revision.');
  }
  for (const existing of Array.from(section.children).filter(
    (child) =>
      child.localName === 'sectPrChange' &&
      child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  const change = document.createElementNS(WORD_NAMESPACE, 'w:sectPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', String(patch.id));
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', patch.author);
  if (patch.date) {
    change.setAttributeNS(WORD_NAMESPACE, 'w:date', patch.date);
  }
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:sectPr');
  if (formatting.pageGeometry) {
    const pageSize = document.createElementNS(WORD_NAMESPACE, 'w:pgSz');
    pageSize.setAttributeNS(
      WORD_NAMESPACE,
      'w:w',
      String(formatting.pageGeometry.width),
    );
    pageSize.setAttributeNS(
      WORD_NAMESPACE,
      'w:h',
      String(formatting.pageGeometry.height),
    );
    if (formatting.pageGeometry.orientation) {
      pageSize.setAttributeNS(
        WORD_NAMESPACE,
        'w:orient',
        formatting.pageGeometry.orientation,
      );
    }
    if (formatting.pageGeometry.code !== undefined) {
      pageSize.setAttributeNS(
        WORD_NAMESPACE,
        'w:code',
        String(formatting.pageGeometry.code),
      );
    }
    prior.append(pageSize);
  } else if (formatting.orientation) {
    const pageSize = document.createElementNS(WORD_NAMESPACE, 'w:pgSz');
    pageSize.setAttributeNS(WORD_NAMESPACE, 'w:orient', formatting.orientation);
    prior.append(pageSize);
  }
  if (formatting.pageMargins) {
    const margins = document.createElementNS(WORD_NAMESPACE, 'w:pgMar');
    for (const key of [
      'top',
      'right',
      'bottom',
      'left',
      'header',
      'footer',
      'gutter',
    ] as const) {
      margins.setAttributeNS(
        WORD_NAMESPACE,
        `w:${key}`,
        String(formatting.pageMargins[key]),
      );
    }
    prior.append(margins);
  }
  if (formatting.paperSource) {
    const paperSource = document.createElementNS(WORD_NAMESPACE, 'w:paperSrc');
    if (formatting.paperSource.first !== undefined) {
      paperSource.setAttributeNS(
        WORD_NAMESPACE,
        'w:first',
        String(formatting.paperSource.first),
      );
    }
    if (formatting.paperSource.other !== undefined) {
      paperSource.setAttributeNS(
        WORD_NAMESPACE,
        'w:other',
        String(formatting.paperSource.other),
      );
    }
    prior.append(paperSource);
  }
  if (formatting.columns) {
    const columns = document.createElementNS(WORD_NAMESPACE, 'w:cols');
    columns.setAttributeNS(
      WORD_NAMESPACE,
      'w:num',
      String(formatting.columns.count),
    );
    columns.setAttributeNS(
      WORD_NAMESPACE,
      'w:space',
      String(Math.round((formatting.columns.spacing * 1440) / 25.4)),
    );
    if (formatting.columns.separator) {
      columns.setAttributeNS(WORD_NAMESPACE, 'w:sep', '1');
    }
    prior.append(columns);
  }
  if (formatting.differentFirstPage !== undefined) {
    const titlePg = document.createElementNS(WORD_NAMESPACE, 'w:titlePg');
    if (!formatting.differentFirstPage) {
      titlePg.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
    }
    prior.append(titlePg);
  }
  if (formatting.rtlGutter !== undefined) {
    const rtlGutter = document.createElementNS(WORD_NAMESPACE, 'w:rtlGutter');
    if (!formatting.rtlGutter) {
      rtlGutter.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
    }
    prior.append(rtlGutter);
  }
  change.append(prior);
  section.append(change);
}

function normalizedRevisionDate(value: string | undefined): string {
  if (!value?.trim()) return '';
  const time = Date.parse(value);
  return Number.isFinite(time) ? new Date(time).toISOString() : '';
}
