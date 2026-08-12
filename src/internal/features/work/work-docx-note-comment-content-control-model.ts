import {
  readStaticContentControlDefinition,
  type StaticContentControlDefinition,
} from './work-docx-note-comment-content-control-properties';
import {
  hasOnlyPassiveContentControlAttributes as validPassiveContainer,
  hasUnsupportedContentControlSemanticChild as hasUnsupportedSemanticChild,
  isContentControlSemanticNamespace as isSemanticNamespace,
  isDocxWordElement as isWordElement,
  wordDirectChildren,
} from './work-docx-note-comment-content-control-xml';
import {
  XML_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';

const MAX_PARAGRAPHS = 65_536;
const MAX_RUNS = 262_144;
const MAX_CONTROLS = 65_536;
const MAX_RUN_SEGMENTS = 4_096;

type DocumentRole = 'generated' | 'source';

export interface ContentControlLimits {
  generatedControls: number;
  generatedParagraphs: number;
  generatedRuns: number;
  sourceControls: number;
  sourceParagraphs: number;
  sourceRuns: number;
}

export interface StaticContentRunRecord {
  element: Element;
  end: number;
  start: number;
  text: string;
}

export interface InlineContentControlRecord
  extends StaticContentControlDefinition {
  end: number;
  runs: StaticContentRunRecord[];
  start: number;
  text: string;
}

export interface StaticContentParagraphRecord {
  controls: InlineContentControlRecord[];
  element: Element;
  runs: StaticContentRunRecord[];
  text: string;
}

export interface BlockContentControlRecord
  extends StaticContentControlDefinition {
  paragraphs: StaticContentParagraphRecord[];
}

export function createContentControlLimits(): ContentControlLimits {
  return {
    generatedControls: 0,
    generatedParagraphs: 0,
    generatedRuns: 0,
    sourceControls: 0,
    sourceParagraphs: 0,
    sourceRuns: 0,
  };
}

export function readStaticInlineParagraph(
  paragraph: Element,
  role: DocumentRole,
  limits: ContentControlLimits,
): StaticContentParagraphRecord | null {
  return readParagraph(paragraph, role, limits, role === 'source');
}

export function readStaticBlockControl(
  control: Element,
  limits: ContentControlLimits,
): BlockContentControlRecord | null {
  const definition = readStaticContentControlDefinition(control);
  if (!definition || !validPassiveContainer(definition.content)) return null;
  const wordChildren = Array.from(definition.content.children).filter(
    isWordElement,
  );
  if (
    !wordChildren.length ||
    wordChildren.some((child) => child.localName !== 'p') ||
    hasUnsupportedSemanticChild(definition.content)
  ) {
    return null;
  }
  incrementLimit(limits, 'sourceControls', 1, 'content-control');
  const paragraphs: StaticContentParagraphRecord[] = [];
  for (const paragraph of wordChildren) {
    const record = readParagraph(paragraph, 'source', limits, false);
    if (!record) return null;
    paragraphs.push(record);
  }
  if (
    !paragraphs.some((paragraph) => paragraph.text) ||
    (definition.type === 'text' &&
      (paragraphs.length !== 1 ||
        paragraphs[0].runs.length !== 1 ||
        (!definition.multiLine && paragraphs[0].text.includes('\n'))))
  ) {
    return null;
  }
  return { ...definition, paragraphs };
}

export function directWordParagraphs(scope: Element): Element[] {
  return Array.from(scope.children).filter(
    (child) => child.localName === 'p' && isWordElement(child),
  );
}

export function directWordControls(scope: Element): Element[] {
  return Array.from(scope.children).filter(
    (child) => child.localName === 'sdt' && isWordElement(child),
  );
}

export function staticRunSpanKey(run: StaticContentRunRecord): string {
  return `${run.start}:${run.end}:${run.text}`;
}

export function alignStaticRunBoundaries(
  paragraph: Element,
  boundaries: readonly number[],
): boolean {
  const runs = directRunRecords(paragraph);
  if (!runs || boundaries.length > MAX_RUN_SEGMENTS) return false;
  const text = runs.map((run) => run.text).join('');
  const allBoundaries = Array.from(new Set(boundaries)).sort(
    (left, right) => left - right,
  );
  if (
    allBoundaries.some(
      (boundary) =>
        boundary < 0 ||
        boundary > text.length ||
        (boundary > 0 &&
          boundary < text.length &&
          splitsSurrogatePair(text, boundary)),
    )
  ) {
    return false;
  }
  const unique = allBoundaries.filter(
    (boundary) => boundary > 0 && boundary < text.length,
  );
  for (const run of runs) {
    const internal = unique.filter(
      (boundary) => boundary > run.start && boundary < run.end,
    );
    if (!internal.length) continue;
    if (!isSimpleTextRun(run.element) || !run.element.parentNode) return false;
    const offsets = [run.start, ...internal, run.end];
    const parent = run.element.parentNode;
    for (let index = 0; index < offsets.length - 1; index += 1) {
      const clone = run.element.cloneNode(true) as Element;
      const textElement = wordDirectChildren(clone, 't')[0];
      if (!textElement) return false;
      const value = text.slice(offsets[index], offsets[index + 1]);
      textElement.textContent = value;
      setTextSpacePreservation(textElement, value);
      parent.insertBefore(clone, run.element);
    }
    run.element.remove();
  }
  return true;
}

function readParagraph(
  paragraph: Element,
  role: DocumentRole,
  limits: ContentControlLimits,
  allowControls: boolean,
): StaticContentParagraphRecord | null {
  if (!isWordElement(paragraph) || paragraph.localName !== 'p') return null;
  incrementLimit(limits, `${role}Paragraphs`, 1, 'paragraph');
  const runs: StaticContentRunRecord[] = [];
  const controls: InlineContentControlRecord[] = [];
  let offset = 0;
  let properties = 0;
  let sawContent = false;
  for (const child of Array.from(paragraph.children)) {
    if (!isWordElement(child)) {
      if (isSemanticNamespace(child.namespaceURI ?? '')) return null;
      continue;
    }
    if (child.localName === 'pPr') {
      properties += 1;
      if (properties > 1 || sawContent) return null;
      continue;
    }
    sawContent = true;
    if (child.localName === 'r') {
      const record = readRun(child, offset, role, limits);
      if (!record) return null;
      offset = record.end;
      if (record.text) runs.push(record);
      continue;
    }
    if (child.localName !== 'sdt' || !allowControls) return null;
    const definition = readStaticContentControlDefinition(child);
    if (!definition || !validPassiveContainer(definition.content)) return null;
    const contentChildren = Array.from(definition.content.children).filter(
      isWordElement,
    );
    if (
      !contentChildren.length ||
      contentChildren.some((item) => item.localName !== 'r') ||
      hasUnsupportedSemanticChild(definition.content)
    ) {
      return null;
    }
    incrementLimit(limits, `${role}Controls`, 1, 'content-control');
    const start = offset;
    const controlRuns: StaticContentRunRecord[] = [];
    for (const run of contentChildren) {
      const record = readRun(run, offset, role, limits);
      if (!record) return null;
      offset = record.end;
      if (record.text) {
        runs.push(record);
        controlRuns.push(record);
      }
    }
    const text = controlRuns.map((run) => run.text).join('');
    if (
      !text ||
      (definition.type === 'text' &&
        (controlRuns.length !== 1 ||
          (!definition.multiLine && text.includes('\n'))))
    ) {
      return null;
    }
    controls.push({
      ...definition,
      end: offset,
      runs: controlRuns,
      start,
      text,
    });
  }
  return {
    controls,
    element: paragraph,
    runs,
    text: runs.map((run) => run.text).join(''),
  };
}

function readRun(
  run: Element,
  start: number,
  role: DocumentRole,
  limits: ContentControlLimits,
): StaticContentRunRecord | null {
  incrementLimit(limits, `${role}Runs`, 1, 'run');
  const text = contentRunText(run);
  return text === null
    ? null
    : { element: run, end: start + text.length, start, text };
}

function directRunRecords(paragraph: Element): StaticContentRunRecord[] | null {
  const records: StaticContentRunRecord[] = [];
  let offset = 0;
  let properties = 0;
  let sawContent = false;
  for (const child of Array.from(paragraph.children)) {
    if (!isWordElement(child)) {
      if (isSemanticNamespace(child.namespaceURI ?? '')) return null;
      continue;
    }
    if (child.localName === 'pPr') {
      properties += 1;
      if (properties > 1 || sawContent) return null;
      continue;
    }
    if (child.localName !== 'r') return null;
    sawContent = true;
    const text = contentRunText(child);
    if (text === null) return null;
    const start = offset;
    offset += text.length;
    if (text) records.push({ element: child, end: offset, start, text });
  }
  return records;
}

function contentRunText(run: Element): string | null {
  let text = '';
  let properties = 0;
  let sawContent = false;
  for (const child of Array.from(run.children)) {
    if (!isWordElement(child)) {
      if (isSemanticNamespace(child.namespaceURI ?? '')) return null;
      continue;
    }
    if (child.localName === 'rPr') {
      properties += 1;
      if (properties > 1 || sawContent) return null;
    } else if (child.localName === 't' || child.localName === 'delText') {
      sawContent = true;
      text += child.textContent ?? '';
    } else if (child.localName === 'tab') {
      sawContent = true;
      text += '\t';
    } else if (child.localName === 'br' || child.localName === 'cr') {
      sawContent = true;
      text += '\n';
    } else if (child.localName === 'noBreakHyphen') {
      sawContent = true;
      text += '\u2011';
    } else if (child.localName === 'softHyphen') {
      sawContent = true;
      text += '\u00ad';
    } else if (
      child.localName !== 'footnoteRef' &&
      child.localName !== 'endnoteRef'
    ) {
      return null;
    } else {
      sawContent = true;
    }
  }
  return text;
}

function isSimpleTextRun(run: Element): boolean {
  const children = Array.from(run.children).filter(isWordElement);
  return (
    children.filter((child) => child.localName === 't').length === 1 &&
    children.every(
      (child) => child.localName === 'rPr' || child.localName === 't',
    )
  );
}

function setTextSpacePreservation(element: Element, text: string): void {
  for (const item of Array.from(element.attributes)) {
    if (
      xmlAttributeLocalName(item) === 'space' &&
      xmlAttributeNamespace(element, item) === XML_NAMESPACE
    ) {
      element.removeAttributeNode(item);
    }
  }
  if (/^\s|\s$/u.test(text)) {
    element.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  }
}

function incrementLimit(
  limits: ContentControlLimits,
  key: keyof ContentControlLimits,
  count: number,
  label: string,
): void {
  limits[key] += count;
  const maximum =
    label === 'run'
      ? MAX_RUNS
      : label === 'paragraph'
        ? MAX_PARAGRAPHS
        : MAX_CONTROLS;
  if (limits[key] > maximum) {
    const role = key.startsWith('source') ? 'Registered source' : 'Generated';
    throw new Error(
      `${role} DOCX exceeds the stable note/comment ${label} limit.`,
    );
  }
}

function splitsSurrogatePair(text: string, offset: number): boolean {
  const before = text.charCodeAt(offset - 1);
  const after = text.charCodeAt(offset);
  return (
    before >= 0xd800 && before <= 0xdbff && after >= 0xdc00 && after <= 0xdfff
  );
}
