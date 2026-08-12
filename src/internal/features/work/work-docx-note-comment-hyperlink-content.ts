import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  readDocxHyperlinkDestination,
  type DocxHyperlinkDestination,
  type DocxHyperlinkRelationshipState,
} from './work-docx-note-comment-hyperlink-relationships';
import { descendants } from './work-ooxml-package';
import {
  XML_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';

const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const MAX_PARAGRAPHS = 65_536;
const MAX_RUNS = 262_144;
const MAX_COMMENT_RUN_SEGMENTS = 4_096;

type DocumentRole = 'generated' | 'source';

export interface NoteCommentHyperlinkContentLimits {
  generatedParagraphs: number;
  generatedRuns: number;
  sourceParagraphs: number;
  sourceRuns: number;
}

export interface NoteCommentParagraphRecord {
  element: Element;
  hyperlinks: NoteCommentHyperlinkRecord[];
  identity: string;
  runs: NoteCommentRunRecord[];
  text: string;
}

type ParagraphContent = Omit<
  NoteCommentParagraphRecord,
  'element' | 'identity'
>;

export interface NoteCommentHyperlinkRecord {
  destination: DocxHyperlinkDestination | null;
  element: Element;
  end: number;
  start: number;
  text: string;
}

export interface NoteCommentRunRecord {
  container: Element | null;
  element: Element;
  end: number;
  start: number;
  text: string;
}

export function noteCommentContentParagraphs(
  scope: Element,
  role: DocumentRole,
  relationships: DocxHyperlinkRelationshipState,
  limits: NoteCommentHyperlinkContentLimits,
): NoteCommentParagraphRecord[] {
  const paragraphs = descendants(scope, 'p').filter(isWordElement);
  const paragraphKey = `${role}Paragraphs` as const;
  limits[paragraphKey] += paragraphs.length;
  if (limits[paragraphKey] > MAX_PARAGRAPHS) {
    throw new Error(
      `${roleLabel(role)} DOCX exceeds the stable note/comment hyperlink paragraph limit.`,
    );
  }
  const result: NoteCommentParagraphRecord[] = [];
  for (const paragraph of paragraphs) {
    const runKey = `${role}Runs` as const;
    limits[runKey] += directContentRunCount(paragraph);
    if (limits[runKey] > MAX_RUNS) {
      throw new Error(
        `${roleLabel(role)} DOCX exceeds the stable note/comment hyperlink run limit.`,
      );
    }
    const content = noteCommentParagraphContent(paragraph, relationships, role);
    if (!content?.text) continue;
    const ancestry = paragraphAncestry(paragraph, scope);
    if (!ancestry) continue;
    result.push({
      element: paragraph,
      hyperlinks: content.hyperlinks,
      identity: `${ancestry}\u0000${content.text}`,
      runs: content.runs,
      text: content.text,
    });
  }
  return result;
}

function directContentRunCount(paragraph: Element): number {
  let count = 0;
  for (const child of Array.from(paragraph.children)) {
    if (!isWordElement(child)) continue;
    if (child.localName === 'r') count += 1;
    if (child.localName === 'hyperlink') {
      count += wordDirectChildren(child, 'r').length;
    }
  }
  return count;
}

export function noteCommentParagraphContent(
  paragraph: Element,
  relationships: DocxHyperlinkRelationshipState,
  role: DocumentRole,
): ParagraphContent | null {
  const runs: NoteCommentRunRecord[] = [];
  const hyperlinks: NoteCommentHyperlinkRecord[] = [];
  let offset = 0;
  let paragraphProperties = 0;
  for (const child of Array.from(paragraph.children)) {
    if (!isWordElement(child)) {
      if (isContentSemanticNamespace(child.namespaceURI ?? '')) return null;
      continue;
    }
    if (child.localName === 'pPr') {
      paragraphProperties += 1;
      if (paragraphProperties > 1) return null;
      continue;
    }
    if (child.localName === 'r') {
      const text = contentRunText(child);
      if (text === null) return null;
      const start = offset;
      offset += text.length;
      if (text) {
        runs.push({
          container: null,
          element: child,
          start,
          end: offset,
          text,
        });
      }
      continue;
    }
    if (child.localName !== 'hyperlink') return null;
    const start = offset;
    for (const hyperlinkChild of Array.from(child.children)) {
      if (!isWordElement(hyperlinkChild)) {
        if (isContentSemanticNamespace(hyperlinkChild.namespaceURI ?? '')) {
          return null;
        }
        continue;
      }
      if (hyperlinkChild.localName !== 'r') return null;
      const text = contentRunText(hyperlinkChild);
      if (text === null) return null;
      const runStart = offset;
      offset += text.length;
      if (text) {
        runs.push({
          container: child,
          element: hyperlinkChild,
          start: runStart,
          end: offset,
          text,
        });
      }
    }
    if (offset > start) {
      hyperlinks.push({
        destination: readDocxHyperlinkDestination(child, relationships, role),
        element: child,
        start,
        end: offset,
        text: runs
          .filter((run) => run.container === child)
          .map((run) => run.text)
          .join(''),
      });
    }
  }
  return { hyperlinks, runs, text: runs.map((run) => run.text).join('') };
}

export function groupNoteCommentParagraphs(
  paragraphs: readonly NoteCommentParagraphRecord[],
): Map<string, NoteCommentParagraphRecord[]> {
  const result = new Map<string, NoteCommentParagraphRecord[]>();
  for (const paragraph of paragraphs) {
    const matches = result.get(paragraph.identity) ?? [];
    matches.push(paragraph);
    result.set(paragraph.identity, matches);
  }
  return result;
}

export function groupNoteCommentHyperlinks(
  hyperlinks: readonly NoteCommentHyperlinkRecord[],
): Map<string, NoteCommentHyperlinkRecord[]> {
  const result = new Map<string, NoteCommentHyperlinkRecord[]>();
  for (const hyperlink of hyperlinks) {
    const key = noteCommentHyperlinkSpanKey(hyperlink);
    const matches = result.get(key) ?? [];
    matches.push(hyperlink);
    result.set(key, matches);
  }
  return result;
}

export function noteCommentHyperlinkSpanKey(
  item: NoteCommentHyperlinkRecord,
): string {
  return `${item.start}:${item.end}:${item.text}`;
}

export function noteCommentHyperlinksOverlap(
  left: NoteCommentHyperlinkRecord,
  right: NoteCommentHyperlinkRecord,
): boolean {
  return left.start < right.end && right.start < left.end;
}

export function restoreElementChildren(
  element: Element,
  children: readonly Node[],
): void {
  while (element.firstChild) element.removeChild(element.firstChild);
  for (const child of children) element.append(child);
}

export function alignCommentRunBoundaries(
  generated: NoteCommentParagraphRecord,
  source: NoteCommentParagraphRecord,
): boolean {
  if (
    generated.hyperlinks.length ||
    generated.runs.length !== 1 ||
    generated.runs[0].container ||
    source.runs.length <= 1 ||
    source.runs.length > MAX_COMMENT_RUN_SEGMENTS ||
    !isSimpleTextRun(generated.runs[0].element) ||
    source.runs.some((run) => !isSimpleTextRun(run.element))
  ) {
    return false;
  }
  const original = generated.runs[0].element;
  if (!original.parentNode) return false;
  const clones: Element[] = [];
  for (const sourceRun of source.runs) {
    const clone = original.cloneNode(true) as Element;
    const textElement = wordDirectChildren(clone, 't')[0];
    if (!textElement) return false;
    textElement.textContent = sourceRun.text;
    setTextSpacePreservation(textElement, sourceRun.text);
    clones.push(clone);
  }
  for (const clone of clones) original.parentNode.insertBefore(clone, original);
  original.remove();
  return true;
}

export function isKnownOoxmlNamespace(namespace: string): boolean {
  if (namespace === MARKUP_COMPATIBILITY_NAMESPACE) return false;
  return (
    DOCX_WORDPROCESSING_NAMESPACES.has(namespace) ||
    namespace.startsWith('http://schemas.microsoft.com/office/') ||
    namespace.startsWith('http://schemas.openxmlformats.org/') ||
    namespace.startsWith('http://purl.oclc.org/ooxml/') ||
    namespace.startsWith('urn:schemas-microsoft-com:') ||
    namespace.startsWith('urn:microsoft-com:office:')
  );
}

function contentRunText(run: Element): string | null {
  let text = '';
  for (const child of Array.from(run.children)) {
    if (!isWordElement(child)) {
      if (isContentSemanticNamespace(child.namespaceURI ?? '')) return null;
      continue;
    }
    if (child.localName === 'rPr') continue;
    if (child.localName === 't' || child.localName === 'delText') {
      text += child.textContent ?? '';
    } else if (child.localName === 'tab') text += '\t';
    else if (child.localName === 'br' || child.localName === 'cr') text += '\n';
    else if (child.localName === 'noBreakHyphen') text += '\u2011';
    else if (child.localName === 'softHyphen') text += '\u00ad';
    else if (
      child.localName !== 'footnoteRef' &&
      child.localName !== 'endnoteRef'
    ) {
      return null;
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

function paragraphAncestry(paragraph: Element, scope: Element): string | null {
  const path: string[] = [];
  let current: Element | null = paragraph;
  while (current && current !== scope) {
    if (!isWordElement(current)) return null;
    path.push(current.localName);
    current = current.parentElement;
  }
  return current === scope ? path.reverse().join('/') : null;
}

function isContentSemanticNamespace(namespace: string): boolean {
  return (
    namespace === MARKUP_COMPATIBILITY_NAMESPACE ||
    isKnownOoxmlNamespace(namespace)
  );
}

function wordDirectChildren(element: Element, localName: string): Element[] {
  return Array.from(element.children).filter(
    (child) => child.localName === localName && isWordElement(child),
  );
}

function isWordElement(element: Element): boolean {
  return DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '');
}

function roleLabel(role: DocumentRole): string {
  return role === 'source' ? 'Registered source' : 'Generated';
}
