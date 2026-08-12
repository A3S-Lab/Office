import {
  DOCX_WORDPROCESSING_NAMESPACES,
  mergeDocxIgnorableExtensionsAtPairs,
  type DocxExtensionDocumentRole,
  type DocxIgnorableExtensionPair,
} from './work-docx-ignorable-extension-preservation';
import { preserveDocxNoteCommentRunProperties } from './work-docx-note-comment-run-properties';
import { descendants } from './work-ooxml-package';
import {
  XML_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';

const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const MAX_CONTENT_PARAGRAPHS = 65_536;
const MAX_CONTENT_RUNS = 262_144;
const MAX_COMMENT_RUN_SEGMENTS = 4_096;

type NoteCommentContentKind = 'comment' | 'note';
type ContentDocumentRole = 'generated' | 'source';

interface ContentLimits {
  generatedParagraphs: number;
  generatedRuns: number;
  sourceParagraphs: number;
  sourceRuns: number;
}

interface ContentParagraph {
  element: Element;
  identity: string;
}

interface ContentRun {
  element: Element;
  end: number;
  start: number;
  text: string;
}

export function preserveDocxNoteCommentRunContent(
  generatedDocument: Document,
  sourceDocument: Document,
  scopes: readonly DocxIgnorableExtensionPair[],
  kind: NoteCommentContentKind,
): void {
  const paragraphPairs: DocxIgnorableExtensionPair[] = [];
  const runPairs: DocxIgnorableExtensionPair[] = [];
  const limits: ContentLimits = {
    generatedParagraphs: 0,
    generatedRuns: 0,
    sourceParagraphs: 0,
    sourceRuns: 0,
  };

  for (const scope of scopes) {
    const pairs = uniqueParagraphPairs(scope.generated, scope.source, limits);
    for (const pair of pairs) {
      if (kind === 'comment') alignCommentRunBoundaries(pair);
      const generatedRuns = contentRuns(pair.generated);
      const sourceRuns = contentRuns(pair.source);
      if (!generatedRuns || !sourceRuns) continue;
      const pairsBySpan = uniqueRunPairs(generatedRuns, sourceRuns);
      for (const runPair of pairsBySpan) {
        preserveDocxNoteCommentRunProperties(
          generatedDocument,
          runPair.generated,
          runPair.source,
          kind,
        );
        runPairs.push(runPair);
      }
      paragraphPairs.push(pair);
    }
  }

  mergeContentExtensions(
    generatedDocument,
    sourceDocument,
    paragraphPairs,
    (_generated, _source, depth) => depth === 0,
  );
  mergeContentExtensions(
    generatedDocument,
    sourceDocument,
    runPairs,
    (_generated, _source, depth) => depth <= 4,
  );
}

function uniqueParagraphPairs(
  generatedScope: Element,
  sourceScope: Element,
  limits: ContentLimits,
): DocxIgnorableExtensionPair[] {
  const generated = groupParagraphs(
    contentParagraphs(generatedScope, 'generated', limits),
  );
  const source = groupParagraphs(
    contentParagraphs(sourceScope, 'source', limits),
  );
  const pairs: DocxIgnorableExtensionPair[] = [];
  for (const [identity, sourceParagraphs] of source) {
    const generatedParagraphs = generated.get(identity) ?? [];
    if (sourceParagraphs.length !== 1 || generatedParagraphs.length !== 1) {
      continue;
    }
    pairs.push({
      generated: generatedParagraphs[0].element,
      source: sourceParagraphs[0].element,
    });
  }
  return pairs;
}

function contentParagraphs(
  scope: Element,
  role: ContentDocumentRole,
  limits: ContentLimits,
): ContentParagraph[] {
  const records: ContentParagraph[] = [];
  const paragraphs = descendants(scope, 'p').filter(isWordElement);
  const paragraphKey = `${role}Paragraphs` as const;
  const runKey = `${role}Runs` as const;
  limits[paragraphKey] += paragraphs.length;
  if (limits[paragraphKey] > MAX_CONTENT_PARAGRAPHS) {
    throw new Error(
      `${role === 'source' ? 'Registered source' : 'Generated'} DOCX exceeds the stable note/comment paragraph limit.`,
    );
  }
  for (const paragraph of paragraphs) {
    limits[runKey] += wordDirectChildren(paragraph, 'r').length;
    if (limits[runKey] > MAX_CONTENT_RUNS) {
      throw new Error(
        `${role === 'source' ? 'Registered source' : 'Generated'} DOCX exceeds the stable note/comment run limit.`,
      );
    }
    const runs = contentRuns(paragraph);
    if (!runs?.length) continue;
    const text = runs.map((run) => run.text).join('');
    if (!text) continue;
    const ancestry = paragraphAncestry(paragraph, scope);
    if (!ancestry) continue;
    records.push({
      element: paragraph,
      identity: `${ancestry}\u0000${text}`,
    });
  }
  return records;
}

function groupParagraphs(
  paragraphs: readonly ContentParagraph[],
): Map<string, ContentParagraph[]> {
  const result = new Map<string, ContentParagraph[]>();
  for (const paragraph of paragraphs) {
    const matches = result.get(paragraph.identity) ?? [];
    matches.push(paragraph);
    result.set(paragraph.identity, matches);
  }
  return result;
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

function contentRuns(paragraph: Element): ContentRun[] | null {
  if (!eligibleParagraphChildren(paragraph)) return null;
  const records: ContentRun[] = [];
  let offset = 0;
  for (const run of wordDirectChildren(paragraph, 'r')) {
    const text = contentRunText(run);
    if (text === null) return null;
    const start = offset;
    offset += text.length;
    if (text) records.push({ element: run, start, end: offset, text });
  }
  return records;
}

function eligibleParagraphChildren(paragraph: Element): boolean {
  return Array.from(paragraph.children).every((child) => {
    if (isWordElement(child)) {
      return child.localName === 'pPr' || child.localName === 'r';
    }
    return !isContentSemanticNamespace(child.namespaceURI ?? '');
  });
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
      continue;
    }
    if (child.localName === 'tab') {
      text += '\t';
      continue;
    }
    if (child.localName === 'br' || child.localName === 'cr') {
      text += '\n';
      continue;
    }
    if (child.localName === 'noBreakHyphen') {
      text += '\u2011';
      continue;
    }
    if (child.localName === 'softHyphen') {
      text += '\u00ad';
      continue;
    }
    if (child.localName === 'footnoteRef' || child.localName === 'endnoteRef') {
      continue;
    }
    return null;
  }
  return text;
}

function uniqueRunPairs(
  generated: readonly ContentRun[],
  source: readonly ContentRun[],
): DocxIgnorableExtensionPair[] {
  const generatedBySpan = groupRunsBySpan(generated);
  const sourceBySpan = groupRunsBySpan(source);
  const result: DocxIgnorableExtensionPair[] = [];
  for (const [span, sourceRuns] of sourceBySpan) {
    const generatedRuns = generatedBySpan.get(span) ?? [];
    if (sourceRuns.length === 1 && generatedRuns.length === 1) {
      result.push({
        generated: generatedRuns[0].element,
        source: sourceRuns[0].element,
      });
    }
  }
  return result;
}

function groupRunsBySpan(
  runs: readonly ContentRun[],
): Map<string, ContentRun[]> {
  const result = new Map<string, ContentRun[]>();
  for (const run of runs) {
    const key = `${run.start}:${run.end}:${run.text}`;
    const matches = result.get(key) ?? [];
    matches.push(run);
    result.set(key, matches);
  }
  return result;
}

function alignCommentRunBoundaries(pair: DocxIgnorableExtensionPair): void {
  const generatedRuns = contentRuns(pair.generated);
  const sourceRuns = contentRuns(pair.source);
  if (
    !generatedRuns ||
    !sourceRuns ||
    generatedRuns.length !== 1 ||
    sourceRuns.length < 2 ||
    sourceRuns.length > MAX_COMMENT_RUN_SEGMENTS ||
    generatedRuns[0].text !== sourceRuns.map((run) => run.text).join('') ||
    !isSimpleTextRun(generatedRuns[0].element) ||
    sourceRuns.some((run) => !isSimpleTextRun(run.element))
  ) {
    return;
  }
  const original = generatedRuns[0].element;
  const textElement = wordDirectChildren(original, 't')[0];
  if (!textElement || !original.parentNode) return;
  for (const sourceRun of sourceRuns) {
    const clone = original.cloneNode(true) as Element;
    const cloneText = wordDirectChildren(clone, 't')[0];
    if (!cloneText) return;
    cloneText.textContent = sourceRun.text;
    setTextSpacePreservation(cloneText, sourceRun.text);
    original.parentNode.insertBefore(clone, original);
  }
  original.remove();
}

function isSimpleTextRun(run: Element): boolean {
  const wordChildren = Array.from(run.children).filter(isWordElement);
  return (
    wordChildren.filter((child) => child.localName === 't').length === 1 &&
    wordChildren.every(
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
  if (/^\s|\s$/.test(text)) {
    element.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  }
}

function mergeContentExtensions(
  generatedDocument: Document,
  sourceDocument: Document,
  pairs: readonly DocxIgnorableExtensionPair[],
  allowMatchedElementMerge: NonNullable<
    Parameters<typeof mergeDocxIgnorableExtensionsAtPairs>[3]
  >['allowMatchedElementMerge'],
): void {
  if (!pairs.length) return;
  mergeDocxIgnorableExtensionsAtPairs(
    generatedDocument,
    sourceDocument,
    pairs,
    {
      semanticKey: contentSemanticKey,
      isAdditionalSemanticNamespace: isKnownOoxmlNamespace,
      allowExtensionNamespace: (namespace) => !isKnownOoxmlNamespace(namespace),
      allowMatchedElementMerge,
    },
  );
}

function contentSemanticKey(
  element: Element,
  _role: DocxExtensionDocumentRole,
): string {
  return isKnownOoxmlNamespace(element.namespaceURI ?? '')
    ? `{semantic}${element.localName}`
    : `{${element.namespaceURI ?? ''}}${element.localName}`;
}

function isKnownOoxmlNamespace(namespace: string): boolean {
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
