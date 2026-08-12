import type JSZip from 'jszip';
import {
  DOCX_WORDPROCESSING_NAMESPACES,
  mergeDocxIgnorableExtensionsAtPairs,
  type DocxExtensionDocumentRole,
  type DocxIgnorableExtensionPair,
} from './work-docx-ignorable-extension-preservation';
import {
  alignCommentRunBoundaries,
  groupNoteCommentHyperlinks,
  groupNoteCommentParagraphs,
  isKnownOoxmlNamespace,
  noteCommentContentParagraphs,
  noteCommentHyperlinkSpanKey,
  noteCommentHyperlinksOverlap,
  noteCommentParagraphContent,
  restoreElementChildren,
  type NoteCommentHyperlinkContentLimits,
  type NoteCommentHyperlinkRecord,
  type NoteCommentParagraphRecord,
  type NoteCommentRunRecord,
} from './work-docx-note-comment-hyperlink-content';
import {
  ensureDocxExternalHyperlinkRelationships,
  flushDocxHyperlinkRelationships,
  loadDocxHyperlinkRelationshipState,
  setDocxHyperlinkDestination,
  type DocxHyperlinkDestination,
  type DocxHyperlinkRelationshipState,
} from './work-docx-note-comment-hyperlink-relationships';
import { preserveDocxNoteCommentRunProperties } from './work-docx-note-comment-run-properties';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
  xmlDeclaredPrefix,
  xmlNamespaceUri,
} from './work-docx-settings-xml';

type NoteCommentKind = 'comment' | 'note';

interface HyperlinkPreservationPlan {
  existing?: NoteCommentHyperlinkRecord;
  relationshipId?: string;
  runs?: NoteCommentRunRecord[];
  source: NoteCommentHyperlinkRecord;
}

export async function preserveDocxNoteCommentHyperlinks(
  generatedArchive: JSZip,
  sourceArchive: JSZip,
  generatedDocument: Document,
  sourceDocument: Document,
  scopes: readonly DocxIgnorableExtensionPair[],
  kind: NoteCommentKind,
  ownerPart: string,
): Promise<void> {
  if (!scopes.length) return;
  const relationships = await loadDocxHyperlinkRelationshipState(
    generatedArchive,
    sourceArchive,
    ownerPart,
  );
  const limits: NoteCommentHyperlinkContentLimits = {
    generatedParagraphs: 0,
    generatedRuns: 0,
    sourceParagraphs: 0,
    sourceRuns: 0,
  };
  const paragraphPairs: DocxIgnorableExtensionPair[] = [];
  const hyperlinkPairs: DocxIgnorableExtensionPair[] = [];
  const runPairs: DocxIgnorableExtensionPair[] = [];
  for (const scope of scopes) {
    const matchedParagraphs = uniqueParagraphPairs(
      scope,
      relationships,
      limits,
    );
    for (const pair of matchedParagraphs) {
      preserveParagraphHyperlinks(
        generatedDocument,
        pair.generated,
        pair.source,
        relationships,
        kind,
        paragraphPairs,
        hyperlinkPairs,
        runPairs,
      );
    }
  }
  mergeTopLevelExtensions(generatedDocument, sourceDocument, paragraphPairs);
  mergeTopLevelExtensions(generatedDocument, sourceDocument, hyperlinkPairs);
  for (const pair of runPairs) {
    preserveDocxNoteCommentRunProperties(
      generatedDocument,
      pair.generated,
      pair.source,
      kind,
    );
  }
  mergeRunExtensions(generatedDocument, sourceDocument, runPairs);
  flushDocxHyperlinkRelationships(relationships);
}

function uniqueParagraphPairs(
  scope: DocxIgnorableExtensionPair,
  relationships: DocxHyperlinkRelationshipState,
  limits: NoteCommentHyperlinkContentLimits,
): Array<{
  generated: NoteCommentParagraphRecord;
  source: NoteCommentParagraphRecord;
}> {
  const generated = groupNoteCommentParagraphs(
    noteCommentContentParagraphs(
      scope.generated,
      'generated',
      relationships,
      limits,
    ),
  );
  const source = groupNoteCommentParagraphs(
    noteCommentContentParagraphs(scope.source, 'source', relationships, limits),
  );
  const pairs: Array<{
    generated: NoteCommentParagraphRecord;
    source: NoteCommentParagraphRecord;
  }> = [];
  for (const [identity, sourceItems] of source) {
    const generatedItems = generated.get(identity) ?? [];
    if (sourceItems.length === 1 && generatedItems.length === 1) {
      pairs.push({ generated: generatedItems[0], source: sourceItems[0] });
    }
  }
  return pairs;
}

function preserveParagraphHyperlinks(
  generatedDocument: Document,
  originalGenerated: NoteCommentParagraphRecord,
  source: NoteCommentParagraphRecord,
  relationships: DocxHyperlinkRelationshipState,
  kind: NoteCommentKind,
  paragraphPairs: DocxIgnorableExtensionPair[],
  pairs: DocxIgnorableExtensionPair[],
  runPairs: DocxIgnorableExtensionPair[],
): void {
  if (!eligibleSourceHyperlinks(source.hyperlinks)) return;
  const originalChildren =
    kind === 'comment'
      ? Array.from(originalGenerated.element.childNodes).map((child) =>
          child.cloneNode(true),
        )
      : [];
  const aligned =
    kind === 'comment' && alignCommentRunBoundaries(originalGenerated, source);
  const generatedContent = noteCommentParagraphContent(
    originalGenerated.element,
    relationships,
    'generated',
  );
  const plans =
    generatedContent?.text === source.text
      ? planParagraphHyperlinks(
          generatedContent.hyperlinks,
          generatedContent.runs,
          source.hyperlinks,
          relationships,
          kind,
        )
      : null;
  if (!plans) {
    if (aligned)
      restoreElementChildren(originalGenerated.element, originalChildren);
    return;
  }
  for (const plan of plans) {
    const generatedHyperlink =
      plan.existing?.element ??
      wrapRuns(generatedDocument, originalGenerated.element, plan.runs ?? []);
    if (
      !generatedHyperlink ||
      !plan.source.destination ||
      (!plan.existing &&
        !setDocxHyperlinkDestination(
          generatedHyperlink,
          generatedDocument.documentElement,
          plan.source.destination,
          plan.relationshipId,
        ))
    ) {
      throw new Error('A planned note/comment hyperlink could not be emitted.');
    }
    preserveHyperlinkMetadata(
      generatedDocument,
      generatedHyperlink,
      plan.source.element,
    );
    pairs.push({ generated: generatedHyperlink, source: plan.source.element });
  }
  const generated = noteCommentParagraphContent(
    originalGenerated.element,
    relationships,
    'generated',
  );
  if (
    !generated ||
    !haveStableHyperlinks(generated.hyperlinks, source.hyperlinks)
  ) {
    throw new Error('A planned note/comment hyperlink mapping was not stable.');
  }
  paragraphPairs.push({
    generated: originalGenerated.element,
    source: source.element,
  });
  runPairs.push(...uniqueRunPairs(generated.runs, source.runs));
}

function planParagraphHyperlinks(
  generated: readonly NoteCommentHyperlinkRecord[],
  generatedRuns: readonly NoteCommentRunRecord[],
  source: readonly NoteCommentHyperlinkRecord[],
  relationships: DocxHyperlinkRelationshipState,
  kind: NoteCommentKind,
): HyperlinkPreservationPlan[] | null {
  const plans: HyperlinkPreservationPlan[] = [];
  const matchedGenerated = new Set<Element>();
  for (const sourceHyperlink of source) {
    if (!sourceHyperlink.destination) return null;
    const sameSpan = generated.filter(
      (item) =>
        noteCommentHyperlinkSpanKey(item) ===
        noteCommentHyperlinkSpanKey(sourceHyperlink),
    );
    if (sameSpan.length) {
      if (
        sameSpan.length !== 1 ||
        !sameDestination(sameSpan[0].destination, sourceHyperlink.destination)
      ) {
        return null;
      }
      matchedGenerated.add(sameSpan[0].element);
      plans.push({ existing: sameSpan[0], source: sourceHyperlink });
      continue;
    }
    if (
      kind !== 'comment' ||
      generated.some((item) =>
        noteCommentHyperlinksOverlap(item, sourceHyperlink),
      )
    ) {
      return null;
    }
    const runs = runsCovering(generatedRuns, sourceHyperlink);
    if (!runs) return null;
    plans.push({ runs, source: sourceHyperlink });
  }
  if (generated.some((item) => !matchedGenerated.has(item.element)))
    return null;
  const externalPlans = plans.filter(
    (plan) => !plan.existing && plan.source.destination?.kind === 'external',
  );
  const relationshipIds = ensureDocxExternalHyperlinkRelationships(
    relationships,
    externalPlans.map((plan) => ({
      sourceId: plan.source.destination?.relationshipId ?? '',
      target: plan.source.destination?.target ?? '',
    })),
  );
  if (!relationshipIds) return null;
  for (const [index, plan] of externalPlans.entries()) {
    plan.relationshipId = relationshipIds[index];
  }
  return plans;
}

function eligibleSourceHyperlinks(
  hyperlinks: readonly NoteCommentHyperlinkRecord[],
): boolean {
  if (!hyperlinks.length || hyperlinks.some((item) => !item.destination)) {
    return false;
  }
  return Array.from(groupNoteCommentHyperlinks(hyperlinks).values()).every(
    (items) => items.length === 1,
  );
}

function runsCovering(
  runs: readonly NoteCommentRunRecord[],
  hyperlink: NoteCommentHyperlinkRecord,
): NoteCommentRunRecord[] | null {
  const selected = runs.filter(
    (run) => run.start >= hyperlink.start && run.end <= hyperlink.end,
  );
  if (
    !selected.length ||
    selected[0].start !== hyperlink.start ||
    selected.at(-1)?.end !== hyperlink.end ||
    selected.some((run) => run.container || !run.element.parentElement)
  ) {
    return null;
  }
  const parent = selected[0].element.parentElement;
  if (!parent || selected.some((run) => run.element.parentElement !== parent)) {
    return null;
  }
  const semanticChildren = Array.from(parent.children).filter((child) =>
    DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? ''),
  );
  const firstIndex = semanticChildren.indexOf(selected[0].element);
  return firstIndex >= 0 &&
    selected.every(
      (run, index) => semanticChildren[firstIndex + index] === run.element,
    )
    ? selected
    : null;
}

function wrapRuns(
  document: Document,
  paragraph: Element,
  runs: readonly NoteCommentRunRecord[],
): Element | null {
  const first = runs[0]?.element;
  const namespace = paragraph.namespaceURI;
  if (!first?.parentNode || !namespace) return null;
  const prefix = ensureNamespacePrefix(
    document.documentElement,
    namespace,
    paragraph.prefix ?? 'w',
  );
  const hyperlink = document.createElementNS(namespace, `${prefix}:hyperlink`);
  first.parentNode.insertBefore(hyperlink, first);
  for (const run of runs) hyperlink.append(run.element);
  return hyperlink;
}

function preserveHyperlinkMetadata(
  document: Document,
  generated: Element,
  source: Element,
): void {
  const tooltip = uniqueWordAttribute(source, 'tooltip');
  if (
    tooltip &&
    tooltip.length <= 260 &&
    !/[\u0000-\u001f\u007f]/u.test(tooltip) &&
    uniqueWordAttribute(generated, 'tooltip') === null
  ) {
    const namespace = generated.namespaceURI;
    if (!namespace) return;
    const prefix = ensureNamespacePrefix(
      document.documentElement,
      namespace,
      generated.prefix ?? 'w',
    );
    generated.setAttributeNS(namespace, `${prefix}:tooltip`, tooltip);
  }
}

function uniqueWordAttribute(
  element: Element,
  localName: string,
): string | null | undefined {
  const matches = Array.from(element.attributes).filter(
    (item) =>
      xmlAttributeLocalName(item) === localName &&
      DOCX_WORDPROCESSING_NAMESPACES.has(
        xmlAttributeNamespace(element, item) ?? '',
      ),
  );
  return matches.length <= 1 ? (matches[0]?.value ?? null) : undefined;
}

function mergeTopLevelExtensions(
  generated: Document,
  source: Document,
  pairs: readonly DocxIgnorableExtensionPair[],
): void {
  if (!pairs.length) return;
  mergeDocxIgnorableExtensionsAtPairs(generated, source, pairs, {
    semanticKey: contentSemanticKey,
    isAdditionalSemanticNamespace: isKnownOoxmlNamespace,
    allowExtensionNamespace: (namespace) => !isKnownOoxmlNamespace(namespace),
    allowMatchedElementMerge: (_generated, _source, depth) => depth === 0,
  });
}

function mergeRunExtensions(
  generated: Document,
  source: Document,
  pairs: readonly DocxIgnorableExtensionPair[],
): void {
  if (!pairs.length) return;
  mergeDocxIgnorableExtensionsAtPairs(generated, source, pairs, {
    semanticKey: contentSemanticKey,
    isAdditionalSemanticNamespace: isKnownOoxmlNamespace,
    allowExtensionNamespace: (namespace) => !isKnownOoxmlNamespace(namespace),
    allowMatchedElementMerge: (_generated, _source, depth) => depth <= 4,
  });
}

function uniqueRunPairs(
  generated: readonly NoteCommentRunRecord[],
  source: readonly NoteCommentRunRecord[],
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
  runs: readonly NoteCommentRunRecord[],
): Map<string, NoteCommentRunRecord[]> {
  const result = new Map<string, NoteCommentRunRecord[]>();
  for (const run of runs) {
    const key = `${run.start}:${run.end}:${run.text}`;
    const matches = result.get(key) ?? [];
    matches.push(run);
    result.set(key, matches);
  }
  return result;
}

function haveStableHyperlinks(
  generated: readonly NoteCommentHyperlinkRecord[],
  source: readonly NoteCommentHyperlinkRecord[],
): boolean {
  if (!source.length || generated.length !== source.length) return false;
  const generatedBySpan = groupNoteCommentHyperlinks(generated);
  const sourceBySpan = groupNoteCommentHyperlinks(source);
  for (const [span, sourceItems] of sourceBySpan) {
    const generatedItems = generatedBySpan.get(span) ?? [];
    if (
      sourceItems.length !== 1 ||
      generatedItems.length !== 1 ||
      !sourceItems[0].destination ||
      !sameDestination(
        generatedItems[0].destination,
        sourceItems[0].destination,
      )
    ) {
      return false;
    }
  }
  return generatedBySpan.size === sourceBySpan.size;
}

function sameDestination(
  left: DocxHyperlinkDestination | null,
  right: DocxHyperlinkDestination,
): boolean {
  return Boolean(
    left && left.kind === right.kind && left.target === right.target,
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

function ensureNamespacePrefix(
  root: Element,
  namespace: string,
  preferred: string,
): string {
  const existing = xmlDeclaredPrefix(root, namespace);
  if (existing) return existing;
  let prefix = preferred;
  let index = 1;
  while (xmlNamespaceUri(root, prefix)) {
    prefix = `${preferred}${index}`;
    index += 1;
  }
  root.setAttributeNS(XMLNS_NAMESPACE, `xmlns:${prefix}`, namespace);
  return prefix;
}
