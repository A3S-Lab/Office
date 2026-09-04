import { type Content, createNodeFromContent, type Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import {
  Fragment,
  type Mark as ProseMirrorMark,
  Node as ProseMirrorNode,
} from '@tiptap/pm/model';
import {
  collectDocumentChanges,
  resolveAllDocumentChanges,
  type WorkDocumentChangeIdentity,
} from './work-document-changes';
import {
  boundedDocumentSequenceAlignment,
  boundedDocumentSequenceDiff,
} from './work-document-compare-diff';
import {
  type ComparisonIdentityFactory,
  createComparisonIdentityFactory,
  type DocumentComparisonOptions,
  type DocumentComparisonSummary,
  emptyComparisonSummary,
  summarizeComparisonChanges,
} from './work-document-compare-identities';
import {
  appendComparisonRevisionUnits,
  appendRevisionUnits,
  type InlineDiffStep,
  type InlineMoveAssignment,
  type InlineMovePair,
  type InlineUnit,
  inferInlineMovePairs,
  inferInlineMovePairsAcrossScopes,
  MAX_INFERRED_MOVE_TEXT,
} from './work-document-compare-moves';
import { boundedMetadata, stableJson } from './work-document-compare-stability';
import {
  isDocumentCharacterFormatMark,
  serializeDocumentCharacterFormatting,
} from './work-document-format-changes';
import {
  DOCUMENT_PARAGRAPH_CHANGE_ATTRIBUTES,
  DOCUMENT_PARAGRAPH_FORMAT_ATTRIBUTES,
  serializeDocumentParagraphFormatting,
} from './work-document-paragraph-format-changes';
import type { WorkDocumentChangeKind } from './work-types';

export type {
  DocumentComparisonMode,
  DocumentComparisonOptions,
  DocumentComparisonSummary,
} from './work-document-compare-identities';

export type DocumentComparisonDiagnosticCode =
  | 'changed-complex-structure'
  | 'combine-baseline-mismatch'
  | 'combine-resolution-invalid'
  | 'combine-structural-revisions'
  | 'combine-without-revisions'
  | 'comparison-limit-exceeded'
  | 'current-revisions-present'
  | 'empty-structural-change'
  | 'invalid-revised-content'
  | 'revised-revisions-present'
  | 'section-layout-mismatch'
  | 'unsupported-inline-review-state';

export interface DocumentComparisonDiagnostic {
  code: DocumentComparisonDiagnosticCode;
  message: string;
  section?: number;
  block?: number;
}

export interface DocumentComparisonApplyResult {
  status: 'applied' | 'unchanged' | 'unsupported';
  summary: DocumentComparisonSummary;
  diagnostics: DocumentComparisonDiagnostic[];
}

interface DocumentComparisonPlan extends DocumentComparisonApplyResult {
  document?: ProseMirrorNode;
}

const MAX_COMPARISON_BLOCKS = 1_024;
const MAX_COMPARISON_TEXT = 1_000_000;
const MAX_BLOCK_ALIGNMENT_CELLS = 1_100_000;
const MAX_INLINE_DIFF_CELLS = 1_100_000;
const BLOCK_CHANGE_ATTRIBUTE_NAMES = [
  'blockChangeKind',
  'blockChangeId',
  'blockChangeActorId',
  'blockChangeAuthor',
  'blockChangeDate',
] as const;
const PARAGRAPH_IDENTITY_ATTRIBUTE_NAMES = ['paragraphId', 'textId'] as const;
const REVIEW_ONLY_MARK_NAMES = new Set(['documentChange', 'documentComment']);

interface InlineComparisonVariant {
  currentUnits: InlineUnit[];
  revisedUnits: InlineUnit[];
  changes: InlineDiffStep[];
}

interface PreparedInlineComparison {
  variant: InlineComparisonVariant;
  attached?: InlineComparisonVariant;
}

interface PreparedPairedBlock {
  current: ProseMirrorNode;
  revised: ProseMirrorNode;
  scope: string;
  inline: PreparedInlineComparison;
}

type ComparisonBlockSlot =
  | { kind: 'node'; node: ProseMirrorNode }
  | { kind: 'paired'; comparison: PreparedPairedBlock };

export function applyDocumentComparison(
  editor: Editor,
  revisedContent: Content,
  options: DocumentComparisonOptions,
): DocumentComparisonApplyResult {
  const revisedDocument = parseComparisonDocument(editor, revisedContent);
  if (!revisedDocument) {
    return unsupportedResult({
      code: 'invalid-revised-content',
      message: 'The selected file cannot be represented by the Writer schema.',
    });
  }
  const plan =
    options.mode === 'combine'
      ? planDocumentCombine(editor.state.doc, revisedDocument)
      : planDocumentCompare(editor.state.doc, revisedDocument, options);
  if (plan.status !== 'applied' || !plan.document) return plan;

  const transaction = editor.state.tr;
  closeHistory(transaction);
  transaction
    .replaceWith(0, editor.state.doc.content.size, plan.document.content)
    .setMeta('documentComparison', {
      mode: options.mode,
      sourceName: boundedMetadata(options.sourceName, 512),
    });
  if (!transaction.docChanged) {
    return {
      status: 'unchanged',
      summary: plan.summary,
      diagnostics: plan.diagnostics,
    };
  }
  editor.view.dispatch(transaction);
  return {
    status: 'applied',
    summary: plan.summary,
    diagnostics: plan.diagnostics,
  };
}

function planDocumentCompare(
  current: ProseMirrorNode,
  revised: ProseMirrorNode,
  options: DocumentComparisonOptions,
): DocumentComparisonPlan {
  if (collectDocumentChanges(current).length) {
    return unsupportedResult({
      code: 'current-revisions-present',
      message: 'Resolve current revisions before comparing another version.',
    });
  }
  if (collectDocumentChanges(revised).length) {
    return unsupportedResult({
      code: 'revised-revisions-present',
      message:
        'Use Combine for a reviewed copy that already contains revisions.',
    });
  }
  const currentSections = documentSections(current);
  const revisedSections = documentSections(revised);
  if (
    !currentSections ||
    !revisedSections ||
    currentSections.length !== revisedSections.length
  ) {
    return unsupportedResult({
      code: 'section-layout-mismatch',
      message: 'Comparison requires the same bounded section structure.',
    });
  }
  const limits = comparisonLimits(currentSections, revisedSections);
  if (limits) return unsupportedResult(limits);

  const diagnostics: DocumentComparisonDiagnostic[] = [];
  const factory = createComparisonIdentityFactory(
    comparisonSemanticSignature(current),
    comparisonSemanticSignature(revised),
    options,
  );
  const comparedSections: ProseMirrorNode[] = [];
  for (
    let sectionIndex = 0;
    sectionIndex < currentSections.length;
    sectionIndex += 1
  ) {
    const currentSection = currentSections[sectionIndex] as ProseMirrorNode;
    const revisedSection = revisedSections[sectionIndex] as ProseMirrorNode;
    if (
      sectionLayoutSignature(currentSection) !==
      sectionLayoutSignature(revisedSection)
    ) {
      diagnostics.push({
        code: 'section-layout-mismatch',
        message:
          'Page and section layout differences are not approximated as text revisions.',
        section: sectionIndex,
      });
      continue;
    }
    const blocks = compareSectionBlocks(
      currentSection,
      revisedSection,
      sectionIndex,
      factory,
      diagnostics,
    );
    if (!blocks) continue;
    comparedSections.push(
      currentSection.type.create(
        currentSection.attrs,
        Fragment.fromArray(blocks),
      ),
    );
  }
  if (diagnostics.length) {
    return {
      status: 'unsupported',
      summary: factory.summary,
      diagnostics,
    };
  }
  const document = current.type.create(current.attrs, comparedSections);
  if (document.eq(current)) {
    return {
      status: 'unchanged',
      summary: emptyComparisonSummary(),
      diagnostics: [],
    };
  }
  return {
    status: 'applied',
    document,
    summary: factory.summary,
    diagnostics: [],
  };
}

function planDocumentCombine(
  current: ProseMirrorNode,
  reviewed: ProseMirrorNode,
): DocumentComparisonPlan {
  if (collectDocumentChanges(current).length) {
    return unsupportedResult({
      code: 'current-revisions-present',
      message:
        'Resolve current revisions before combining another reviewed copy.',
    });
  }
  const reviewedChanges = collectDocumentChanges(reviewed);
  if (!reviewedChanges.length) {
    return unsupportedResult({
      code: 'combine-without-revisions',
      message: 'The selected reviewed copy contains no revisions to combine.',
    });
  }
  if (hasStructuralBlockRevisions(reviewed)) {
    return unsupportedResult({
      code: 'combine-structural-revisions',
      message:
        'This reviewed copy changes the paragraph tree; combine supports inline and formatting revisions.',
    });
  }
  const baseline = resolveAllDocumentChanges(reviewed, 'reject');
  if (!baseline) {
    return unsupportedResult({
      code: 'combine-resolution-invalid',
      message: 'The reviewed copy contains a malformed formatting revision.',
    });
  }
  if (!simpleDocumentShape(current) || !simpleDocumentShape(baseline)) {
    return unsupportedResult({
      code: 'changed-complex-structure',
      message:
        'Combine currently accepts text-only paragraph and heading review copies.',
    });
  }
  if (
    comparisonSemanticSignature(current) !==
    comparisonSemanticSignature(baseline)
  ) {
    return unsupportedResult({
      code: 'combine-baseline-mismatch',
      message:
        'Rejecting the imported revisions does not reproduce the current document.',
    });
  }
  const document = transferCombineIdentities(current, reviewed);
  if (!document) {
    return unsupportedResult({
      code: 'combine-structural-revisions',
      message:
        'This reviewed copy changes the paragraph tree; combine supports inline and formatting revisions.',
    });
  }
  return {
    status: 'applied',
    document,
    summary: summarizeComparisonChanges(reviewedChanges),
    diagnostics: [],
  };
}

function compareSectionBlocks(
  currentSection: ProseMirrorNode,
  revisedSection: ProseMirrorNode,
  sectionIndex: number,
  factory: ComparisonIdentityFactory,
  diagnostics: DocumentComparisonDiagnostic[],
): ProseMirrorNode[] | null {
  const currentBlocks = childNodes(currentSection);
  const revisedBlocks = childNodes(revisedSection);
  const alignment = boundedDocumentSequenceAlignment(
    currentBlocks,
    revisedBlocks,
    blockSubstitutionCost,
    MAX_BLOCK_ALIGNMENT_CELLS,
  );
  if (!alignment) {
    diagnostics.push({
      code: 'comparison-limit-exceeded',
      message:
        'The block-alignment matrix exceeds the bounded comparison limit.',
      section: sectionIndex,
    });
    return null;
  }
  const slots: ComparisonBlockSlot[] = [];
  for (const [blockIndex, step] of alignment.entries()) {
    if (step.kind === 'equal') {
      slots.push({ kind: 'node', node: step.left });
      continue;
    }
    if (step.kind === 'delete') {
      const deleted = structuralChangeBlock(
        step.left,
        'deletion',
        factory,
        `section-${sectionIndex}-delete-${blockIndex}`,
      );
      if (deleted) slots.push({ kind: 'node', node: deleted });
      else
        diagnostics.push(
          structuralBlockDiagnostic(step.left, sectionIndex, blockIndex),
        );
      continue;
    }
    if (step.kind === 'insert') {
      const inserted = structuralChangeBlock(
        step.right,
        'insertion',
        factory,
        `section-${sectionIndex}-insert-${blockIndex}`,
      );
      if (inserted) slots.push({ kind: 'node', node: inserted });
      else
        diagnostics.push(
          structuralBlockDiagnostic(step.right, sectionIndex, blockIndex),
        );
      continue;
    }
    const prepared = preparePairedBlock(
      step.left,
      step.right,
      `section-${sectionIndex}-block-${blockIndex}`,
    );
    if (prepared) {
      slots.push({ kind: 'paired', comparison: prepared });
      continue;
    }
    const deleted = structuralChangeBlock(
      step.left,
      'deletion',
      factory,
      `section-${sectionIndex}-replace-delete-${blockIndex}`,
    );
    const inserted = structuralChangeBlock(
      step.right,
      'insertion',
      factory,
      `section-${sectionIndex}-replace-insert-${blockIndex}`,
    );
    if (deleted) slots.push({ kind: 'node', node: deleted });
    if (inserted) slots.push({ kind: 'node', node: inserted });
    if (!deleted || !inserted)
      diagnostics.push(
        structuralBlockDiagnostic(step.left, sectionIndex, blockIndex),
      );
  }

  const pairedSlots = slots.filter(
    (slot): slot is Extract<ComparisonBlockSlot, { kind: 'paired' }> =>
      slot.kind === 'paired',
  );
  const initialComparisons = pairedSlots.map(({ comparison }) => ({
    scope: comparison.scope,
    changes: comparison.inline.variant.changes,
  }));
  const initialMovePairs = inferInlineMovePairsAcrossScopes(
    initialComparisons,
    marksEqual,
  );
  const attachedComparisons = pairedSlots.map(({ comparison }) => ({
    scope: comparison.scope,
    changes:
      comparison.inline.attached?.changes ?? comparison.inline.variant.changes,
  }));
  const attachedMovePairs = inferInlineMovePairsAcrossScopes(
    attachedComparisons,
    marksEqual,
  );
  const pairedByScope = new Map(
    pairedSlots.map((slot) => [slot.comparison.scope, slot.comparison]),
  );
  const attachedScopes = new Set<string>();
  for (const pair of [...initialMovePairs, ...attachedMovePairs]) {
    const deletionScope = pair.deletion.scope ?? '';
    const insertionScope = pair.insertion.scope ?? '';
    if (deletionScope === insertionScope) continue;
    const deletionBlock = pairedByScope.get(deletionScope);
    const insertionBlock = pairedByScope.get(insertionScope);
    if (deletionBlock?.inline.attached && insertionBlock?.inline.attached) {
      attachedScopes.add(deletionScope);
      attachedScopes.add(insertionScope);
    }
  }
  const selectedComparisons = pairedSlots.map(({ comparison }) => ({
    scope: comparison.scope,
    changes:
      (attachedScopes.has(comparison.scope)
        ? comparison.inline.attached
        : comparison.inline.variant
      )?.changes ?? comparison.inline.variant.changes,
  }));
  const movePairs = inferInlineMovePairsAcrossScopes(
    selectedComparisons,
    marksEqual,
  );
  const useAttachedVariants =
    attachedScopes.size > 0 && movePairs.length >= initialMovePairs.length;
  const movesByScope = new Map<string, InlineMoveAssignment[]>();
  for (const pair of useAttachedVariants ? movePairs : initialMovePairs) {
    const identity = factory.create('move');
    appendMoveAssignment(movesByScope, pair.deletion, 'from', identity);
    appendMoveAssignment(movesByScope, pair.insertion, 'to', identity);
  }

  return slots.flatMap((slot) => {
    if (slot.kind === 'node') return [slot.node];
    const compared = renderPairedBlock(
      slot.comparison,
      factory,
      movesByScope.get(slot.comparison.scope),
      useAttachedVariants && attachedScopes.has(slot.comparison.scope)
        ? (slot.comparison.inline.attached ?? slot.comparison.inline.variant)
        : slot.comparison.inline.variant,
    );
    return compared ? [compared] : [];
  });
}

function appendMoveAssignment(
  target: Map<string, InlineMoveAssignment[]>,
  candidate: InlineMovePair['deletion'],
  role: InlineMoveAssignment['role'],
  identity: WorkDocumentChangeIdentity,
): void {
  const scope = candidate.scope ?? '';
  const assignments = target.get(scope) ?? [];
  assignments.push({ role, candidate, identity });
  target.set(scope, assignments);
}

function preparePairedBlock(
  current: ProseMirrorNode,
  revised: ProseMirrorNode,
  scope: string,
): PreparedPairedBlock | null {
  if (!isSimpleTextBlock(current) || !isSimpleTextBlock(revised)) return null;
  if (
    current.type !== revised.type ||
    blockStructuralSignature(current) !== blockStructuralSignature(revised)
  ) {
    return null;
  }
  const comparison = prepareInlineComparison(current, revised);
  return comparison ? { current, revised, scope, inline: comparison } : null;
}

function renderPairedBlock(
  comparison: PreparedPairedBlock,
  factory: ComparisonIdentityFactory,
  moves: readonly InlineMoveAssignment[] | undefined,
  variant: InlineComparisonVariant,
): ProseMirrorNode | null {
  const { current, revised } = comparison;
  const inline = renderInlineComparison(current, variant, factory, moves);
  if (!inline) return null;
  const before = serializeDocumentParagraphFormatting(current.attrs);
  const after = serializeDocumentParagraphFormatting(revised.attrs);
  let attributes: Record<string, unknown> = {
    ...current.attrs,
    ...paragraphFormattingAttributes(revised.attrs),
  };
  if (before !== after) {
    const identity = factory.create('paragraph-formatting', before);
    attributes = {
      ...attributes,
      paragraphChangeKind: 'paragraph-formatting',
      paragraphChangeId: identity.id,
      paragraphChangeActorId: identity.actorId ?? '',
      paragraphChangeAuthor: identity.author,
      paragraphChangeDate: identity.date,
      paragraphChangeBefore: before,
    };
  }
  return current.type.create(attributes, inline);
}

function prepareInlineComparison(
  current: ProseMirrorNode,
  revised: ProseMirrorNode,
): PreparedInlineComparison | null {
  const baseCurrentUnits = inlineUnits(current);
  const baseRevisedUnits = inlineUnits(revised);
  if (!baseCurrentUnits || !baseRevisedUnits) return null;
  const diff = boundedDocumentSequenceDiff(
    baseCurrentUnits,
    baseRevisedUnits,
    (left, right) => left.text === right.text,
    MAX_INLINE_DIFF_CELLS,
  );
  const baseVariant: InlineComparisonVariant = {
    currentUnits: baseCurrentUnits,
    revisedUnits: baseRevisedUnits,
    changes: diff ?? [
      { kind: 'delete' as const, left: baseCurrentUnits, right: [] as [] },
      { kind: 'insert' as const, left: [] as [], right: baseRevisedUnits },
    ],
  };
  const variant = baseVariant;
  const baseMovePairs = inferInlineMovePairs(variant.changes, marksEqual);
  // Keep the long-standing comparison boundaries for ordinary edits, but
  // rerun the bounded diff with leading separators attached to lexical units
  // when that exposes a real move. This mirrors Word's move range (the same
  // leading space travels with the phrase) without changing deletion text for
  // documents that contain no inferred move.
  const moveCurrentUnits = inlineUnits(current, true);
  const moveRevisedUnits = inlineUnits(revised, true);
  if (
    moveCurrentUnits &&
    moveRevisedUnits &&
    current.textContent.length + revised.textContent.length <=
      MAX_INFERRED_MOVE_TEXT * 2
  ) {
    const moveDiff = boundedDocumentSequenceDiff(
      moveCurrentUnits,
      moveRevisedUnits,
      (left, right) => left.text === right.text,
      MAX_INLINE_DIFF_CELLS,
    );
    if (moveDiff) {
      const moveChanges = moveDiff as InlineDiffStep[];
      const moveVariant: InlineComparisonVariant = {
        currentUnits: moveCurrentUnits,
        revisedUnits: moveRevisedUnits,
        changes: moveChanges,
      };
      const movePairs = inferInlineMovePairs(moveChanges, marksEqual);
      if (movePairs.length >= baseMovePairs.length && movePairs.length > 0) {
        return { variant: moveVariant, attached: moveVariant };
      }
      return { variant, attached: moveVariant };
    }
  }
  return { variant };
}

function renderInlineComparison(
  current: ProseMirrorNode,
  variant: InlineComparisonVariant,
  factory: ComparisonIdentityFactory,
  moves: readonly InlineMoveAssignment[] | undefined,
): Fragment | null {
  const { changes } = variant;
  const nodes: ProseMirrorNode[] = [];
  for (const [stepIndex, change] of changes.entries()) {
    if (change.kind === 'delete') {
      appendComparisonRevisionUnits(
        nodes,
        change.left,
        'deletion',
        current.type.schema,
        factory,
        withoutReviewMarks,
        appendRevisionUnits,
        moves?.filter((move) => move.candidate.stepIndex === stepIndex),
      );
      continue;
    }
    if (change.kind === 'insert') {
      appendComparisonRevisionUnits(
        nodes,
        change.right,
        'insertion',
        current.type.schema,
        factory,
        withoutReviewMarks,
        appendRevisionUnits,
        moves?.filter((move) => move.candidate.stepIndex === stepIndex),
      );
      continue;
    }
    let formattingIdentity: WorkDocumentChangeIdentity | null = null;
    let formattingSignature = '';
    for (let index = 0; index < change.left.length; index += 1) {
      const left = change.left[index] as InlineUnit;
      const right = change.right[index] as InlineUnit;
      if (marksEqual(left.marks, right.marks)) {
        nodes.push(current.type.schema.text(left.text, [...left.marks]));
        formattingIdentity = null;
        formattingSignature = '';
        continue;
      }
      if (nonCharacterMarksEqual(left.marks, right.marks)) {
        const before = serializeDocumentCharacterFormatting(left.marks);
        const after = serializeDocumentCharacterFormatting(right.marks);
        const signature = `${before}\u0000${after}`;
        if (!formattingIdentity || formattingSignature !== signature) {
          formattingIdentity = factory.create('formatting', before);
          formattingSignature = signature;
        }
        nodes.push(
          current.type.schema.text(right.text, [
            ...withoutReviewMarks(right.marks),
            comparisonMark(current, 'formatting', formattingIdentity, before),
          ]),
        );
        continue;
      }
      const deletion = factory.create('deletion');
      const insertion = factory.create('insertion');
      appendRevisionUnits(
        nodes,
        [left],
        'deletion',
        deletion,
        current.type.schema,
        withoutReviewMarks,
      );
      appendRevisionUnits(
        nodes,
        [right],
        'insertion',
        insertion,
        current.type.schema,
        withoutReviewMarks,
      );
      formattingIdentity = null;
      formattingSignature = '';
    }
  }
  return Fragment.fromArray(nodes);
}

function structuralChangeBlock(
  block: ProseMirrorNode,
  kind: 'insertion' | 'deletion',
  factory: ComparisonIdentityFactory,
  identityChannel: string,
): ProseMirrorNode | null {
  if (!isSimpleTextBlock(block) || !block.textContent) return null;
  const identity = factory.create(kind);
  const paragraphIdentity =
    kind === 'insertion'
      ? factory.paragraphIdentity(identityChannel)
      : {
          paragraphId: block.attrs.paragraphId,
          textId: block.attrs.textId,
        };
  const nodes: ProseMirrorNode[] = [];
  for (const unit of inlineUnits(block) ?? []) {
    appendRevisionUnits(
      nodes,
      [unit],
      kind,
      identity,
      block.type.schema,
      withoutReviewMarks,
    );
  }
  return block.type.create(
    {
      ...block.attrs,
      ...paragraphIdentity,
      blockChangeKind: kind,
      blockChangeId: identity.id,
      blockChangeActorId: identity.actorId ?? '',
      blockChangeAuthor: identity.author,
      blockChangeDate: identity.date,
    },
    Fragment.fromArray(nodes),
  );
}

function comparisonMark(
  node: ProseMirrorNode,
  kind: WorkDocumentChangeKind,
  identity: WorkDocumentChangeIdentity,
  before: string,
): ProseMirrorMark {
  return node.type.schema.marks.documentChange.create({
    kind,
    id: identity.id,
    actorId: identity.actorId ?? '',
    author: identity.author,
    date: identity.date,
    before,
  });
}

function parseComparisonDocument(
  editor: Editor,
  content: Content,
): ProseMirrorNode | null {
  try {
    const parsed = createNodeFromContent(content, editor.schema, {
      errorOnInvalidContent: true,
      slice: false,
    });
    return parsed instanceof ProseMirrorNode ? parsed : null;
  } catch {
    return null;
  }
}

function documentSections(document: ProseMirrorNode): ProseMirrorNode[] | null {
  const sections = childNodes(document);
  return sections.every((node) => node.type.name === 'documentSection')
    ? sections
    : null;
}

function comparisonLimits(
  current: readonly ProseMirrorNode[],
  revised: readonly ProseMirrorNode[],
): DocumentComparisonDiagnostic | null {
  const currentBlocks = current.reduce(
    (sum, section) => sum + section.childCount,
    0,
  );
  const revisedBlocks = revised.reduce(
    (sum, section) => sum + section.childCount,
    0,
  );
  const textLength =
    current.reduce((sum, section) => sum + section.textContent.length, 0) +
    revised.reduce((sum, section) => sum + section.textContent.length, 0);
  return currentBlocks > MAX_COMPARISON_BLOCKS ||
    revisedBlocks > MAX_COMPARISON_BLOCKS ||
    textLength > MAX_COMPARISON_TEXT
    ? {
        code: 'comparison-limit-exceeded',
        message: `Comparison accepts at most ${MAX_COMPARISON_BLOCKS} blocks per version and ${MAX_COMPARISON_TEXT} combined text characters.`,
      }
    : null;
}

function blockSubstitutionCost(
  current: ProseMirrorNode,
  revised: ProseMirrorNode,
): number {
  if (
    comparisonSemanticSignature(current) ===
    comparisonSemanticSignature(revised)
  )
    return 0;
  if (!isSimpleTextBlock(current) || !isSimpleTextBlock(revised)) return 2.1;
  if (current.type !== revised.type) return 2.1;
  if (!current.textContent && !revised.textContent) return 0.8;
  return documentTextSimilarity(current.textContent, revised.textContent) >=
    0.18
    ? 0.8
    : 2.1;
}

function documentTextSimilarity(left: string, right: string): number {
  if (left === right) return 1;
  if (!left || !right) return 0;
  const maximum = Math.max(left.length, right.length);
  let prefix = 0;
  while (
    prefix < Math.min(left.length, right.length) &&
    left[prefix] === right[prefix]
  ) {
    prefix += 1;
  }
  let suffix = 0;
  while (
    suffix < Math.min(left.length, right.length) - prefix &&
    left[left.length - suffix - 1] === right[right.length - suffix - 1]
  ) {
    suffix += 1;
  }
  const boundaryScore = (prefix + suffix) / maximum;
  const leftWords = comparisonWords(left);
  const rightWords = comparisonWords(right);
  let shared = 0;
  for (const word of leftWords) if (rightWords.has(word)) shared += 1;
  const union = new Set([...leftWords, ...rightWords]).size;
  return Math.max(boundaryScore, union ? shared / union : 0);
}

function comparisonWords(value: string): Set<string> {
  return new Set(
    (value.toLocaleLowerCase().match(/[\p{L}\p{N}_]+/gu) ?? []).slice(0, 256),
  );
}

function isSimpleTextBlock(node: ProseMirrorNode): boolean {
  if (node.type.name !== 'paragraph' && node.type.name !== 'heading')
    return false;
  let simple = true;
  node.forEach((child) => {
    if (
      !child.isText ||
      child.marks.some((mark) => REVIEW_ONLY_MARK_NAMES.has(mark.type.name))
    ) {
      simple = false;
    }
  });
  return simple;
}

function simpleDocumentShape(document: ProseMirrorNode): boolean {
  const sections = documentSections(document);
  return Boolean(
    sections?.every((section) =>
      childNodes(section).every((block) => isSimpleTextBlock(block)),
    ),
  );
}

function hasStructuralBlockRevisions(document: ProseMirrorNode): boolean {
  let found = false;
  document.descendants((node) => {
    if (
      (node.type.name === 'paragraph' || node.type.name === 'heading') &&
      (node.attrs.blockChangeKind === 'insertion' ||
        node.attrs.blockChangeKind === 'deletion')
    ) {
      found = true;
    }
    return !found;
  });
  return found;
}

function inlineUnits(
  node: ProseMirrorNode,
  attachLeadingWhitespace = false,
): InlineUnit[] | null {
  if (!isSimpleTextBlock(node)) return null;
  const units: InlineUnit[] = [];
  node.forEach((child) => {
    const text = child.text ?? '';
    let pendingWhitespace = '';
    for (const part of text.match(/\s+|[\p{L}\p{N}_]+|[^\s\p{L}\p{N}_]/gu) ??
      []) {
      if (attachLeadingWhitespace && part.trim() === '') {
        pendingWhitespace += part;
        continue;
      }
      units.push({
        text: attachLeadingWhitespace ? `${pendingWhitespace}${part}` : part,
        marks: child.marks,
        ...(attachLeadingWhitespace ? { leadingWhitespaceAttached: true } : {}),
      });
      pendingWhitespace = '';
    }
    if (attachLeadingWhitespace && pendingWhitespace) {
      units.push({
        text: pendingWhitespace,
        marks: child.marks,
        leadingWhitespaceAttached: true,
      });
    }
  });
  return units;
}

function marksEqual(
  left: readonly ProseMirrorMark[],
  right: readonly ProseMirrorMark[],
): boolean {
  const normalizedLeft = withoutReviewMarks(left);
  const normalizedRight = withoutReviewMarks(right);
  return (
    normalizedLeft.length === normalizedRight.length &&
    normalizedLeft.every((mark, index) =>
      mark.eq(normalizedRight[index] as ProseMirrorMark),
    )
  );
}

function nonCharacterMarksEqual(
  left: readonly ProseMirrorMark[],
  right: readonly ProseMirrorMark[],
): boolean {
  const normalizedLeft = withoutReviewMarks(left).filter(
    (mark) => !isDocumentCharacterFormatMark(mark.type.name),
  );
  const normalizedRight = withoutReviewMarks(right).filter(
    (mark) => !isDocumentCharacterFormatMark(mark.type.name),
  );
  return (
    normalizedLeft.length === normalizedRight.length &&
    normalizedLeft.every((mark, index) =>
      mark.eq(normalizedRight[index] as ProseMirrorMark),
    )
  );
}

function withoutReviewMarks(
  marks: readonly ProseMirrorMark[],
): ProseMirrorMark[] {
  return marks.filter((mark) => !REVIEW_ONLY_MARK_NAMES.has(mark.type.name));
}

function blockStructuralSignature(node: ProseMirrorNode): string {
  const ignored = new Set<string>([
    ...DOCUMENT_PARAGRAPH_FORMAT_ATTRIBUTES,
    ...DOCUMENT_PARAGRAPH_CHANGE_ATTRIBUTES,
    ...BLOCK_CHANGE_ATTRIBUTE_NAMES,
    ...PARAGRAPH_IDENTITY_ATTRIBUTE_NAMES,
  ]);
  return stableJson({
    attrs: filteredAttributes(node.attrs, ignored),
    type: node.type.name,
  });
}

function paragraphFormattingAttributes(
  attributes: Record<string, unknown>,
): Record<string, unknown> {
  return Object.fromEntries(
    DOCUMENT_PARAGRAPH_FORMAT_ATTRIBUTES.map((name) => [
      name,
      attributes[name],
    ]),
  );
}

function sectionLayoutSignature(section: ProseMirrorNode): string {
  return stableJson(filteredAttributes(section.attrs, new Set(['id'])));
}

function comparisonSemanticSignature(node: ProseMirrorNode): string {
  return stableJson(comparisonSemanticValue(node));
}

function comparisonSemanticValue(node: ProseMirrorNode): unknown {
  const ignored = new Set<string>([
    ...DOCUMENT_PARAGRAPH_CHANGE_ATTRIBUTES,
    ...BLOCK_CHANGE_ATTRIBUTE_NAMES,
    ...PARAGRAPH_IDENTITY_ATTRIBUTE_NAMES,
    ...(node.type.name === 'documentSection' ? ['id'] : []),
  ]);
  return {
    type: node.type.name,
    ...(node.isText ? { text: node.text } : {}),
    ...(Object.keys(node.attrs).length
      ? { attrs: filteredAttributes(node.attrs, ignored) }
      : {}),
    ...(node.marks.length
      ? { marks: withoutReviewMarks(node.marks).map((mark) => mark.toJSON()) }
      : {}),
    ...(node.childCount
      ? { content: childNodes(node).map(comparisonSemanticValue) }
      : {}),
  };
}

function transferCombineIdentities(
  current: ProseMirrorNode,
  reviewed: ProseMirrorNode,
): ProseMirrorNode | null {
  if (current.type !== reviewed.type) return null;
  if (reviewed.type.name === 'paragraph' || reviewed.type.name === 'heading') {
    return reviewed.type.create(
      {
        ...reviewed.attrs,
        paragraphId: current.attrs.paragraphId,
        textId: current.attrs.textId,
      },
      reviewed.content,
      reviewed.marks,
    );
  }
  if (current.childCount !== reviewed.childCount) return null;
  const content: ProseMirrorNode[] = [];
  for (let index = 0; index < current.childCount; index += 1) {
    const currentChild = current.child(index);
    const reviewedChild = reviewed.child(index);
    const transferred = transferCombineIdentities(currentChild, reviewedChild);
    if (!transferred) return null;
    content.push(transferred);
  }
  const attributes = { ...reviewed.attrs };
  if (reviewed.type.name === 'documentSection')
    attributes.id = current.attrs.id;
  return reviewed.isText
    ? reviewed
    : reviewed.type.create(
        attributes,
        Fragment.fromArray(content),
        reviewed.marks,
      );
}

function structuralBlockDiagnostic(
  block: ProseMirrorNode,
  section: number,
  blockIndex: number,
): DocumentComparisonDiagnostic {
  return block.type.name === 'paragraph' || block.type.name === 'heading'
    ? {
        code: 'empty-structural-change',
        message:
          'An empty inserted or deleted paragraph cannot carry a reviewable text revision.',
        section,
        block: blockIndex,
      }
    : {
        code: 'changed-complex-structure',
        message: `Changed ${block.type.name} content is not approximated as plain text.`,
        section,
        block: blockIndex,
      };
}

function unsupportedResult(
  diagnostic: DocumentComparisonDiagnostic,
): DocumentComparisonApplyResult {
  return {
    status: 'unsupported',
    summary: emptyComparisonSummary(),
    diagnostics: [diagnostic],
  };
}

function childNodes(node: ProseMirrorNode): ProseMirrorNode[] {
  const children: ProseMirrorNode[] = [];
  node.forEach((child) => {
    children.push(child);
  });
  return children;
}

function filteredAttributes(
  attributes: Record<string, unknown>,
  ignored: ReadonlySet<string>,
): Record<string, unknown> {
  return Object.fromEntries(
    Object.entries(attributes).filter(([name]) => !ignored.has(name)),
  );
}
