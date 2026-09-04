import type {
  Mark as ProseMirrorMark,
  Node as ProseMirrorNode,
  Schema,
} from '@tiptap/pm/model';
import type {
  WorkDocumentChangeIdentity,
  WorkDocumentChangeKind,
} from './work-document-changes';
import type { WorkDocumentMoveRole } from './work-types';

export interface InlineUnit {
  text: string;
  marks: readonly ProseMirrorMark[];
  leadingWhitespaceAttached?: boolean;
}

export type InlineDiffStep =
  | { kind: 'equal'; left: InlineUnit[]; right: InlineUnit[] }
  | { kind: 'delete'; left: InlineUnit[]; right: [] }
  | { kind: 'insert'; left: []; right: InlineUnit[] };

export interface InlineMoveCandidate {
  stepIndex: number;
  kind: 'delete' | 'insert';
  units: InlineUnit[];
  start: number;
  end: number;
  startOffset: number;
  endOffset: number;
  text: string;
}

export interface InlineMovePair {
  deletion: InlineMoveCandidate;
  insertion: InlineMoveCandidate;
}

export interface InlineMoveAssignment {
  role: WorkDocumentMoveRole;
  candidate: InlineMoveCandidate;
  identity: WorkDocumentChangeIdentity;
}

export interface ComparisonMoveIdentityFactory {
  create(kind: WorkDocumentChangeKind): WorkDocumentChangeIdentity;
}

type MarksEqual = (
  left: readonly ProseMirrorMark[],
  right: readonly ProseMirrorMark[],
) => boolean;

type StripReviewMarks = (
  marks: readonly ProseMirrorMark[],
) => ProseMirrorMark[];

type AppendRevisionUnits = (
  target: ProseMirrorNode[],
  units: readonly InlineUnit[],
  kind: 'insertion' | 'deletion',
  identity: WorkDocumentChangeIdentity,
  schema: Schema,
  stripReviewMarks: StripReviewMarks,
) => void;

export const MAX_INFERRED_MOVE_TEXT = 65_536;
const MAX_INFERRED_MOVES = 256;

export function appendRevisionUnits(
  target: ProseMirrorNode[],
  units: readonly InlineUnit[],
  kind: 'insertion' | 'deletion',
  identity: WorkDocumentChangeIdentity,
  schema: Schema,
  stripReviewMarks: StripReviewMarks,
): void {
  for (const unit of units) {
    target.push(
      schema.text(unit.text, [
        ...stripReviewMarks(unit.marks),
        schema.marks.documentChange.create({
          kind,
          id: identity.id,
          actorId: identity.actorId ?? '',
          author: identity.author,
          date: identity.date,
          before: '',
        }),
      ]),
    );
  }
}

/**
 * Pairs only deterministic same-text ranges inside one block. Duplicate
 * candidates and ranges whose marks differ are deliberately left as ordinary
 * insertions/deletions so Compare never guesses at an ambiguous move.
 */
export function inferInlineMovePairs(
  changes: readonly InlineDiffStep[],
  marksEqual: MarksEqual,
): InlineMovePair[] {
  const fullCandidates: InlineMoveCandidate[] = [];
  for (const [stepIndex, change] of changes.entries()) {
    if (change.kind === 'equal') continue;
    const units = change.kind === 'delete' ? change.left : change.right;
    const candidate = inlineMoveCandidate(stepIndex, change.kind, units);
    if (candidate) fullCandidates.push(candidate);
  }
  const fullPairs = pairInlineMoveCandidates(fullCandidates, marksEqual);
  const consumedSteps = new Set<number>();
  for (const pair of fullPairs) {
    consumedSteps.add(pair.deletion.stepIndex);
    consumedSteps.add(pair.insertion.stepIndex);
  }

  // A sequence diff can coalesce adjacent reordered words into one delete and
  // one insert chunk. Pair unique lexical units in that case; duplicate words
  // remain ordinary revisions instead of being assigned to an arbitrary side.
  const tokenCandidates: InlineMoveCandidate[] = [];
  for (const [stepIndex, change] of changes.entries()) {
    if (change.kind === 'equal' || consumedSteps.has(stepIndex)) continue;
    const units = change.kind === 'delete' ? change.left : change.right;
    tokenCandidates.push(
      ...inlineMoveTokenCandidates(stepIndex, change.kind, units),
    );
  }
  return [
    ...fullPairs,
    ...pairInlineMoveCandidates(tokenCandidates, marksEqual),
  ]
    .sort(
      (left, right) =>
        left.deletion.stepIndex - right.deletion.stepIndex ||
        left.insertion.stepIndex - right.insertion.stepIndex ||
        left.deletion.start - right.deletion.start,
    )
    .slice(0, MAX_INFERRED_MOVES);
}

/**
 * Splits an ordinary diff step around paired move ranges while retaining the
 * exact whitespace and mark boundaries needed to reconstruct either version.
 */
export function appendComparisonRevisionUnits(
  target: ProseMirrorNode[],
  units: InlineUnit[],
  kind: 'insertion' | 'deletion',
  schema: Schema,
  factory: ComparisonMoveIdentityFactory,
  stripReviewMarks: StripReviewMarks,
  appendRevisionUnits: AppendRevisionUnits,
  moves?: readonly InlineMoveAssignment[],
): void {
  if (!moves?.length) {
    appendRevisionUnits(
      target,
      units,
      kind,
      factory.create(kind),
      schema,
      stripReviewMarks,
    );
    return;
  }
  const byUnit = new Map<number, InlineMoveAssignment>();
  for (const move of [...moves].sort(
    (left, right) => left.candidate.start - right.candidate.start,
  )) {
    for (
      let index = move.candidate.start;
      index < move.candidate.end;
      index += 1
    ) {
      if (byUnit.has(index)) {
        // Overlapping candidates are not safe to represent as one native move
        // range. Leave the complete step as an ordinary revision rather than
        // emitting ambiguous nested identities.
        appendRevisionUnits(
          target,
          units,
          kind,
          factory.create(kind),
          schema,
          stripReviewMarks,
        );
        return;
      }
      byUnit.set(index, move);
    }
  }
  let ordinary: InlineUnit[] = [];
  const flushOrdinary = () => {
    if (!ordinary.length) return;
    appendRevisionUnits(
      target,
      ordinary,
      kind,
      factory.create(kind),
      schema,
      stripReviewMarks,
    );
    ordinary = [];
  };
  for (const [index, unit] of units.entries()) {
    const move = byUnit.get(index);
    if (!move) {
      ordinary.push(unit);
      continue;
    }
    const start =
      index === move.candidate.start ? move.candidate.startOffset : 0;
    const end =
      index === move.candidate.end - 1
        ? move.candidate.endOffset
        : unit.text.length;
    const prefix = unit.text.slice(0, start);
    const marked = unit.text.slice(start, end);
    const suffix = unit.text.slice(end);
    if (prefix) ordinary.push({ text: prefix, marks: unit.marks });
    flushOrdinary();
    if (marked) {
      appendMoveRevisionUnits(
        target,
        [{ text: marked, marks: unit.marks }],
        move.role,
        move.identity,
        schema,
        stripReviewMarks,
      );
    }
    ordinary = suffix ? [{ text: suffix, marks: unit.marks }] : [];
  }
  flushOrdinary();
}

function pairInlineMoveCandidates(
  candidates: readonly InlineMoveCandidate[],
  marksEqual: MarksEqual,
): InlineMovePair[] {
  const grouped = new Map<
    string,
    { deletion: InlineMoveCandidate[]; insertion: InlineMoveCandidate[] }
  >();
  for (const candidate of candidates) {
    const group = grouped.get(candidate.text) ?? {
      deletion: [],
      insertion: [],
    };
    (candidate.kind === 'delete' ? group.deletion : group.insertion).push(
      candidate,
    );
    grouped.set(candidate.text, group);
  }
  return Array.from(grouped.values())
    .filter(
      ({ deletion, insertion }) =>
        deletion.length === 1 &&
        insertion.length === 1 &&
        inlineMoveCandidatesEqual(
          deletion[0] as InlineMoveCandidate,
          insertion[0] as InlineMoveCandidate,
          marksEqual,
        ),
    )
    .map(({ deletion, insertion }) => ({
      deletion: deletion[0] as InlineMoveCandidate,
      insertion: insertion[0] as InlineMoveCandidate,
    }))
    .sort(
      (left, right) =>
        left.deletion.stepIndex - right.deletion.stepIndex ||
        left.insertion.stepIndex - right.insertion.stepIndex,
    )
    .slice(0, MAX_INFERRED_MOVES);
}

function inlineMoveCandidate(
  stepIndex: number,
  kind: 'delete' | 'insert',
  units: InlineUnit[],
): InlineMoveCandidate | null {
  let start = 0;
  while (start < units.length && isWhitespaceUnit(units[start] as InlineUnit)) {
    start += 1;
  }
  let end = units.length;
  while (end > start && isWhitespaceUnit(units[end - 1] as InlineUnit)) {
    end -= 1;
  }
  if (start === end) return null;
  const first = units[start] as InlineUnit;
  const last = units[end - 1] as InlineUnit;
  const startOffset = first.leadingWhitespaceAttached
    ? 0
    : first.text.length - first.text.trimStart().length;
  const endOffset = last.leadingWhitespaceAttached
    ? last.text.length
    : last.text.trimEnd().length;
  const text = units
    .slice(start, end)
    .map((unit, index, selected) =>
      unit.text.slice(
        index === 0 ? startOffset : 0,
        index === selected.length - 1 ? endOffset : unit.text.length,
      ),
    )
    .join('');
  if (
    !text ||
    text.length > MAX_INFERRED_MOVE_TEXT ||
    !/[\p{L}\p{N}]/u.test(text)
  ) {
    return null;
  }
  return {
    stepIndex,
    kind,
    units,
    start,
    end,
    startOffset,
    endOffset,
    text,
  };
}

function inlineMoveTokenCandidates(
  stepIndex: number,
  kind: 'delete' | 'insert',
  units: InlineUnit[],
): InlineMoveCandidate[] {
  const candidates: InlineMoveCandidate[] = [];
  for (const [index, unit] of units.entries()) {
    const leadingWhitespace = unit.text.length - unit.text.trimStart().length;
    const text = unit.text.slice(leadingWhitespace);
    if (
      !text ||
      text.length > MAX_INFERRED_MOVE_TEXT ||
      !/[\p{L}\p{N}]/u.test(text)
    ) {
      continue;
    }
    candidates.push({
      stepIndex,
      kind,
      units,
      start: index,
      end: index + 1,
      startOffset: leadingWhitespace,
      endOffset: unit.text.length,
      text,
    });
  }
  return candidates;
}

function inlineMoveCandidatesEqual(
  left: InlineMoveCandidate,
  right: InlineMoveCandidate,
  marksEqual: MarksEqual,
): boolean {
  const leftUnits = selectedMoveUnits(left);
  const rightUnits = selectedMoveUnits(right);
  return (
    leftUnits.length === rightUnits.length &&
    leftUnits.every(
      (unit, index) =>
        unit.text === rightUnits[index]?.text &&
        marksEqual(unit.marks, rightUnits[index]?.marks ?? []),
    )
  );
}

function selectedMoveUnits(candidate: InlineMoveCandidate): InlineUnit[] {
  return candidate.units
    .slice(candidate.start, candidate.end)
    .map((unit, index, selected) => ({
      text: unit.text.slice(
        index === 0 ? candidate.startOffset : 0,
        index === selected.length - 1 ? candidate.endOffset : unit.text.length,
      ),
      marks: unit.marks,
    }))
    .filter((unit) => unit.text.length > 0);
}

function isWhitespaceUnit(unit: InlineUnit): boolean {
  return unit.text.trim() === '';
}

function appendMoveRevisionUnits(
  target: ProseMirrorNode[],
  units: readonly InlineUnit[],
  role: WorkDocumentMoveRole,
  identity: WorkDocumentChangeIdentity,
  schema: Schema,
  stripReviewMarks: StripReviewMarks,
): void {
  for (const unit of units) {
    target.push(
      schema.text(unit.text, [
        ...stripReviewMarks(unit.marks),
        schema.marks.documentChange.create({
          kind: 'move',
          moveRole: role,
          id: identity.id,
          actorId: identity.actorId ?? '',
          author: identity.author,
          date: identity.date,
          before: '',
        }),
      ]),
    );
  }
}
