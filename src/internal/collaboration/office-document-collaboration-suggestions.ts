import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import type { Transaction } from '@tiptap/pm/state';
import { AddMarkStep, RemoveMarkStep, ReplaceStep } from '@tiptap/pm/transform';
import type {
  WorkDocumentChange,
  WorkDocumentChangeIdentity,
} from '../features/work/work-document-changes';
import type {
  WorkDocumentChangeDecision,
  WorkDocumentChangeDecisionAction,
  WorkDocumentChangeKind,
} from '../features/work/work-types';
import { parseDocumentCharacterFormatting } from '../features/work/work-document-format-changes';
import type { WorkOfficeCollaborationSession } from './office-collaboration';

interface StrictDocumentChange extends WorkDocumentChangeIdentity {
  kind: WorkDocumentChangeKind;
  text: string;
}

export function workOfficeDocumentSuggestionTransactionAllowed(
  session: WorkOfficeCollaborationSession,
  transaction: Transaction,
): boolean {
  const actor = session.actor;
  if (!actor || !transaction.docChanged || transaction.steps.length === 0) {
    return false;
  }
  if (
    !transaction.steps.every(
      (step) =>
        step instanceof ReplaceStep ||
        ((step instanceof AddMarkStep || step instanceof RemoveMarkStep) &&
          step.mark.type.name === 'documentChange'),
    )
  ) {
    return false;
  }

  const before = strictDocumentChanges(transaction.before);
  const after = strictDocumentChanges(transaction.doc);
  if (!before || !after) return false;
  if (
    JSON.stringify(documentSuggestionBaseline(transaction.before)) !==
    JSON.stringify(documentSuggestionBaseline(transaction.doc))
  ) {
    return false;
  }

  const beforeById = new Map(before.map((change) => [change.id, change]));
  const afterById = new Map(after.map((change) => [change.id, change]));
  for (const previous of before) {
    const current = afterById.get(previous.id);
    if (!current) {
      if (previous.actorId !== actor.id) return false;
      continue;
    }
    if (!sameChangeIdentity(previous, current)) return false;
    if (previous.actorId !== actor.id && previous.text !== current.text) {
      return false;
    }
    if (
      previous.actorId === actor.id &&
      previous.kind === 'deletion' &&
      previous.text !== current.text
    ) {
      return false;
    }
  }
  for (const current of after) {
    if (beforeById.has(current.id)) continue;
    if (
      current.actorId !== actor.id ||
      current.author !== actor.name ||
      !current.date ||
      !current.text
    ) {
      return false;
    }
  }
  return true;
}

export function createWorkOfficeDocumentChangeDecisions(
  session: WorkOfficeCollaborationSession,
  changes: readonly WorkDocumentChange[],
  decision: WorkDocumentChangeDecisionAction,
  decidedAt = new Date().toISOString(),
): WorkDocumentChangeDecision[] {
  const decidedBy = session.actor?.name ?? 'A3S Work 用户';
  return changes.map((change) => ({
    id: `${change.kind}:${change.id}`,
    changeId: change.id,
    changeKind: change.kind,
    ...(change.actorId ? { suggestedByActorId: change.actorId } : {}),
    suggestedBy: change.author,
    suggestedAt: change.date,
    text: change.text,
    decision,
    ...(session.actor ? { decidedByActorId: session.actor.id } : {}),
    decidedBy,
    decidedAt,
  }));
}

function strictDocumentChanges(
  document: ProseMirrorNode,
): StrictDocumentChange[] | null {
  const changes = new Map<string, StrictDocumentChange>();
  let valid = true;
  document.descendants((node) => {
    if (!valid || !node.isText || !node.text) return;
    const marks = node.marks.filter(
      (mark) => mark.type.name === 'documentChange',
    );
    if (marks.length === 0) return;
    if (marks.length !== 1) {
      valid = false;
      return false;
    }
    const mark = marks[0];
    const id = strictString(mark.attrs.id);
    const kind = strictChangeKind(mark.attrs.kind);
    const author = strictString(mark.attrs.author);
    const date = strictString(mark.attrs.date);
    const actorId = optionalStrictString(mark.attrs.actorId);
    if (
      !id ||
      !kind ||
      author === null ||
      date === null ||
      actorId === null ||
      (kind === 'formatting' &&
        !parseDocumentCharacterFormatting(mark.attrs.before))
    ) {
      valid = false;
      return false;
    }
    if (kind === 'formatting') return;
    const current = changes.get(id);
    const candidate: StrictDocumentChange = {
      id,
      kind,
      ...(actorId ? { actorId } : {}),
      author,
      date,
      text: node.text,
    };
    if (!current) {
      changes.set(id, candidate);
      return;
    }
    if (!sameChangeIdentity(current, candidate)) {
      valid = false;
      return false;
    }
    current.text += node.text;
  });
  return valid ? Array.from(changes.values()) : null;
}

function documentSuggestionBaseline(node: ProseMirrorNode): unknown {
  const json = node.toJSON();
  if (node.isText) {
    const change = node.marks.find(
      (mark) => mark.type.name === 'documentChange',
    );
    if (change?.attrs.kind === 'insertion') return null;
    const marks = node.marks
      .filter(
        (mark) =>
          mark.type.name !== 'documentChange' ||
          mark.attrs.kind === 'formatting',
      )
      .map((mark) => mark.toJSON());
    if (marks.length > 0) return { ...json, marks };
    const { marks: _marks, ...withoutMarks } = json;
    return withoutMarks;
  }
  const content: unknown[] = [];
  node.forEach((child) => {
    const projected = documentSuggestionBaseline(child);
    if (projected !== null) appendSuggestionBaselineNode(content, projected);
  });
  if (content.length > 0) return { ...json, content };
  const { content: _content, ...withoutContent } = json;
  return withoutContent;
}

function appendSuggestionBaselineNode(
  content: unknown[],
  projected: unknown,
): void {
  const previous = content.at(-1);
  if (
    isSuggestionBaselineText(previous) &&
    isSuggestionBaselineText(projected) &&
    JSON.stringify(previous.marks ?? []) ===
      JSON.stringify(projected.marks ?? [])
  ) {
    content[content.length - 1] = {
      ...previous,
      text: previous.text + projected.text,
    };
    return;
  }
  content.push(projected);
}

function isSuggestionBaselineText(
  value: unknown,
): value is { type: 'text'; text: string; marks?: unknown } {
  if (!value || typeof value !== 'object') return false;
  const candidate = value as { type?: unknown; text?: unknown };
  return candidate.type === 'text' && typeof candidate.text === 'string';
}

function sameChangeIdentity(
  left: StrictDocumentChange,
  right: StrictDocumentChange,
): boolean {
  return (
    left.id === right.id &&
    left.kind === right.kind &&
    left.actorId === right.actorId &&
    left.author === right.author &&
    left.date === right.date
  );
}

function strictString(value: unknown): string | null {
  return typeof value === 'string' && value.trim() === value && value
    ? value
    : null;
}

function optionalStrictString(value: unknown): string | undefined | null {
  if (value === undefined || value === '') return undefined;
  return strictString(value);
}

function strictChangeKind(value: unknown): WorkDocumentChangeKind | null {
  return value === 'insertion' || value === 'deletion' || value === 'formatting'
    ? value
    : null;
}
