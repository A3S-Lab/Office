import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { type EditorState, Plugin, type Transaction } from '@tiptap/pm/state';
import { Mapping } from '@tiptap/pm/transform';
import {
  DOCUMENT_INTEGRITY_NOTE,
  documentHasIntegrityFeature,
} from './work-document-integrity-index';
import {
  documentNoteKey,
  documentNoteKind,
  type WorkDocumentNoteKind,
} from './work-document-notes';
import { createWorkId } from './work-templates';

interface NoteNodeAtPosition {
  node: ProseMirrorNode;
  position: number;
  id: string;
  kind: WorkDocumentNoteKind;
  key: string;
}

interface NoteReferenceAssignment {
  reference: NoteNodeAtPosition;
  id: string;
  kind: WorkDocumentNoteKind;
  key: string;
  number: number;
}

interface NoteDefinitionInsertion {
  assignment: NoteReferenceAssignment;
  source: ProseMirrorNode | null;
  position: number;
}

export function createDocumentNoteIntegrityPlugin(
  referenceNodeName = 'documentNoteReference',
  definitionNodeName = 'documentNote',
): Plugin {
  return new Plugin({
    view(view) {
      const transaction = normalizeDocumentNoteGraph(
        view.state,
        referenceNodeName,
        definitionNodeName,
      );
      if (transaction) {
        transaction.setMeta('addToHistory', false);
        view.dispatch(transaction);
      }
      return {};
    },
    appendTransaction(transactions, oldState, newState) {
      if (!transactions.some((transaction) => transaction.docChanged)) {
        return null;
      }
      return normalizeDocumentNoteGraph(
        newState,
        referenceNodeName,
        definitionNodeName,
        oldState,
        transactions,
      );
    },
  });
}

function normalizeDocumentNoteGraph(
  state: EditorState,
  referenceNodeName: string,
  definitionNodeName: string,
  oldState?: EditorState,
  transactions: readonly Transaction[] = [],
): Transaction | null {
  if (
    referenceNodeName === 'documentNoteReference' &&
    definitionNodeName === 'documentNote' &&
    !documentHasIntegrityFeature(state.doc, DOCUMENT_INTEGRITY_NOTE)
  ) {
    return null;
  }
  const allReferences = documentNoteNodes(state.doc, referenceNodeName);
  const nestedReferences = allReferences.filter((reference) =>
    positionInsideNode(state.doc, reference.position, definitionNodeName),
  );
  const nestedReferencePositions = new Set(
    nestedReferences.map(({ position }) => position),
  );
  const references = allReferences.filter(
    (reference) => !nestedReferencePositions.has(reference.position),
  );
  const definitions = documentNoteNodes(state.doc, definitionNodeName);
  if (!references.length && !definitions.length) return null;

  const retainedReferences = oldState
    ? retainedNoteNodePositions(
        oldState,
        state,
        transactions,
        referenceNodeName,
      )
    : new Set<number>();
  const retainedDefinitions = oldState
    ? retainedNoteNodePositions(
        oldState,
        state,
        transactions,
        definitionNodeName,
      )
    : new Set<number>();
  const definitionsByKey = groupNoteNodesByKey(definitions);
  const oldDefinitionsByKey = oldState
    ? groupNoteNodesByKey(documentNoteNodes(oldState.doc, definitionNodeName))
    : new Map<string, NoteNodeAtPosition[]>();
  const deletedReferences = new Set<number>();
  if (oldState) {
    for (const reference of references) {
      if (
        retainedReferences.has(reference.position) &&
        reference.key &&
        oldDefinitionsByKey.has(reference.key) &&
        !definitionsByKey.has(reference.key)
      ) {
        deletedReferences.add(reference.position);
      }
    }
  }

  const keptReferences = references.filter(
    (reference) => !deletedReferences.has(reference.position),
  );
  const assignmentByPosition = assignNoteReferenceIdentities(
    keptReferences,
    retainedReferences,
  );
  const assignments = keptReferences.map(
    (reference) =>
      assignmentByPosition.get(reference.position) as NoteReferenceAssignment,
  );
  assignNoteNumbers(assignments);

  const usedDefinitions = new Set<number>();
  const definitionUpdates = new Map<number, NoteReferenceAssignment>();
  const definitionInsertions: NoteDefinitionInsertion[] = [];
  const assignmentsByRetention = [
    ...assignments.filter((assignment) =>
      retainedReferences.has(assignment.reference.position),
    ),
    ...assignments.filter(
      (assignment) => !retainedReferences.has(assignment.reference.position),
    ),
  ];
  for (const assignment of assignmentsByRetention) {
    const candidates = definitionsByKey.get(assignment.reference.key) ?? [];
    const definition = preferredDefinition(
      candidates,
      usedDefinitions,
      retainedDefinitions,
      assignment.key === assignment.reference.key,
    );
    if (definition) {
      usedDefinitions.add(definition.position);
      definitionUpdates.set(definition.position, assignment);
      continue;
    }
    const source =
      candidates[0]?.node ??
      oldDefinitionsByKey.get(assignment.reference.key)?.[0]?.node ??
      null;
    definitionInsertions.push({
      assignment,
      source,
      position: noteDefinitionInsertionPosition(
        state.doc,
        assignment.reference.position,
        assignment.kind,
      ),
    });
  }

  const transaction = state.tr;
  for (const assignment of assignments) {
    const { reference } = assignment;
    if (sameNoteAttributes(reference.node, assignment)) continue;
    transaction.setNodeMarkup(reference.position, undefined, {
      ...reference.node.attrs,
      id: assignment.id,
      kind: assignment.kind,
      number: assignment.number,
    });
  }
  for (const definition of definitions) {
    const assignment = definitionUpdates.get(definition.position);
    if (!assignment || sameNoteAttributes(definition.node, assignment)) {
      continue;
    }
    transaction.setNodeMarkup(definition.position, undefined, {
      ...definition.node.attrs,
      id: assignment.id,
      kind: assignment.kind,
      number: assignment.number,
    });
  }

  const deletionCandidates = [
    ...nestedReferences,
    ...references.filter((reference) =>
      deletedReferences.has(reference.position),
    ),
    ...definitions.filter(
      (definition) => !usedDefinitions.has(definition.position),
    ),
  ];
  const deletions = deletionCandidates
    .filter(
      (candidate) =>
        !deletionCandidates.some(
          (container) =>
            container !== candidate &&
            container.position < candidate.position &&
            container.position + container.node.nodeSize >=
              candidate.position + candidate.node.nodeSize,
        ),
    )
    .sort((left, right) => right.position - left.position);
  for (const deletion of deletions) {
    transaction.delete(
      deletion.position,
      deletion.position + deletion.node.nodeSize,
    );
  }

  const definitionType = state.schema.nodes[definitionNodeName];
  const paragraphType = state.schema.nodes.paragraph;
  if (definitionType && paragraphType) {
    for (const insertion of definitionInsertions) {
      const { assignment, source } = insertion;
      const content = source?.content.size
        ? source.content
        : paragraphType.create();
      const definition = definitionType.create(
        {
          ...(source?.attrs ?? {}),
          id: assignment.id,
          kind: assignment.kind,
          number: assignment.number,
        },
        content,
      );
      transaction.insert(
        transaction.mapping.map(insertion.position, 1),
        definition,
      );
    }
  }
  return transaction.docChanged ? transaction : null;
}

function assignNoteReferenceIdentities(
  references: readonly NoteNodeAtPosition[],
  retained: ReadonlySet<number>,
): Map<number, NoteReferenceAssignment> {
  const ordered = [
    ...references.filter((reference) => retained.has(reference.position)),
    ...references.filter((reference) => !retained.has(reference.position)),
  ];
  const keys = new Set<string>();
  const assignments = new Map<number, NoteReferenceAssignment>();
  for (const reference of ordered) {
    let id = reference.id;
    let key = reference.key;
    if (!id || keys.has(key)) {
      do {
        id = createWorkId(reference.kind);
        key = documentNoteKey(reference.kind, id);
      } while (keys.has(key));
    }
    keys.add(key);
    assignments.set(reference.position, {
      reference,
      id,
      kind: reference.kind,
      key,
      number: 1,
    });
  }
  return assignments;
}

function assignNoteNumbers(assignments: NoteReferenceAssignment[]): void {
  const counters: Record<WorkDocumentNoteKind, number> = {
    footnote: 0,
    endnote: 0,
  };
  for (const assignment of assignments) {
    counters[assignment.kind] += 1;
    assignment.number = counters[assignment.kind];
  }
}

function preferredDefinition(
  candidates: readonly NoteNodeAtPosition[],
  used: ReadonlySet<number>,
  retained: ReadonlySet<number>,
  preserveIdentity: boolean,
): NoteNodeAtPosition | null {
  const available = candidates.filter(
    (candidate) => !used.has(candidate.position),
  );
  if (!available.length) return null;
  return preserveIdentity
    ? (available.find((candidate) => retained.has(candidate.position)) ??
        available[0] ??
        null)
    : (available.find((candidate) => !retained.has(candidate.position)) ??
        null);
}

function retainedNoteNodePositions(
  oldState: EditorState,
  newState: EditorState,
  transactions: readonly Transaction[],
  nodeName: string,
): Set<number> {
  const mapping = transactionMapping(transactions);
  const retained = new Set<number>();
  const current = documentNoteNodes(newState.doc, nodeName);
  const currentByKey = groupNoteNodesByKey(current);
  for (const previous of documentNoteNodes(oldState.doc, nodeName)) {
    const mapped = mapping.mapResult(previous.position, 1);
    const node = newState.doc.nodeAt(mapped.pos);
    if (node?.type.name === nodeName && noteNodeKey(node) === previous.key) {
      retained.add(mapped.pos);
      continue;
    }
    const matches = currentByKey.get(previous.key);
    if (matches?.length === 1) retained.add(matches[0]?.position ?? -1);
  }
  retained.delete(-1);
  return retained;
}

function transactionMapping(transactions: readonly Transaction[]): Mapping {
  const mapping = new Mapping();
  for (const transaction of transactions) {
    mapping.appendMapping(transaction.mapping);
  }
  return mapping;
}

function documentNoteNodes(
  document: ProseMirrorNode,
  nodeName: string,
): NoteNodeAtPosition[] {
  const nodes: NoteNodeAtPosition[] = [];
  document.descendants((node, position) => {
    if (node.type.name !== nodeName) return;
    const kind =
      documentNoteKind(stringAttribute(node.attrs.kind)) ?? 'footnote';
    const id = stringAttribute(node.attrs.id);
    nodes.push({
      node,
      position,
      id,
      kind,
      key: id ? documentNoteKey(kind, id) : '',
    });
  });
  return nodes;
}

function groupNoteNodesByKey(
  nodes: readonly NoteNodeAtPosition[],
): Map<string, NoteNodeAtPosition[]> {
  const groups = new Map<string, NoteNodeAtPosition[]>();
  for (const node of nodes) {
    const matches = groups.get(node.key) ?? [];
    matches.push(node);
    groups.set(node.key, matches);
  }
  return groups;
}

function noteNodeKey(node: ProseMirrorNode): string {
  const kind = documentNoteKind(stringAttribute(node.attrs.kind));
  const id = stringAttribute(node.attrs.id);
  return kind && id ? documentNoteKey(kind, id) : '';
}

function positionInsideNode(
  document: ProseMirrorNode,
  position: number,
  nodeName: string,
): boolean {
  const resolved = document.resolve(position);
  for (let depth = resolved.depth; depth > 0; depth -= 1) {
    if (resolved.node(depth).type.name === nodeName) return true;
  }
  return false;
}

function noteDefinitionInsertionPosition(
  document: ProseMirrorNode,
  referencePosition: number,
  kind: WorkDocumentNoteKind,
): number {
  const sections: Array<{ position: number; nodeSize: number }> = [];
  document.descendants((node, position) => {
    if (node.type.name === 'documentSection') {
      sections.push({ position, nodeSize: node.nodeSize });
      return false;
    }
  });
  const target =
    kind === 'endnote'
      ? sections.at(-1)
      : sections.find(
          (section) =>
            referencePosition > section.position &&
            referencePosition < section.position + section.nodeSize,
        );
  return target ? target.position + target.nodeSize - 1 : document.content.size;
}

function sameNoteAttributes(
  node: ProseMirrorNode,
  assignment: NoteReferenceAssignment,
): boolean {
  return (
    stringAttribute(node.attrs.id) === assignment.id &&
    documentNoteKind(stringAttribute(node.attrs.kind)) === assignment.kind &&
    Number(node.attrs.number) === assignment.number
  );
}

function stringAttribute(value: unknown): string {
  return typeof value === 'string' ? value.trim() : '';
}
