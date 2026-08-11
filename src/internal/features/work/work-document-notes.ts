export type WorkDocumentNoteKind = 'footnote' | 'endnote';

export interface WorkDocumentNote {
  id: string;
  kind: WorkDocumentNoteKind;
  number: number;
  html: string;
}

export interface WorkDocumentNoteCollection {
  html: string;
  notes: WorkDocumentNote[];
}

const NOTE_REFERENCE_SELECTOR = 'sup[data-document-note-reference]';
const NOTE_DEFINITION_SELECTOR = 'aside[data-document-note]';

export function normalizeDocumentNotesHtml(source: string): string {
  const document = new DOMParser().parseFromString(source, 'text/html');
  const definitions = new Map<string, HTMLElement[]>();
  const definitionTemplates = new Map<string, HTMLElement>();
  for (const element of Array.from(
    document.body.querySelectorAll<HTMLElement>(NOTE_DEFINITION_SELECTOR),
  )) {
    const kind = documentNoteKind(element.dataset.noteKind);
    const id = element.dataset.noteId?.trim();
    if (!kind || !id) {
      element.remove();
      continue;
    }
    const key = documentNoteKey(kind, id);
    const matches = definitions.get(key) ?? [];
    matches.push(element);
    definitions.set(key, matches);
    if (!definitionTemplates.has(key)) definitionTemplates.set(key, element);
  }

  const counters: Record<WorkDocumentNoteKind, number> = {
    footnote: 0,
    endnote: 0,
  };
  const allReferenceElements = Array.from(
    document.body.querySelectorAll<HTMLElement>(NOTE_REFERENCE_SELECTOR),
  );
  const referenceElements = allReferenceElements.filter(
    (element) => !element.closest(NOTE_DEFINITION_SELECTOR),
  );
  const referenceSet = new Set(referenceElements);
  for (const element of allReferenceElements) {
    if (!referenceSet.has(element)) element.remove();
  }
  const reservedKeys = new Set(definitions.keys());
  for (const element of referenceElements) {
    const kind = documentNoteKind(element.dataset.noteKind) ?? 'footnote';
    const id = element.dataset.noteId?.trim();
    if (id) reservedKeys.add(documentNoteKey(kind, id));
  }

  const assignedKeys = new Set<string>();
  const usedDefinitions = new Set<HTMLElement>();
  const assignments: Array<{
    reference: HTMLElement;
    definition: HTMLElement | null;
    template: HTMLElement | null;
    kind: WorkDocumentNoteKind;
    id: string;
    number: number;
  }> = [];
  for (const element of referenceElements) {
    const kind = documentNoteKind(element.dataset.noteKind) ?? 'footnote';
    const sourceId = element.dataset.noteId?.trim() ?? '';
    const sourceKey = sourceId ? documentNoteKey(kind, sourceId) : '';
    let id = sourceId;
    let key = sourceKey;
    if (!id || assignedKeys.has(key)) {
      id = nextDocumentNoteId(kind, counters[kind] + 1, reservedKeys);
      key = documentNoteKey(kind, id);
      reservedKeys.add(key);
    }
    assignedKeys.add(key);
    counters[kind] += 1;
    const number = counters[kind];
    const definition =
      (definitions.get(sourceKey) ?? []).find(
        (candidate) => !usedDefinitions.has(candidate),
      ) ?? null;
    if (definition) usedDefinitions.add(definition);
    applyDocumentNoteAttributes(element, kind, id, number, true);
    element.textContent = String(number);
    assignments.push({
      reference: element,
      definition,
      template: definitionTemplates.get(sourceKey) ?? null,
      kind,
      id,
      number,
    });
  }

  for (const matches of definitions.values()) {
    for (const element of matches) {
      if (!usedDefinitions.has(element)) element.remove();
    }
  }

  for (const assignment of assignments) {
    const { reference, template, kind, id, number } = assignment;
    let definition = assignment.definition;
    if (!definition) {
      definition = template
        ? (template.cloneNode(true) as HTMLElement)
        : document.createElement('aside');
    }
    applyDocumentNoteAttributes(definition, kind, id, number, false);
    ensureDocumentNoteBlocks(definition);
    if (assignment.definition) continue;
    const target =
      kind === 'footnote'
        ? reference?.closest('section[data-document-section]')
        : document.body.querySelector(
            'section[data-document-section]:last-of-type',
          );
    (target ?? document.body).append(definition);
  }
  return document.body.innerHTML;
}

export function collectDocumentNotes(
  source: string,
): WorkDocumentNoteCollection {
  const html = normalizeDocumentNotesHtml(source);
  const document = new DOMParser().parseFromString(html, 'text/html');
  const definitions = new Map<string, HTMLElement>();
  for (const element of Array.from(
    document.body.querySelectorAll<HTMLElement>(NOTE_DEFINITION_SELECTOR),
  )) {
    const kind = documentNoteKind(element.dataset.noteKind);
    const id = element.dataset.noteId?.trim();
    if (kind && id) definitions.set(documentNoteKey(kind, id), element);
  }
  const seen = new Set<string>();
  const notes: WorkDocumentNote[] = [];
  for (const reference of Array.from(
    document.body.querySelectorAll<HTMLElement>(NOTE_REFERENCE_SELECTOR),
  )) {
    const kind = documentNoteKind(reference.dataset.noteKind);
    const id = reference.dataset.noteId?.trim();
    if (!kind || !id) continue;
    const key = documentNoteKey(kind, id);
    if (seen.has(key)) continue;
    const definition = definitions.get(key);
    if (!definition) continue;
    seen.add(key);
    notes.push({
      id,
      kind,
      number: positiveInteger(reference.dataset.noteNumber, notes.length + 1),
      html: definition.innerHTML || '<p></p>',
    });
  }
  return { html, notes };
}

export function documentNoteReferenceKeys(source: string): string[] {
  const document = new DOMParser().parseFromString(source, 'text/html');
  const keys: string[] = [];
  for (const reference of Array.from(
    document.body.querySelectorAll<HTMLElement>(NOTE_REFERENCE_SELECTOR),
  )) {
    const kind = documentNoteKind(reference.dataset.noteKind);
    const id = reference.dataset.noteId?.trim();
    if (kind && id) keys.push(documentNoteKey(kind, id));
  }
  return keys;
}

export function removeDocumentNoteDefinitions(source: string): string {
  const document = new DOMParser().parseFromString(source, 'text/html');
  for (const note of Array.from(
    document.body.querySelectorAll(NOTE_DEFINITION_SELECTOR),
  ))
    note.remove();
  return document.body.innerHTML;
}

export function createDocumentNoteElement(
  document: Document,
  note: WorkDocumentNote,
): HTMLElement {
  const element = document.createElement('aside');
  applyDocumentNoteAttributes(element, note.kind, note.id, note.number, false);
  element.innerHTML = note.html || '<p></p>';
  ensureDocumentNoteBlocks(element);
  return element;
}

export function documentNoteKey(
  kind: WorkDocumentNoteKind,
  id: string,
): string {
  return `${kind}:${id}`;
}

export function documentNoteKind(
  value: string | undefined,
): WorkDocumentNoteKind | null {
  if (value === 'footnote' || value === 'endnote') return value;
  return null;
}

function applyDocumentNoteAttributes(
  element: HTMLElement,
  kind: WorkDocumentNoteKind,
  id: string,
  number: number,
  reference: boolean,
) {
  element.setAttribute(
    reference ? 'data-document-note-reference' : 'data-document-note',
    'true',
  );
  element.setAttribute('data-note-kind', kind);
  element.setAttribute('data-note-id', id);
  element.setAttribute('data-note-number', String(number));
}

function nextDocumentNoteId(
  kind: WorkDocumentNoteKind,
  seed: number,
  existing: ReadonlySet<string>,
): string {
  let suffix = seed;
  while (existing.has(documentNoteKey(kind, `document-${kind}-${suffix}`)))
    suffix += 1;
  return `document-${kind}-${suffix}`;
}

function ensureDocumentNoteBlocks(element: HTMLElement) {
  if (element.children.length) return;
  const text = element.textContent ?? '';
  element.replaceChildren();
  const paragraph = element.ownerDocument.createElement('p');
  paragraph.textContent = text;
  element.append(paragraph);
}

function positiveInteger(value: string | undefined, fallback: number): number {
  const number = Number(value);
  return Number.isSafeInteger(number) && number > 0 ? number : fallback;
}
