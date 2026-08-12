import {
  attribute,
  descendants,
  directChildren,
  type OoxmlPackage,
} from './work-ooxml-package';
import type { WorkCompatibilityIssue } from './work-types';

interface DocxNotePart {
  kind: 'footnote' | 'endnote';
  reference: 'footnoteReference' | 'endnoteReference';
  document: Document;
}

export async function diagnoseDocxNotes(
  archive: OoxmlPackage,
  document: Document | null,
): Promise<WorkCompatibilityIssue[]> {
  const noteParts = (
    [
      {
        path: 'word/footnotes.xml',
        kind: 'footnote',
        reference: 'footnoteReference',
      },
      {
        path: 'word/endnotes.xml',
        kind: 'endnote',
        reference: 'endnoteReference',
      },
    ] as const
  ).filter(({ path }) => archive.has(path));
  const noteDocuments: DocxNotePart[] = await Promise.all(
    noteParts.map(async ({ path, kind, reference }) => ({
      kind,
      reference,
      document: await archive.xml(path),
    })),
  );
  const hasNotes =
    Boolean(
      document &&
        (descendants(document, 'footnoteReference').length ||
          descendants(document, 'endnoteReference').length),
    ) ||
    noteDocuments.some(
      ({ document: notes, kind }) => normalNoteDefinitions(notes, kind).length,
    );
  if (!hasNotes) return [];

  const issues: WorkCompatibilityIssue[] = [
    noteIssue(
      'docx.notes',
      'Footnote and endnote references, stable native IDs, editable note text, common inline formatting, safe web, mail, and internal hyperlinks, preview placement, native DOCX note parts, and eligible passive definition, paragraph, hyperlink, and text-stable run metadata are preserved.',
      'info',
    ),
  ];
  if (hasInvalidNoteStructure(document, noteDocuments)) {
    issues.push(
      noteIssue(
        'docx.notes.structure',
        'Missing, duplicate, unreferenced, or nested footnote and endnote identities cannot form a unique editable pair; unmatched content is omitted and repeated references receive new identities during editing import.',
      ),
    );
  }
  if (
    noteDocuments.some(
      ({ document: notes }) =>
        descendants(notes, 'tbl').length > 0 ||
        descendants(notes, 'drawing').length > 0 ||
        descendants(notes, 'pict').length > 0,
    )
  ) {
    issues.push(
      noteIssue(
        'docx.notes.rich-content',
        'Tables, drawings, and embedded media inside notes may be flattened or converted to inline content.',
      ),
    );
  }

  const settings = archive.has('word/settings.xml')
    ? await archive.xml('word/settings.xml')
    : null;
  const noteProperties = [
    ...(document ? descendants(document, 'footnotePr') : []),
    ...(document ? descendants(document, 'endnotePr') : []),
    ...(settings ? descendants(settings, 'footnotePr') : []),
    ...(settings ? descendants(settings, 'endnotePr') : []),
  ];
  if (
    noteProperties.some((properties) => directChildren(properties).length > 0)
  ) {
    issues.push(
      noteIssue(
        'docx.notes.numbering',
        'Custom note symbols, numbering formats, restart rules, separators, and placement settings normalize to continuous Arabic numbering with footnotes per page and endnotes at document end.',
      ),
    );
  }
  return issues;
}

function hasInvalidNoteStructure(
  document: Document | null,
  parts: readonly DocxNotePart[],
): boolean {
  for (const part of parts) {
    const definitions = normalNoteDefinitions(part.document, part.kind);
    const definitionsById = groupedById(definitions);
    const references = document ? descendants(document, part.reference) : [];
    const referencesById = groupedById(references);
    const referenceIds = new Set(
      references.flatMap((reference) => {
        const id = attribute(reference, 'id')?.trim() ?? '';
        return id ? [id] : [];
      }),
    );
    if (
      references.some((reference) => !attribute(reference, 'id')?.trim()) ||
      Array.from(referencesById.values()).some(
        (matches) => matches.length > 1,
      ) ||
      references.some((reference) => {
        const id = attribute(reference, 'id')?.trim() ?? '';
        return id && (definitionsById.get(id)?.length ?? 0) !== 1;
      }) ||
      definitions.some((definition) => !attribute(definition, 'id')?.trim()) ||
      Array.from(definitionsById.values()).some(
        (matches) => matches.length > 1,
      ) ||
      Array.from(definitionsById.keys()).some((id) => !referenceIds.has(id)) ||
      descendants(part.document, 'footnoteReference').length > 0 ||
      descendants(part.document, 'endnoteReference').length > 0
    ) {
      return true;
    }
  }
  if (!document) return false;
  return (
    (descendants(document, 'footnoteReference').length > 0 &&
      !parts.some(({ kind }) => kind === 'footnote')) ||
    (descendants(document, 'endnoteReference').length > 0 &&
      !parts.some(({ kind }) => kind === 'endnote'))
  );
}

function normalNoteDefinitions(
  document: Document,
  kind: DocxNotePart['kind'],
): Element[] {
  return descendants(document, kind).filter((note) => {
    const type = attribute(note, 'type')?.trim().toLowerCase() ?? '';
    return !type || type === 'normal';
  });
}

function groupedById(elements: readonly Element[]): Map<string, Element[]> {
  const groups = new Map<string, Element[]>();
  for (const element of elements) {
    const id = attribute(element, 'id')?.trim() ?? '';
    const matches = groups.get(id) ?? [];
    matches.push(element);
    groups.set(id, matches);
  }
  return groups;
}

function noteIssue(
  code: string,
  message: string,
  severity: WorkCompatibilityIssue['severity'] = 'warning',
): WorkCompatibilityIssue {
  return {
    code,
    feature: 'Footnotes and endnotes',
    message,
    severity,
  };
}
