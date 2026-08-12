import {
  inspectDocxEquation,
  isSupportedDocxEquationPlacement,
} from './work-docx-equation-import';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
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

const WORDPROCESSINGML_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  'http://purl.oclc.org/ooxml/wordprocessingml/main',
]);
const WORDPROCESSING_DRAWING_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  'http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing',
]);
const DRAWINGML_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/drawingml/2006/main',
  'http://purl.oclc.org/ooxml/drawingml/main',
]);
const PICTURE_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/drawingml/2006/picture',
  'http://purl.oclc.org/ooxml/drawingml/picture',
]);
const RELATIONSHIP_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
]);

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
      'Footnote and endnote references, stable native IDs, editable note text, native tables, DrawingML pictures, supported structured OMML equations, common inline formatting, safe web, mail, and internal hyperlinks, eligible static text content controls, preview placement, native DOCX note parts, and eligible passive definition, paragraph, table, drawing, wrapper, and text-stable run metadata are preserved.',
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
        descendants(notes, 'pict').length > 0 ||
        noteHasUnsupportedEquation(notes) ||
        descendants(notes, 'drawing').some(
          (drawing) => !isSupportedNotePicture(drawing),
        ),
    )
  ) {
    issues.push(
      noteIssue(
        'docx.notes.rich-content',
        'Legacy VML pictures, shapes, SmartArt, unsupported or malformed equations, and non-picture drawings inside notes may be flattened or converted to bounded inline content. Native paragraphs, tables, validated DrawingML pictures, and supported structured OMML equations remain editable within the declared static-content boundary.',
      ),
    );
  }
  if (
    noteDocuments.some(
      ({ document: notes }) => descendants(notes, 'sdt').length > 0,
    )
  ) {
    issues.push(
      noteIssue(
        'docx.notes.content-controls',
        'Text-stable static rich-text and plain-text controls retain eligible inline or contiguous block wrappers and metadata. Rich-text block controls may include uniquely matched stable tables and nested tables while generated geometry remains authoritative. Data-bound, placeholder, form, nested, relationship-bound, ambiguous, edited-structure, math, drawing-bearing, or mixed-hyperlink control wrappers normalize instead of being reattached unsafely.',
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

function noteHasUnsupportedEquation(document: Document): boolean {
  return Array.from(document.querySelectorAll('*')).some(
    (element) =>
      (element.localName === 'oMath' || element.localName === 'oMathPara') &&
      (inspectDocxEquation(element).status !== 'supported' ||
        !isSupportedDocxEquationPlacement(element)),
  );
}

function isSupportedNotePicture(drawing: Element): boolean {
  if (!WORDPROCESSINGML_NAMESPACES.has(drawing.namespaceURI ?? '')) {
    return false;
  }
  const containers = directChildren(drawing).filter(
    (element) =>
      (element.localName === 'anchor' || element.localName === 'inline') &&
      WORDPROCESSING_DRAWING_NAMESPACES.has(element.namespaceURI ?? ''),
  );
  if (containers.length !== 1) return false;
  const container = containers[0];
  const graphicData = descendants(container, 'graphicData').filter((element) =>
    DRAWINGML_NAMESPACES.has(element.namespaceURI ?? ''),
  );
  const pictures = descendants(container, 'pic').filter((element) =>
    PICTURE_NAMESPACES.has(element.namespaceURI ?? ''),
  );
  const blips = descendants(container, 'blip').filter((element) =>
    DRAWINGML_NAMESPACES.has(element.namespaceURI ?? ''),
  );
  return (
    graphicData.length === 1 &&
    PICTURE_NAMESPACES.has(graphicData[0].getAttribute('uri') ?? '') &&
    pictures.length === 1 &&
    blips.length === 1 &&
    hasRelationshipAttribute(blips[0], 'embed')
  );
}

function hasRelationshipAttribute(element: Element, name: string): boolean {
  return Array.from(element.attributes).some(
    (item) =>
      xmlAttributeLocalName(item) === name &&
      RELATIONSHIP_NAMESPACES.has(xmlAttributeNamespace(element, item) ?? '') &&
      Boolean(item.value.trim()),
  );
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
