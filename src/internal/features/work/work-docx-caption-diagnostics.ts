import { docxBookmarkReferenceTarget } from './work-document-bookmark-references';
import { documentCitationTagsFromInstruction } from './work-document-citations';
import { docxDocumentFieldKind } from './work-document-fields';
import {
  docxCaptionBookmark,
  docxCaptionSequenceKind,
} from './work-docx-caption-fields';
import {
  type DocxFieldOccurrence,
  docxFieldOccurrenceIsInlineEditable,
  docxFieldOccurrences,
  hasInvalidDocxFieldStructure,
} from './work-docx-field-instructions';
import { attribute, descendants } from './work-ooxml-package';
import type { WorkCompatibilityIssue } from './work-types';

export interface DocxCaptionDiagnostics {
  issues: WorkCompatibilityIssue[];
  hasUnsupportedFields: boolean;
}

export function diagnoseDocxCaptions(
  document: Document,
): DocxCaptionDiagnostics {
  const fields = docxFieldOccurrences(document);
  const sequences = fields.filter((field) =>
    docxCaptionSequenceKind(field.instruction),
  );
  const bodyFields = fields.filter(
    (field) =>
      field.syntax !== 'orphan' &&
      Boolean(docxDocumentFieldKind(field.instruction)),
  );
  const editableBodyFields = bodyFields.filter(
    docxFieldOccurrenceIsInlineEditable,
  );
  const hasUnsupportedFieldStructure =
    hasInvalidDocxFieldStructure(document) ||
    bodyFields.some((field) => !docxFieldOccurrenceIsInlineEditable(field));
  const bookmarkNames = new Set([
    ...captionBookmarkNames(sequences),
    ...pairedBookmarkNames(document),
  ]);
  return {
    issues: [
      ...(sequences.length
        ? [
            {
              code: 'docx.captions',
              feature: 'Captions and cross-references',
              message:
                'Figure and table SEQ fields, independent numbering, caption bookmarks, and caption REF fields are preserved, editable, and updated in native DOCX output.',
              severity: 'info',
            } satisfies WorkCompatibilityIssue,
          ]
        : []),
      ...(editableBodyFields.length
        ? [
            {
              code: 'docx.fields.body',
              feature: 'Body fields',
              message:
                'PAGE, NUMPAGES, SECTION, SECTIONPAGES, DATE, and TIME fields remain live, editable body fields and update in preview, PDF, and native DOCX output.',
              severity: 'info',
            } satisfies WorkCompatibilityIssue,
          ]
        : []),
      ...(hasUnsupportedFieldStructure
        ? [
            {
              code: 'docx.fields.structure',
              feature: 'Field structure',
              message:
                'Incomplete, nested, cross-paragraph, deleted, or instructionless fields cannot form one editable inline field; their current results are retained as text.',
              severity: 'warning',
            } satisfies WorkCompatibilityIssue,
          ]
        : []),
    ],
    hasUnsupportedFields:
      hasUnsupportedFieldStructure ||
      fields.some(
        (field) => !isSupportedCaptionField(field.instruction, bookmarkNames),
      ),
  };
}

function captionBookmarkNames(fields: DocxFieldOccurrence[]): Set<string> {
  const names = new Set<string>();
  for (const field of fields) {
    const paragraph = closestAncestor(field.start, 'p');
    if (!paragraph) continue;
    const bookmark = docxCaptionBookmark(paragraph, field);
    const name = attribute(bookmark ?? paragraph, 'name')?.trim();
    if (name) names.add(name.toLowerCase());
  }
  return names;
}

function isSupportedCaptionField(
  instruction: string,
  bookmarkNames: Set<string>,
): boolean {
  if (
    documentCitationTagsFromInstruction(instruction).length ||
    /^\s*BIBLIOGRAPHY\b/i.test(instruction)
  ) {
    return true;
  }
  if (docxDocumentFieldKind(instruction)) return true;
  if (docxCaptionSequenceKind(instruction)) return true;
  const target = docxBookmarkReferenceTarget(instruction);
  return Boolean(target && bookmarkNames.has(target.toLowerCase()));
}

function pairedBookmarkNames(document: Document): Set<string> {
  const endIds = new Set(
    descendants(document, 'bookmarkEnd').map(
      (bookmark) => attribute(bookmark, 'id')?.trim() ?? '',
    ),
  );
  return new Set(
    descendants(document, 'bookmarkStart').flatMap((bookmark) => {
      const id = attribute(bookmark, 'id')?.trim() ?? '';
      const name = attribute(bookmark, 'name')?.trim().toLowerCase() ?? '';
      return id && name && endIds.has(id) ? [name] : [];
    }),
  );
}

function closestAncestor(element: Element, localName: string): Element | null {
  let current: Element | null = element;
  while (current) {
    if (current.localName === localName) return current;
    current = current.parentElement;
  }
  return null;
}
