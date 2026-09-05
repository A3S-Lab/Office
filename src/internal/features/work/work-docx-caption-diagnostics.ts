import { docxBookmarkReferenceTarget } from './work-document-bookmark-references';
import { documentCitationTagsFromInstruction } from './work-document-citations';
import {
  docxDocumentFieldKind,
  docxDocumentFieldTarget,
  supportedDocxDocumentFieldInstruction,
} from './work-document-fields';
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
import {
  supportedDocxIndexEntryField,
  supportedDocxIndexField,
} from './work-docx-index-import';
import { supportedDocxTableOfContentsField } from './work-docx-table-of-contents-import';
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
  const tableOfContentsFields = fields.filter(
    supportedDocxTableOfContentsField,
  );
  const indexEntryFields = fields.filter(supportedDocxIndexEntryField);
  const indexFields = fields.filter(supportedDocxIndexField);
  const tableOfContentsContainers = new Set(
    tableOfContentsFields.flatMap((field) => {
      const container = closestAncestor(field.start, 'sdt');
      return container ? [container] : [];
    }),
  );
  const isTableOfContentsField = (field: DocxFieldOccurrence) => {
    const container = closestAncestor(field.start, 'sdt');
    return Boolean(container && tableOfContentsContainers.has(container));
  };
  const indexContainers = new Set(
    indexFields.flatMap((field) => {
      const container = closestAncestor(field.start, 'sdt');
      return container ? [container] : [];
    }),
  );
  const isIndexField = (field: DocxFieldOccurrence) => {
    const container = closestAncestor(field.start, 'sdt');
    return Boolean(container && indexContainers.has(container));
  };
  const isSupportedBlockField = (field: DocxFieldOccurrence) =>
    isTableOfContentsField(field) || isIndexField(field);
  const sequences = fields.filter((field) =>
    docxCaptionSequenceKind(field.instruction),
  );
  const bookmarkNames = new Set([
    ...captionBookmarkNames(sequences),
    ...pairedBookmarkNames(document),
  ]);
  const bodyFields = fields.filter(
    (field) =>
      field.syntax !== 'orphan' &&
      Boolean(docxDocumentFieldKind(field.instruction)),
  );
  const supportedBodyFields = bodyFields.filter((field) =>
    supportedBodyField(field, bookmarkNames),
  );
  const unsupportedBodyFields = bodyFields.filter(
    (field) => !supportedBodyField(field, bookmarkNames),
  );
  const editableBodyFields = supportedBodyFields.filter(
    docxFieldOccurrenceIsInlineEditable,
  );
  const hasUnsupportedFieldStructure =
    hasInvalidDocxFieldStructure(document, isSupportedBlockField) ||
    bodyFields.some((field) => !docxFieldOccurrenceIsInlineEditable(field));
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
                'PAGE, NUMPAGES, SECTION, SECTIONPAGES, NUMWORDS, NUMCHARS, DATE, TIME, and bookmark-backed PAGEREF fields remain live, editable body fields and update in preview, PDF, and native DOCX output; common numeric switches (Arabic, ROMAN, alphabetic, and Ordinal) are retained.',
              severity: 'info',
            } satisfies WorkCompatibilityIssue,
          ]
        : []),
      ...(unsupportedBodyFields.length
        ? [
            {
              code: 'docx.fields.instructions',
              feature: 'Field instructions',
              message:
                'Unsupported field switches or PAGEREF targets remain at their cached displayed value; only the bounded common instruction subset is made editable.',
              severity: 'warning',
            } satisfies WorkCompatibilityIssue,
          ]
        : []),
      ...(tableOfContentsFields.length
        ? [
            {
              code: 'docx.tableOfContents',
              feature: 'Table of contents',
              message:
                'Heading ranges, hyperlinks, page-number visibility and alignment, leader styles, cached entries, and refreshable native TOC fields remain editable and round-trip in DOCX output.',
              severity: 'info',
            } satisfies WorkCompatibilityIssue,
          ]
        : []),
      ...(indexEntryFields.length || indexFields.length
        ? [
            {
              code: 'docx.index',
              feature: 'Index',
              message:
                'Native XE entries, primary and secondary terms, cross-references, page emphasis, cached INDEX rows, columns, and leader styles remain editable and round-trip in DOCX output.',
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
        (field) =>
          !isSupportedBlockField(field) &&
          !supportedDocxIndexEntryField(field) &&
          !isSupportedCaptionField(field.instruction, bookmarkNames),
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
  if (docxDocumentFieldKind(instruction)) {
    return (
      supportedDocxDocumentFieldInstruction(instruction) &&
      (docxDocumentFieldKind(instruction) !== 'pageReference' ||
        Boolean(
          docxDocumentFieldTarget(instruction) &&
            bookmarkNames.has(
              docxDocumentFieldTarget(instruction)?.toLowerCase() ?? '',
            ),
        ))
    );
  }
  if (docxCaptionSequenceKind(instruction)) return true;
  const target = docxBookmarkReferenceTarget(instruction);
  return Boolean(target && bookmarkNames.has(target.toLowerCase()));
}

function supportedBodyField(
  field: DocxFieldOccurrence,
  bookmarkNames: Set<string>,
): boolean {
  if (!supportedDocxDocumentFieldInstruction(field.instruction)) return false;
  const kind = docxDocumentFieldKind(field.instruction);
  return (
    kind !== 'pageReference' ||
    Boolean(
      docxDocumentFieldTarget(field.instruction) &&
        bookmarkNames.has(
          docxDocumentFieldTarget(field.instruction)?.toLowerCase() ?? '',
        ),
    )
  );
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
