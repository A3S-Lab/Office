import {
  normalizeDocumentBookmarkName,
  normalizeDocumentBookmarkNativeId,
} from './work-document-bookmarks';
import { normalizeDocumentHref } from './work-document-links';
import {
  attribute,
  descendants,
  type OoxmlPackage,
} from './work-ooxml-package';
import type { WorkCompatibilityIssue } from './work-types';

export async function diagnoseDocxBookmarksAndLinks(
  archive: OoxmlPackage,
  document: Document,
): Promise<WorkCompatibilityIssue[]> {
  const starts = descendants(document, 'bookmarkStart');
  const ends = descendants(document, 'bookmarkEnd');
  const hyperlinks = descendants(document, 'hyperlink');
  if (!starts.length && !ends.length && !hyperlinks.length) return [];

  const issues: WorkCompatibilityIssue[] = [
    bookmarkIssue(
      'docx.bookmarks-links',
      'Paired body bookmarks remain editable native bookmark ranges. Internal hyperlinks retain Word anchor semantics without external relationships, while supported web and email links retain external hyperlink relationships.',
      'info',
    ),
  ];
  const startsById = groupedByAttribute(starts, 'id');
  const endsById = groupedByAttribute(ends, 'id');
  const boundaryIds = new Set([...startsById.keys(), ...endsById.keys()]);
  if (
    starts.some((start) => !attribute(start, 'id')?.trim()) ||
    ends.some((end) => !attribute(end, 'id')?.trim()) ||
    Array.from(boundaryIds).some(
      (id) =>
        (startsById.get(id)?.length ?? 0) !== (endsById.get(id)?.length ?? 0),
    )
  ) {
    issues.push(
      bookmarkIssue(
        'docx.bookmarks.structure',
        'Unpaired bookmark boundaries cannot form a complete editable range and are omitted during editing import.',
      ),
    );
  }

  const names = new Set<string>();
  const nativeIds = new Set<number>();
  let normalizedIdentity = false;
  for (const start of starts) {
    const sourceName = attribute(start, 'name')?.trim() ?? '';
    const name = normalizeDocumentBookmarkName(sourceName);
    const sourceNativeId = attribute(start, 'id');
    const nativeId = normalizeDocumentBookmarkNativeId(sourceNativeId);
    const foldedName = name?.toLowerCase();
    if (
      !name ||
      !foldedName ||
      names.has(foldedName) ||
      nativeId === null ||
      nativeIds.has(nativeId)
    ) {
      normalizedIdentity = true;
    }
    if (foldedName) names.add(foldedName);
    if (nativeId !== null) nativeIds.add(nativeId);
  }
  if (normalizedIdentity) {
    issues.push(
      bookmarkIssue(
        'docx.bookmarks.identity',
        'Invalid, duplicate, or over-40-character bookmark names are made unique and Word-compatible. Missing or reused numeric identifiers are reassigned, and matching internal links are redirected to the retained target.',
      ),
    );
  }
  if (
    starts.some(
      (start) =>
        attribute(start, 'colFirst') !== null ||
        attribute(start, 'colLast') !== null,
    )
  ) {
    issues.push(
      bookmarkIssue(
        'docx.bookmarks.columns',
        'Column-scoped table bookmarks are imported as ordinary editable bookmark ranges; column boundary metadata is normalized.',
      ),
    );
  }

  const targetNames = new Set(
    starts
      .map((start) => attribute(start, 'name')?.trim().toLowerCase())
      .filter((name): name is string => Boolean(name)),
  );
  const danglingAnchors = hyperlinks.some((hyperlink) => {
    const anchor = attribute(hyperlink, 'anchor')?.trim();
    return Boolean(anchor && !targetNames.has(anchor.toLowerCase()));
  });
  if (danglingAnchors) {
    issues.push(
      bookmarkIssue(
        'docx.hyperlinks.missing-target',
        'One or more internal hyperlinks do not resolve to a body bookmark. Their anchors are retained and shown with an explicit missing-target state.',
      ),
    );
  }

  const relationships = await archive.relationships('word/document.xml');
  const unsupportedExternalLink = hyperlinks.some((hyperlink) => {
    const relationshipId = attribute(hyperlink, 'r:id')?.trim();
    if (!relationshipId) return false;
    const relationship = relationships.get(relationshipId);
    if (!relationship) return true;
    return (
      !relationship.type.endsWith('/hyperlink') ||
      relationship.targetMode?.toLowerCase() !== 'external' ||
      normalizeDocumentHref(relationship.target) === null
    );
  });
  if (unsupportedExternalLink) {
    issues.push(
      bookmarkIssue(
        'docx.hyperlinks.external',
        'Hyperlink relationships outside http, https, and mailto, or unresolved hyperlink relationships, are converted to plain linked text.',
      ),
    );
  }
  if (
    hyperlinks.some(
      (hyperlink) =>
        attribute(hyperlink, 'tooltip') !== null ||
        attribute(hyperlink, 'tgtFrame') !== null ||
        attribute(hyperlink, 'docLocation') !== null ||
        (attribute(hyperlink, 'anchor') !== null &&
          attribute(hyperlink, 'r:id') !== null),
    )
  ) {
    issues.push(
      bookmarkIssue(
        'docx.hyperlinks.metadata',
        'Hyperlink tooltips, target frames, external document locations, and combined external-plus-anchor destinations normalize to the supported link destination.',
      ),
    );
  }
  return issues;
}

function groupedByAttribute(
  elements: readonly Element[],
  name: string,
): Map<string, Element[]> {
  const groups = new Map<string, Element[]>();
  for (const element of elements) {
    const value = attribute(element, name)?.trim() ?? '';
    const matches = groups.get(value) ?? [];
    matches.push(element);
    groups.set(value, matches);
  }
  return groups;
}

function bookmarkIssue(
  code: string,
  message: string,
  severity: WorkCompatibilityIssue['severity'] = 'warning',
): WorkCompatibilityIssue {
  return {
    code,
    feature: 'Bookmarks and hyperlinks',
    message,
    severity,
  };
}
