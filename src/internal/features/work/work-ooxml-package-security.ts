export function isSafeOpcPartPath(path: string): boolean {
  if (
    !path ||
    path.startsWith('/') ||
    path.includes('\\') ||
    /[\u0000-\u001f\u007f]/.test(path)
  ) {
    return false;
  }
  const segments = path.split('/');
  return segments.every(
    (segment) => Boolean(segment) && segment !== '.' && segment !== '..',
  );
}

export function isExcludedDocxPart(path: string): boolean {
  const normalized = path.toLowerCase();
  return (
    normalized.startsWith('_xmlsignatures/') ||
    normalized.startsWith('customui/') ||
    normalized.startsWith('word/activex/') ||
    /(^|\/)vbaproject(?:signature)?\.bin(?:\.rels)?$/.test(normalized) ||
    normalized === 'word/vbadata.xml'
  );
}

export function isRegeneratedDocxNoteCommentPart(path: string): boolean {
  return /^word\/(?:footnotes|endnotes|comments|commentsExtended|commentsIds|commentsExtensible|people)\.xml$/i.test(
    path,
  );
}

export function isRegeneratedDocxNoteCommentContentType(
  type: string | undefined,
): boolean {
  const normalized = type?.trim().toLowerCase() ?? '';
  return /^application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.(?:footnotes|endnotes|comments|commentsextended|commentsids|commentsextensible|people)\+xml$/.test(
    normalized,
  );
}

export function isExcludedRelationshipType(type: string): boolean {
  const normalized = type.toLowerCase();
  return [
    'digital-signature',
    'vbaproject',
    'activex',
    'ui/extensibility',
  ].some((marker) => normalized.includes(marker));
}

export function isExcludedDocxContentType(type: string | undefined): boolean {
  const normalized = type?.toLowerCase() ?? '';
  return [
    'digital-signature',
    'vbaproject',
    'activex',
    'customui',
    'custom-ui',
    'ribbon',
  ].some((marker) => normalized.includes(marker));
}
