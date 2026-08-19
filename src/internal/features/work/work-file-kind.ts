import type { WorkArtifactKind } from './work-types';

const WORK_FILE_KIND_BY_EXTENSION = Object.freeze({
  docx: 'document',
  html: 'document',
  htm: 'document',
  txt: 'document',
  md: 'markdown',
  markdown: 'markdown',
  xlsx: 'spreadsheet',
  xls: 'spreadsheet',
  csv: 'spreadsheet',
  ods: 'spreadsheet',
  pptx: 'presentation',
  pdf: 'pdf',
} satisfies Record<string, WorkArtifactKind>);

export const WORK_SUPPORTED_FILE_EXTENSIONS = Object.freeze(
  Object.keys(WORK_FILE_KIND_BY_EXTENSION),
);

export function workFileExtension(fileName: string): string {
  const separator = fileName.lastIndexOf('.');
  if (separator <= 0 || separator === fileName.length - 1) return '';
  return fileName.slice(separator + 1).toLowerCase();
}

export function workKindForFileName(fileName: string): WorkArtifactKind | null {
  const extension = workFileExtension(fileName);
  return (
    WORK_FILE_KIND_BY_EXTENSION[
      extension as keyof typeof WORK_FILE_KIND_BY_EXTENSION
    ] ?? null
  );
}

export function workKindForFile(
  file: Pick<File, 'name'>,
): WorkArtifactKind | null {
  return workKindForFileName(file.name);
}
