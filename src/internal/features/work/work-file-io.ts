import {
  createWorkDocumentBlob,
  importWorkDocumentFile,
} from './work-document-file-io';
import {
  downloadBlob,
  fileNameWithoutExtension,
  safeFileName,
} from './work-file-download';
import { workFileExtension, workKindForFile } from './work-file-kind';
import { materializeWorkFileSource } from './work-file-data';
import {
  type WorkFileImportContext,
  WorkFileImportController,
  type WorkFileImportOptions,
} from './work-file-import';
import {
  createWorkMarkdownBlob,
  importWorkMarkdownFile,
} from './work-markdown-file-io';
import {
  createWorkPresentationBlob,
  importWorkPresentationFile,
  type WorkPresentationExportOptions,
} from './work-presentation-file-io';
import {
  moveWorkSourceBlob,
  readWorkSourceBlob,
  rememberWorkSourceBlob,
} from './work-repository';
import { createWorkSpreadsheetBlob } from './work-spreadsheet-file-export';
import { importWorkSpreadsheetFile } from './work-spreadsheet-file-import';
import { createWorkArtifact } from './work-templates';
import { type WorkArtifact, workArtifactExtension } from './work-types';

export { WORK_IMPORT_ACCEPT } from './work-file-contract';

export type WorkArtifactExportOptions = WorkPresentationExportOptions;

export async function importWorkFile(
  file: File,
  options: WorkFileImportOptions = {},
): Promise<WorkArtifact> {
  const reservedArtifactId = normalizedReservedArtifactId(options.artifactId);
  const reservedSpreadsheetSheetIds = normalizedReservedSpreadsheetSheetIds(
    options.spreadsheetSheetIds,
  );
  const extension = workFileExtension(file.name);
  const kind = workKindForFile(file);
  if (!kind) {
    throw new Error(
      '目前可导入 DOCX、XLSX、XLS、ODS、CSV、PPTX、PDF、HTML、Markdown 和文本文件。',
    );
  }
  const controller = new WorkFileImportController(options, file.size);
  const source = await materializeWorkFileSource(file, controller);
  const context: WorkFileImportContext = {
    bytes: source.bytes,
    controller,
    ...(reservedSpreadsheetSheetIds
      ? { spreadsheetSheetIds: reservedSpreadsheetSheetIds }
      : {}),
  };
  controller.report('parsing', 0);
  let artifact: WorkArtifact;
  if (kind === 'markdown') {
    artifact = await importWorkMarkdownFile(source.file, context);
  } else if (kind === 'document') {
    artifact = await importWorkDocumentFile(source.file, extension, context);
  } else if (kind === 'spreadsheet') {
    artifact = await importWorkSpreadsheetFile(source.file, extension, context);
  } else if (kind === 'presentation') {
    artifact = await importWorkPresentationFile(source.file, context);
  } else {
    artifact = await importPdf(source.file);
  }
  if (reservedArtifactId && reservedArtifactId !== artifact.id) {
    moveWorkSourceBlob(artifact.id, reservedArtifactId);
    artifact.id = reservedArtifactId;
  }
  controller.report('analyzing', 1);
  controller.report('finalizing', 0);
  controller.complete();
  return artifact;
}

function normalizedReservedArtifactId(value: unknown): string | null {
  if (value === undefined) return null;
  if (typeof value !== 'string' || !value.trim() || value.length > 256) {
    throw new Error('A reserved artifact ID must be 1 to 256 characters.');
  }
  return value;
}

function normalizedReservedSpreadsheetSheetIds(
  value: unknown,
): readonly string[] | null {
  if (value === undefined) return null;
  if (!Array.isArray(value) || value.length > 16_384) {
    throw new Error(
      'Reserved spreadsheet sheet IDs must be an array with at most 16,384 entries.',
    );
  }
  const ids: string[] = [];
  for (const id of value) {
    if (typeof id !== 'string' || !id.trim() || id.length > 256) {
      throw new Error(
        'A reserved spreadsheet sheet ID must be 1 to 256 characters.',
      );
    }
    ids.push(id);
  }
  if (new Set(ids).size !== ids.length) {
    throw new Error('Reserved spreadsheet sheet IDs must be unique.');
  }
  return Object.freeze(ids);
}

export async function exportWorkArtifact(
  artifact: WorkArtifact,
  options?: WorkArtifactExportOptions,
): Promise<void> {
  const blob = await createWorkArtifactBlob(artifact, options);
  downloadBlob(
    blob,
    `${safeFileName(artifact.title)}.${workArtifactExtension(artifact.kind)}`,
  );
}

export async function createWorkArtifactBlob(
  artifact: WorkArtifact,
  options?: WorkArtifactExportOptions,
): Promise<Blob> {
  if (artifact.kind === 'document') return createWorkDocumentBlob(artifact);
  if (artifact.kind === 'markdown') return createWorkMarkdownBlob(artifact);
  if (artifact.kind === 'spreadsheet')
    return createWorkSpreadsheetBlob(artifact);
  if (artifact.kind === 'presentation')
    return createWorkPresentationBlob(artifact, options);
  return readWorkSourceBlob(artifact);
}

async function importPdf(file: File): Promise<WorkArtifact> {
  const contentType = file.type || 'application/pdf';
  const source = new Blob([file], { type: contentType });
  const artifact = createWorkArtifact('blank-document');
  artifact.kind = 'pdf';
  artifact.title = fileNameWithoutExtension(file.name);
  artifact.content = { type: 'pdf' };
  artifact.source = {
    name: file.name,
    contentType,
    size: file.size,
    updatedAt: file.lastModified || Date.now(),
  };
  rememberWorkSourceBlob(artifact.id, source);
  return artifact;
}

export { workKindForFile } from './work-file-kind';
