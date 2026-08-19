import { expect, test } from '@rstest/core';
import { WORK_IMPORT_ACCEPT } from '../src/internal/features/work/work-file-contract';
import {
  WORK_SUPPORTED_FILE_EXTENSIONS,
  workFileExtension,
  workKindForFile,
  workKindForFileName,
} from '../src/internal/features/work/work-file-kind';

const SUPPORTED_FILES = Object.freeze({
  'brief.docx': 'document',
  'index.html': 'document',
  'legacy.htm': 'document',
  'notes.txt': 'document',
  'readme.md': 'markdown',
  'guide.markdown': 'markdown',
  'budget.xlsx': 'spreadsheet',
  'legacy.xls': 'spreadsheet',
  'data.csv': 'spreadsheet',
  'portable.ods': 'spreadsheet',
  'deck.pptx': 'presentation',
  'report.pdf': 'pdf',
} as const);

test('maps every accepted extension through one file-kind contract', () => {
  expect(WORK_SUPPORTED_FILE_EXTENSIONS).toEqual([
    'docx',
    'html',
    'htm',
    'txt',
    'md',
    'markdown',
    'xlsx',
    'xls',
    'csv',
    'ods',
    'pptx',
    'pdf',
  ]);
  expect(WORK_IMPORT_ACCEPT).toBe(
    WORK_SUPPORTED_FILE_EXTENSIONS.map((extension) => `.${extension}`).join(
      ',',
    ),
  );

  for (const [fileName, kind] of Object.entries(SUPPORTED_FILES)) {
    expect(workKindForFileName(fileName)).toBe(kind);
    expect(workKindForFile({ name: fileName.toUpperCase() })).toBe(kind);
  }
});

test('requires a real terminal extension before selecting an editor', () => {
  expect(workFileExtension('archive.backup.PDF')).toBe('pdf');
  expect(workFileExtension('pdf')).toBe('');
  expect(workFileExtension('.pdf')).toBe('');
  expect(workFileExtension('report.')).toBe('');
  expect(workKindForFileName('pdf')).toBeNull();
  expect(workKindForFileName('report.pdf.exe')).toBeNull();
});
