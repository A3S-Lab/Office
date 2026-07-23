import { describe, expect, test } from '@rstest/core';
import {
  artifactExtension,
  createArtifact,
  createArtifactBlob,
  defaultPptxRuntimeUrl,
  OFFICE_FILE_ACCEPT,
  officeKindForFile,
} from '../src/core';

describe('office core', () => {
  test('creates complete blank artifacts', () => {
    const document = createArtifact('blank-document');
    const spreadsheet = createArtifact('blank-spreadsheet');
    const presentation = createArtifact('blank-presentation');

    expect(document.kind).toBe('document');
    expect(document.content.type).toBe('document');
    expect(spreadsheet.kind).toBe('spreadsheet');
    expect(spreadsheet.content.type).toBe('spreadsheet');
    expect(presentation.kind).toBe('presentation');
    expect(presentation.content.type).toBe('presentation');
    expect(artifactExtension(presentation.kind)).toBe('pptx');
  });

  test('detects supported files without reading their contents', () => {
    expect(officeKindForFile(new File([], 'proposal.docx'))).toBe('document');
    expect(officeKindForFile(new File([], 'forecast.xlsx'))).toBe(
      'spreadsheet',
    );
    expect(officeKindForFile(new File([], 'deck.pptx'))).toBe('presentation');
    expect(officeKindForFile(new File([], 'contract.pdf'))).toBe('pdf');
    expect(officeKindForFile(new File([], 'archive.zip'))).toBeNull();
    expect(OFFICE_FILE_ACCEPT).toContain('.docx');
    expect(OFFICE_FILE_ACCEPT).toContain('.pdf');
  });

  test('exports a blank spreadsheet as a real workbook blob', async () => {
    const blob = await createArtifactBlob(createArtifact('blank-spreadsheet'));

    expect(blob.type).toBe(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    expect(blob.size).toBeGreaterThan(1_000);
  });

  test('publishes a colocated default PowerPoint runtime URL', () => {
    expect(new URL(defaultPptxRuntimeUrl).pathname).toMatch(
      /\/pptxgen\.bundle\.js$/,
    );
  });
});
