import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import {
  artifactExtension,
  createArtifact,
  createArtifactBlob,
  defaultPptxRuntimeUrl,
  importOfficeFile,
  OFFICE_FILE_ACCEPT,
  officeKindForFile,
} from '../src/core';

describe('office core', () => {
  test('creates complete blank artifacts', () => {
    const document = createArtifact('blank-document');
    const markdown = createArtifact('blank-markdown');
    const spreadsheet = createArtifact('blank-spreadsheet');
    const presentation = createArtifact('blank-presentation');

    expect(document.kind).toBe('document');
    expect(document.content).toEqual({
      type: 'document',
      pageSize: 'a4',
      html: '<p></p>',
    });
    expect(markdown.kind).toBe('markdown');
    expect(markdown.content).toEqual({ type: 'markdown', markdown: '' });
    expect(spreadsheet.kind).toBe('spreadsheet');
    expect(spreadsheet.content.type).toBe('spreadsheet');
    expect(presentation.kind).toBe('presentation');
    expect(presentation.content.type).toBe('presentation');
    expect(artifactExtension(presentation.kind)).toBe('pptx');
  });

  test('detects supported files without reading their contents', () => {
    expect(officeKindForFile(new File([], 'proposal.docx'))).toBe('document');
    expect(officeKindForFile(new File([], 'readme.md'))).toBe('markdown');
    expect(officeKindForFile(new File([], 'notes.markdown'))).toBe('markdown');
    expect(officeKindForFile(new File([], 'forecast.xlsx'))).toBe(
      'spreadsheet',
    );
    expect(officeKindForFile(new File([], 'deck.pptx'))).toBe('presentation');
    expect(officeKindForFile(new File([], 'contract.pdf'))).toBe('pdf');
    expect(officeKindForFile(new File([], 'archive.zip'))).toBeNull();
    expect(OFFICE_FILE_ACCEPT).toContain('.docx');
    expect(OFFICE_FILE_ACCEPT).toContain('.md');
    expect(OFFICE_FILE_ACCEPT).toContain('.pdf');
  });

  test('imports and exports Markdown without converting its source', async () => {
    const source = '# Release notes\n\n- Fast\n- Controlled\n';
    const artifact = await importOfficeFile(
      new File([source], 'release-notes.md', { type: 'text/markdown' }),
    );

    expect(artifact.kind).toBe('markdown');
    expect(artifact.title).toBe('release-notes');
    expect(artifact.content).toEqual({ type: 'markdown', markdown: source });
    expect(artifactExtension(artifact.kind)).toBe('md');

    const output = await createArtifactBlob(artifact);
    expect(output.type).toContain('text/markdown');
    expect(await output.text()).toBe(source);
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

  test('keeps document typography and paragraph layout in DOCX export', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document')
      throw new Error('Expected a document artifact.');
    artifact.content.html = [
      '<p style="text-align: justify; line-height: 1.5; margin-left: 24px;"',
      ' data-office-indent-level="1">',
      '<span style="font-family: Arial; font-size: 12pt;">A3S Office</span>',
      '</p>',
    ].join('');

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = await archive.file('word/document.xml')?.async('string');

    expect(xml).toBeDefined();
    expect(xml).toContain('<w:rFonts w:ascii="Arial"');
    expect(xml).toContain('w:hAnsi="Arial"');
    expect(xml).toContain('<w:sz w:val="24"/>');
    expect(xml).toContain('<w:jc w:val="both"/>');
    expect(xml).toContain('<w:spacing w:after="120" w:line="360"/>');
    expect(xml).toContain('<w:ind w:left="360"/>');
  });
});
