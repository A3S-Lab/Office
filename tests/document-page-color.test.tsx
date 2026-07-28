import { describe, expect, test } from '@rstest/core';
import { render, screen, within } from '@testing-library/react';
import JSZip from 'jszip';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import { WorkDocumentPdfPages } from '../src/internal/features/work/components/work-document-pages';
import {
  documentPageColor,
  normalizeDocumentPageColor,
} from '../src/internal/features/work/work-document-page-color';

describe('document page color', () => {
  test('normalizes supported colors and falls back to a white page', () => {
    expect(normalizeDocumentPageColor('#FC0')).toBe('#ffcc00');
    expect(normalizeDocumentPageColor('  #D9EAF7 ')).toBe('#d9eaf7');
    expect(normalizeDocumentPageColor('red')).toBeUndefined();
    expect(documentPageColor(undefined)).toBe('#ffffff');
  });

  test('round-trips the page color through DOCX', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document')
      throw new Error('Expected a document artifact.');
    artifact.content.pageColor = '#d9eaf7';
    artifact.content.html = '<p>Colored paper</p>';

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const documentXml = await archive.file('word/document.xml')?.async('text');
    const settingsXml = await archive.file('word/settings.xml')?.async('text');
    expect(documentXml).toMatch(
      /<w:background\b[^>]*\bw:color="D9EAF7"[^>]*\/?\s*>/,
    );
    expect(settingsXml).toMatch(/<w:displayBackgroundShape\b/);

    const imported = await importOfficeFile(
      new File([blob], 'colored-paper.docx', { type: blob.type }),
    );
    expect(imported.content.type).toBe('document');
    if (imported.content.type !== 'document') return;
    expect(imported.content.pageColor).toBe('#d9eaf7');
  });

  test('uses the page color on the PDF export surface', () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document')
      throw new Error('Expected a document artifact.');
    artifact.content.pageColor = '#fff2cc';

    render(<WorkDocumentPdfPages content={artifact.content} />);

    const page = screen.getByRole('region', {
      name: '文字打印预览第 1 页',
    });
    expect(page).toHaveStyle({ backgroundColor: '#fff2cc' });
    expect(page.querySelector(':scope > header')).toBeNull();
  });

  test('renders explicit PDF page chrome without inventing a filename header', () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document')
      throw new Error('Expected a document artifact.');
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: '<p>Quarterly report</p>',
        footerHtml: '<p>Internal</p>',
        showPageNumber: true,
      },
      first: { headerHtml: '', footerHtml: '', showPageNumber: false },
      even: { headerHtml: '', footerHtml: '', showPageNumber: false },
    };

    render(<WorkDocumentPdfPages content={artifact.content} />);

    const page = screen.getByRole('region', {
      name: '文字打印预览第 1 页',
    });
    expect(page.querySelector(':scope > header')).toHaveTextContent(
      'Quarterly report',
    );
    expect(page.querySelector(':scope > footer')).toHaveTextContent(
      'Internal1',
    );
    expect(within(page).queryByText(artifact.title)).toBeNull();
  });
});
