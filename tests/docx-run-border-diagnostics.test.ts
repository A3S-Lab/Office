import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { diagnoseDocxRunBorders } from '../src/internal/features/work/work-docx-run-border-diagnostics';
import { OoxmlPackage } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('DOCX run-border diagnostics', () => {
  test('counts native line borders and explicit resets across Word stories', async () => {
    const { archive, document } = await diagnosticPackage({
      'word/document.xml': documentXml(
        run('<w:bdr w:val="single" w:color="auto" w:sz="4"/>', 'Body'),
      ),
      'word/styles.xml': rootXml(
        'styles',
        '<w:style w:type="character" w:styleId="Frame"><w:rPr><w:bdr w:val="double" w:sz="8"/></w:rPr></w:style>',
      ),
      'word/header1.xml': rootXml(
        'hdr',
        `<w:p>${run('<w:bdr w:val="nil"/>', 'Header')}</w:p>`,
      ),
      'word/footer1.xml': rootXml(
        'ftr',
        `<w:p>${run('<w:bdr w:val="none"/>', 'Footer')}</w:p>`,
      ),
      'word/footnotes.xml': rootXml(
        'footnotes',
        `<w:footnote w:id="1"><w:p>${run('<w:bdr w:val="wave" w:sz="8"/>', 'Footnote')}</w:p></w:footnote>`,
      ),
      'word/endnotes.xml': rootXml(
        'endnotes',
        `<w:endnote w:id="1"><w:p>${run('<w:bdr w:val="dashed" w:sz="8"/>', 'Endnote')}</w:p></w:endnote>`,
      ),
      'word/comments.xml': rootXml(
        'comments',
        `<w:comment w:id="0"><w:p>${run('<w:bdr w:val="dotted" w:sz="6"/>', 'Comment')}</w:p></w:comment>`,
      ),
    });

    expect(await diagnoseDocxRunBorders(archive, document)).toEqual([
      expect.objectContaining({
        code: 'docx.run-borders',
        severity: 'info',
        message: expect.stringContaining(
          '7 native character border(s), including 2 explicit reset(s)',
        ),
      }),
    ]);
  });

  test('reports malformed, misplaced, duplicated, art, and spoofed borders', async () => {
    const { archive, document } = await diagnosticPackage({
      'word/document.xml': documentXml(
        [
          run('<w:bdr w:val="single" w:sz="4"/>', 'Valid'),
          run('<w:bdr w:val="apples"/>', 'Art'),
          run('<w:bdr w:val="single"/><w:bdr w:val="double"/>', 'Duplicate'),
          run('<w:bdr w:val="single" w:extra="1"/>', 'Extra'),
          '<w:p><w:bdr w:val="single"/></w:p>',
          `<w:p xmlns:evil="https://example.test/evil">${run(
            '<evil:bdr evil:val="single"/>',
            'Spoofed',
          )}</w:p>`,
        ].join(''),
      ),
    });

    expect(
      (await diagnoseDocxRunBorders(archive, document)).map(({ code }) => code),
    ).toEqual(['docx.run-borders', 'docx.run-borders.invalid']);
  });
});

async function diagnosticPackage(
  parts: Readonly<Record<string, string>>,
): Promise<{ archive: OoxmlPackage; document: Document }> {
  const zip = new JSZip();
  for (const [path, source] of Object.entries(parts)) zip.file(path, source);
  const archive = await OoxmlPackage.load(
    await zip.generateAsync({ type: 'arraybuffer' }),
  );
  return { archive, document: await archive.xml('word/document.xml') };
}

function documentXml(body: string): string {
  return rootXml('document', `<w:body>${body}</w:body>`);
}

function rootXml(name: string, content: string): string {
  return `<w:${name} xmlns:w="${WORD_NAMESPACE}">${content}</w:${name}>`;
}

function run(properties: string, text: string): string {
  return `<w:r><w:rPr>${properties}</w:rPr><w:t>${text}</w:t></w:r>`;
}
