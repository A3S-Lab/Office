import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { diagnoseDocxRunShading } from '../src/internal/features/work/work-docx-run-shading-diagnostics';
import { OoxmlPackage } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('DOCX run-shading diagnostics', () => {
  test('counts native patterns and explicit resets across Word stories', async () => {
    const { archive, document } = await diagnosticPackage({
      'word/document.xml': documentXml(
        run('<w:shd w:val="pct20" w:color="112233" w:fill="DDEEFF"/>', 'Body'),
      ),
      'word/styles.xml': rootXml(
        'styles',
        '<w:style w:type="character" w:styleId="Shade"><w:rPr><w:shd w:val="diagCross" w:color="4472C4"/></w:rPr></w:style>',
      ),
      'word/header1.xml': rootXml(
        'hdr',
        `<w:p>${run('<w:shd w:val="nil"/>', 'Header')}</w:p>`,
      ),
      'word/footer1.xml': rootXml(
        'ftr',
        `<w:p>${run('<w:shd w:val="clear" w:fill="FFF2CC"/>', 'Footer')}</w:p>`,
      ),
      'word/footnotes.xml': rootXml(
        'footnotes',
        `<w:footnote w:id="1"><w:p>${run('<w:shd w:val="solid" w:color="C00000"/>', 'Footnote')}</w:p></w:footnote>`,
      ),
      'word/endnotes.xml': rootXml(
        'endnotes',
        `<w:endnote w:id="1"><w:p>${run('<w:shd w:val="horzStripe"/>', 'Endnote')}</w:p></w:endnote>`,
      ),
      'word/comments.xml': rootXml(
        'comments',
        `<w:comment w:id="0"><w:p>${run('<w:shd w:val="pct75"/>', 'Comment')}</w:p></w:comment>`,
      ),
    });

    expect(await diagnoseDocxRunShading(archive, document)).toEqual([
      expect.objectContaining({
        code: 'docx.run-shading',
        severity: 'info',
        message: expect.stringContaining(
          '7 native character shading value(s), including 1 explicit reset(s)',
        ),
      }),
    ]);
  });

  test('reports malformed, misplaced, duplicated, and spoofed run shading', async () => {
    const { archive, document } = await diagnosticPackage({
      'word/document.xml': documentXml(
        [
          run('<w:shd w:val="clear" w:fill="FFF2CC"/>', 'Valid'),
          run('<w:shd w:val="unknown"/>', 'Unknown'),
          run('<w:shd w:val="clear"/><w:shd w:val="solid"/>', 'Duplicate'),
          run('<w:shd w:val="clear" w:extra="1"/>', 'Extra'),
          '<w:p><w:shd w:val="clear"/></w:p>',
          '<w:p><w:pPr><w:shd w:val="clear" w:fill="DDEEFF"/></w:pPr></w:p>',
          `<w:p xmlns:evil="https://example.test/evil">${run(
            '<evil:shd evil:val="clear"/>',
            'Spoofed',
          )}</w:p>`,
        ].join(''),
      ),
    });

    expect(
      (await diagnoseDocxRunShading(archive, document)).map(({ code }) => code),
    ).toEqual(['docx.run-shading', 'docx.run-shading.invalid']);
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
