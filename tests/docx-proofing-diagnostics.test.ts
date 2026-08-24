import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { diagnoseDocxProofing } from '../src/internal/features/work/work-docx-proofing-diagnostics';
import { OoxmlPackage } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('DOCX proofing-language diagnostics', () => {
  test('counts native language slots and explicit proofing states across Word stories', async () => {
    const { archive, document } = await diagnosticPackage({
      'word/document.xml': documentXml(
        run(
          '<w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/><w:noProof/>',
          'Body',
        ),
      ),
      'word/styles.xml': rootXml(
        'styles',
        '<w:style w:type="character" w:styleId="Proof"><w:rPr><w:lang w:val="fr-FR"/><w:noProof w:val="0"/></w:rPr></w:style>',
      ),
      'word/header1.xml': rootXml(
        'hdr',
        `<w:p>${run('<w:lang w:val="en-GB"/>', 'Header')}</w:p>`,
      ),
      'word/footer1.xml': rootXml(
        'ftr',
        `<w:p>${run('<w:noProof w:val="false"/>', 'Footer')}</w:p>`,
      ),
      'word/footnotes.xml': rootXml(
        'footnotes',
        `<w:footnote w:id="1"><w:p>${run('<w:lang w:eastAsia="ja-JP"/>', 'Footnote')}</w:p></w:footnote>`,
      ),
      'word/endnotes.xml': rootXml(
        'endnotes',
        `<w:endnote w:id="1"><w:p>${run('<w:lang w:bidi="he-IL"/><w:noProof w:val="1"/>', 'Endnote')}</w:p></w:endnote>`,
      ),
      'word/comments.xml': rootXml(
        'comments',
        `<w:comment w:id="0"><w:p>${run('<w:lang w:val="de-DE"/>', 'Comment')}</w:p></w:comment>`,
      ),
    });

    expect(await diagnoseDocxProofing(archive, document)).toEqual([
      expect.objectContaining({
        code: 'docx.proofing-language',
        severity: 'info',
        message: expect.stringContaining(
          '6 native language declaration(s) and 4 explicit proofing state(s)',
        ),
      }),
    ]);
  });

  test('reports malformed, misplaced, duplicated, and spoofed proofing properties', async () => {
    const { archive, document } = await diagnosticPackage({
      'word/document.xml': documentXml(
        [
          run('<w:lang w:val="en-US"/><w:noProof/>', 'Valid'),
          run('<w:lang w:val="en_US"/>', 'Invalid tag'),
          run('<w:lang w:val="en-US"/><w:lang w:val="fr-FR"/>', 'Duplicate'),
          run('<w:noProof w:val="yes"/>', 'Invalid switch'),
          run('<w:noProof/><w:noProof w:val="0"/>', 'Duplicate switch'),
          '<w:p><w:lang w:val="en-US"/><w:noProof/></w:p>',
          `<w:p xmlns:evil="https://example.test/evil">${run(
            '<evil:lang evil:val="en-US"/><evil:noProof/>',
            'Spoofed',
          )}</w:p>`,
        ].join(''),
      ),
    });

    expect(
      (await diagnoseDocxProofing(archive, document)).map(({ code }) => code),
    ).toEqual(['docx.proofing-language', 'docx.proofing-language.invalid']);
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
