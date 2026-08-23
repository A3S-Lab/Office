import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { diagnoseDocxLegacyTextEffects } from '../src/internal/features/work/work-docx-legacy-text-effects-diagnostics';
import { OoxmlPackage } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('DOCX legacy text-effect diagnostics', () => {
  test('counts each native effect and explicit false reset across Word stories', async () => {
    const { archive, document } = await diagnosticPackage({
      'word/document.xml': documentXml(run('<w:outline/><w:shadow/>', 'Body')),
      'word/styles.xml': rootXml(
        'styles',
        '<w:style w:type="character" w:styleId="Relief"><w:rPr><w:emboss/></w:rPr></w:style>',
      ),
      'word/header1.xml': rootXml(
        'hdr',
        `<w:p>${run('<w:imprint/>', 'Header')}</w:p>`,
      ),
      'word/footer1.xml': rootXml(
        'ftr',
        `<w:p>${run('<w:outline w:val="0"/><w:shadow w:val="false"/>', 'Footer')}</w:p>`,
      ),
      'word/footnotes.xml': rootXml(
        'footnotes',
        `<w:footnote w:id="1"><w:p>${run('<w:emboss w:val="off"/>', 'Footnote')}</w:p></w:footnote>`,
      ),
      'word/endnotes.xml': rootXml(
        'endnotes',
        `<w:endnote w:id="1"><w:p>${run('<w:imprint w:val="0"/>', 'Endnote')}</w:p></w:endnote>`,
      ),
      'word/comments.xml': rootXml(
        'comments',
        `<w:comment w:id="0"><w:p>${run('<w:shadow/>', 'Comment')}</w:p></w:comment>`,
      ),
    });

    expect(await diagnoseDocxLegacyTextEffects(archive, document)).toEqual([
      expect.objectContaining({
        code: 'docx.legacy-text-effects',
        severity: 'info',
        message: expect.stringContaining(
          '9 native legacy text-effect value(s), including 4 explicit false reset(s)',
        ),
      }),
    ]);
  });

  test('reports malformed, conflicting, misplaced, and spoofed properties separately', async () => {
    const { archive, document } = await diagnosticPackage({
      'word/document.xml': documentXml(
        [
          run('<w:outline/>', 'Valid'),
          run('<w:outline/><w:emboss/>', 'Conflict'),
          run('<w:shadow w:val="TRUE"/>', 'Unknown token'),
          run('<w:emboss/><w:emboss w:val="0"/>', 'Duplicate'),
          '<w:p><w:imprint/></w:p>',
          `<w:p xmlns:evil="https://example.test/evil">${run(
            '<evil:outline/>',
            'Spoofed',
          )}</w:p>`,
        ].join(''),
      ),
    });

    expect(
      (await diagnoseDocxLegacyTextEffects(archive, document)).map(
        ({ code }) => code,
      ),
    ).toEqual(['docx.legacy-text-effects', 'docx.legacy-text-effects.invalid']);
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
