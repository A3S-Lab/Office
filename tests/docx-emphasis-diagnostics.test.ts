import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { diagnoseDocxEmphasisMarks } from '../src/internal/features/work/work-docx-emphasis-diagnostics';
import { OoxmlPackage } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('DOCX emphasis diagnostics', () => {
  test('reports inherited, body, and page-chrome marks', async () => {
    const { archive, document } = await diagnosticPackage(
      '<w:p><w:r><w:rPr><w:em w:val="none"/></w:rPr><w:t>Body</w:t></w:r></w:p>',
      `<w:styles xmlns:w="${WORD_NAMESPACE}"><w:docDefaults><w:rPrDefault><w:rPr><w:em w:val="dot"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
      `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:rPr><w:em w:val="underDot"/></w:rPr><w:t>Header</w:t></w:r></w:p></w:hdr>`,
    );

    expect(await diagnoseDocxEmphasisMarks(archive, document)).toEqual([
      expect.objectContaining({
        code: 'docx.emphasis-mark',
        severity: 'info',
        message: expect.stringContaining('3 native emphasis mark(s)'),
      }),
    ]);
  });

  test('separates malformed and misplaced properties from valid evidence', async () => {
    const { archive, document } = await diagnosticPackage(
      [
        '<w:p><w:r><w:rPr><w:em w:val="circle"/></w:rPr><w:t>Valid</w:t></w:r></w:p>',
        '<w:p><w:r><w:rPr><w:em w:val="triangle"/></w:rPr><w:t>Invalid</w:t></w:r></w:p>',
        '<w:p><w:em w:val="dot"/></w:p>',
      ].join(''),
      '',
      '',
    );

    expect(
      (await diagnoseDocxEmphasisMarks(archive, document)).map(
        ({ code }) => code,
      ),
    ).toEqual(['docx.emphasis-mark', 'docx.emphasis-mark.invalid']);
  });
});

async function diagnosticPackage(
  body: string,
  styles: string,
  header: string,
): Promise<{ archive: OoxmlPackage; document: Document }> {
  const zip = new JSZip();
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
  if (styles) zip.file('word/styles.xml', styles);
  if (header) zip.file('word/header1.xml', header);
  const archive = await OoxmlPackage.load(
    await zip.generateAsync({ type: 'arraybuffer' }),
  );
  return { archive, document: await archive.xml('word/document.xml') };
}
