import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { diagnoseDocxKerning } from '../src/internal/features/work/work-docx-kerning-diagnostics';
import { OoxmlPackage } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('DOCX kerning diagnostics', () => {
  test('reports valid inherited, body, and page-chrome thresholds', async () => {
    const { archive, document } = await diagnosticPackage(
      '<w:p><w:r><w:rPr><w:kern w:val="0"/></w:rPr><w:t>Body</w:t></w:r></w:p>',
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:kern w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>',
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:kern w:val="30"/></w:rPr><w:t>Header</w:t></w:r></w:p></w:hdr>',
    );

    expect(await diagnoseDocxKerning(archive, document)).toEqual([
      expect.objectContaining({
        code: 'docx.kerning',
        severity: 'info',
        message: expect.stringContaining('3 native kerning threshold(s)'),
      }),
    ]);
  });

  test('separates malformed and misplaced properties from valid evidence', async () => {
    const { archive, document } = await diagnosticPackage(
      [
        '<w:p><w:r><w:rPr><w:kern w:val="24"/></w:rPr><w:t>Valid</w:t></w:r></w:p>',
        '<w:p><w:r><w:rPr><w:kern w:val="3278"/></w:rPr><w:t>Invalid</w:t></w:r></w:p>',
        '<w:p><w:kern w:val="24"/></w:p>',
      ].join(''),
      '',
      '',
    );

    expect(
      (await diagnoseDocxKerning(archive, document)).map(({ code }) => code),
    ).toEqual(['docx.kerning', 'docx.kerning.invalid']);
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
