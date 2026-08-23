import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { diagnoseDocxRunFonts } from '../src/internal/features/work/work-docx-run-fonts-diagnostics';
import { OoxmlPackage } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('DOCX script-font diagnostics', () => {
  test('reports native slots across body, styles, headers, and comments', async () => {
    const { archive, document } = await diagnosticPackage({
      body: '<w:p><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsiTheme="minorHAnsi" w:hint="default"/></w:rPr><w:t>Body</w:t></w:r></w:p>',
      styles:
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:eastAsia="SimSun"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>',
      header:
        '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:rFonts w:cs="Arial" w:cstheme="majorBidi"/></w:rPr><w:t>Header</w:t></w:r></w:p></w:hdr>',
      comments:
        '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0"><w:p><w:r><w:rPr><w:rFonts w:hAnsi="Calibri"/></w:rPr><w:t>Comment</w:t></w:r></w:p></w:comment></w:comments>',
    });

    expect(await diagnoseDocxRunFonts(archive, document)).toEqual([
      expect.objectContaining({
        code: 'docx.script-fonts',
        severity: 'info',
        message: expect.stringContaining('4 native run-font property set(s)'),
      }),
    ]);
  });

  test('separates malformed, namespace-spoofed, and misplaced fonts', async () => {
    const { archive, document } = await diagnosticPackage({
      body: [
        '<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>Valid</w:t></w:r></w:p>',
        '<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:rFonts w:eastAsia="SimSun"/></w:rPr><w:t>Duplicate</w:t></w:r></w:p>',
        '<w:p><w:r><w:rPr><w:rFonts xmlns:evil="urn:spoofed" evil:ascii="Arial"/></w:rPr><w:t>Spoofed</w:t></w:r></w:p>',
        '<w:p><w:rFonts w:ascii="Arial"/></w:p>',
      ].join(''),
    });

    expect(
      (await diagnoseDocxRunFonts(archive, document)).map(({ code }) => code),
    ).toEqual(['docx.script-fonts', 'docx.script-fonts.invalid']);
  });
});

async function diagnosticPackage({
  body,
  styles = '',
  header = '',
  comments = '',
}: {
  body: string;
  styles?: string;
  header?: string;
  comments?: string;
}): Promise<{ archive: OoxmlPackage; document: Document }> {
  const zip = new JSZip();
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
  if (styles) zip.file('word/styles.xml', styles);
  if (header) zip.file('word/header1.xml', header);
  if (comments) zip.file('word/comments.xml', comments);
  const archive = await OoxmlPackage.load(
    await zip.generateAsync({ type: 'arraybuffer' }),
  );
  return { archive, document: await archive.xml('word/document.xml') };
}
