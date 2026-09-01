import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { diagnoseDocxOpenTypeTypography } from '../src/internal/features/work/work-docx-opentype-diagnostics';
import { OoxmlPackage } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WORD_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordml';

describe('DOCX OpenType typography diagnostics', () => {
  test('reports body, inherited, and page-chrome native property sets', async () => {
    const { archive, document } = await diagnosticPackage(
      '<w:p><w:r><w:rPr><w14:ligatures w14:val="all"/><w14:numForm w14:val="oldStyle"/></w:rPr><w:t>Body</w:t></w:r></w:p>',
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:docDefaults><w:rPrDefault><w:rPr><w14:numSpacing w14:val="tabular"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>',
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:p><w:r><w:rPr><w14:stylisticSets><w14:styleSet w14:id="4"/></w14:stylisticSets><w14:cntxtAlts w14:val="0"/></w:rPr><w:t>Header</w:t></w:r></w:p></w:hdr>',
    );

    expect(await diagnoseDocxOpenTypeTypography(archive, document)).toEqual([
      expect.objectContaining({
        code: 'docx.typography.opentype',
        severity: 'info',
        message: expect.stringContaining(
          '3 native OpenType run-property set(s)',
        ),
      }),
    ]);
  });

  test('separates malformed, misplaced, and namespace-spoofed properties', async () => {
    const { archive, document } = await diagnosticPackage(
      [
        '<w:p><w:r><w:rPr><w14:ligatures w14:val="standard"/></w:rPr><w:t>Valid</w:t></w:r></w:p>',
        '<w:p><w:r><w:rPr><w14:numForm w14:val="unknown"/><evil:numSpacing evil:val="tabular"/></w:rPr><w:t>Invalid</w:t></w:r></w:p>',
        '<w:p><w14:cntxtAlts w14:val="1"/></w:p>',
      ].join(''),
      '',
      '',
    );

    expect(
      (await diagnoseDocxOpenTypeTypography(archive, document)).map(
        ({ code }) => code,
      ),
    ).toEqual(['docx.typography.opentype', 'docx.typography.opentype.invalid']);
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
    `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:w14="${WORD_2010_NAMESPACE}" xmlns:evil="https://example.test/evil"><w:body>${body}</w:body></w:document>`,
  );
  if (styles) zip.file('word/styles.xml', styles);
  if (header) zip.file('word/header1.xml', header);
  const archive = await OoxmlPackage.load(
    await zip.generateAsync({ type: 'arraybuffer' }),
  );
  return { archive, document: await archive.xml('word/document.xml') };
}
