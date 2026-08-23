import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { diagnoseDocxHiddenText } from '../src/internal/features/work/work-docx-hidden-text-diagnostics';
import { OoxmlPackage } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('DOCX hidden-text diagnostics', () => {
  test('reports hidden values and explicit visible resets across every Word story and styles', async () => {
    const { archive, document } = await diagnosticPackage({
      'word/document.xml': documentXml(run('<w:vanish/>', 'Body')),
      'word/styles.xml': rootXml(
        'styles',
        `<w:style w:type="character" w:styleId="Visible"><w:rPr><w:vanish w:val="0"/></w:rPr></w:style>`,
      ),
      'word/header1.xml': rootXml(
        'hdr',
        `<w:p>${run('<w:vanish w:val="on"/>', 'Header')}</w:p>`,
      ),
      'word/footer1.xml': rootXml(
        'ftr',
        `<w:p>${run('<w:vanish w:val="off"/>', 'Footer')}</w:p>`,
      ),
      'word/footnotes.xml': rootXml(
        'footnotes',
        `<w:footnote w:id="1"><w:p>${run('<w:vanish w:val="true"/>', 'Footnote')}</w:p></w:footnote>`,
      ),
      'word/endnotes.xml': rootXml(
        'endnotes',
        `<w:endnote w:id="1"><w:p>${run('<w:vanish w:val="false"/>', 'Endnote')}</w:p></w:endnote>`,
      ),
      'word/comments.xml': rootXml(
        'comments',
        `<w:comment w:id="0"><w:p>${run('<w:vanish w:val="1"/>', 'Comment')}</w:p></w:comment>`,
      ),
    });

    expect(await diagnoseDocxHiddenText(archive, document)).toEqual([
      expect.objectContaining({
        code: 'docx.hidden-text',
        severity: 'info',
        message: expect.stringContaining(
          '7 native hidden-text value(s), including 4 hidden value(s) and 3 explicit visible reset(s)',
        ),
      }),
    ]);
    expect(
      (await diagnoseDocxHiddenText(archive, document))[0]?.message,
    ).toContain(
      'unchanged comment XML remains source-preserved rather than becoming a rich comment-editing claim',
    );
  });

  test('separates malformed, duplicated, misplaced, and namespace-spoofed values', async () => {
    const { archive, document } = await diagnosticPackage({
      'word/document.xml': documentXml(
        [
          run('<w:vanish/>', 'Valid'),
          run('<w:vanish w:val="TRUE"/>', 'Unknown token'),
          run('<w:vanish/><w:vanish w:val="0"/>', 'Duplicate'),
          '<w:p><w:vanish/></w:p>',
          `<w:p xmlns:evil="https://example.test/evil">${run(
            '<evil:vanish/>',
            'Spoofed',
          )}</w:p>`,
        ].join(''),
      ),
    });

    expect(
      (await diagnoseDocxHiddenText(archive, document)).map(({ code }) => code),
    ).toEqual(['docx.hidden-text', 'docx.hidden-text.invalid']);
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
