import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { diagnoseDocxNotes } from '../src/internal/features/work/work-docx-note-diagnostics';
import { OoxmlPackage } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('DOCX note diagnostics', () => {
  test('accepts matched footnote and endnote identifiers', async () => {
    const { archive, document } = await diagnosticPackage(
      [
        '<w:p><w:r><w:footnoteReference w:id="2"/></w:r>',
        '<w:r><w:endnoteReference w:id="3"/></w:r></w:p>',
      ].join(''),
      [
        '<w:footnote w:type="separator" w:id="-1"><w:p/></w:footnote>',
        '<w:footnote w:id="2"><w:p><w:r><w:t>Footnote</w:t></w:r></w:p></w:footnote>',
      ].join(''),
      '<w:endnote w:id="3"><w:p><w:r><w:t>Endnote</w:t></w:r></w:p></w:endnote>',
    );

    expect(await diagnoseDocxNotes(archive, document)).toEqual([
      expect.objectContaining({ code: 'docx.notes', severity: 'info' }),
    ]);
  });

  test('reports missing, duplicate, unreferenced, and nested note identities', async () => {
    const { archive, document } = await diagnosticPackage(
      [
        '<w:p><w:r><w:footnoteReference w:id="2"/></w:r>',
        '<w:r><w:endnoteReference w:id="3"/></w:r></w:p>',
      ].join(''),
      [
        '<w:footnote w:id="4"><w:p><w:r><w:t>Unused</w:t></w:r></w:p></w:footnote>',
        '<w:footnote w:id="4"><w:p><w:r><w:footnoteReference w:id="4"/></w:r></w:p></w:footnote>',
      ].join(''),
      [
        '<w:endnote w:id="3"><w:p><w:r><w:t>First</w:t></w:r></w:p></w:endnote>',
        '<w:endnote w:id="3"><w:p><w:r><w:t>Duplicate</w:t></w:r></w:p></w:endnote>',
      ].join(''),
    );

    expect(
      (await diagnoseDocxNotes(archive, document)).map(({ code }) => code),
    ).toEqual(['docx.notes', 'docx.notes.structure']);
  });

  test('reports repeated references to one definition', async () => {
    const { archive, document } = await diagnosticPackage(
      [
        '<w:p><w:r><w:footnoteReference w:id="2"/></w:r>',
        '<w:r><w:footnoteReference w:id="2"/></w:r></w:p>',
      ].join(''),
      '<w:footnote w:id="2"><w:p><w:r><w:t>Footnote</w:t></w:r></w:p></w:footnote>',
      '',
    );

    expect(
      (await diagnoseDocxNotes(archive, document)).map(({ code }) => code),
    ).toEqual(['docx.notes', 'docx.notes.structure']);
  });

  test('reports the static versus active content-control boundary', async () => {
    const { archive, document } = await diagnosticPackage(
      '<w:p><w:r><w:footnoteReference w:id="2"/></w:r></w:p>',
      [
        '<w:footnote w:id="2">',
        '<w:p><w:sdt><w:sdtPr><w:tag w:val="field"/></w:sdtPr>',
        '<w:sdtContent><w:r><w:t>Controlled</w:t></w:r></w:sdtContent>',
        '</w:sdt></w:p></w:footnote>',
      ].join(''),
      '',
    );

    expect(
      (await diagnoseDocxNotes(archive, document)).map(({ code }) => code),
    ).toEqual(['docx.notes', 'docx.notes.content-controls']);
  });

  test('treats table-only note content as native and supported', async () => {
    const { archive, document } = await diagnosticPackage(
      '<w:p><w:r><w:footnoteReference w:id="2"/></w:r></w:p>',
      [
        '<w:footnote w:id="2"><w:sdt><w:sdtPr><w:richText/></w:sdtPr>',
        '<w:sdtContent><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
        '</w:sdtContent></w:sdt></w:footnote>',
      ].join(''),
      '',
    );

    expect(
      (await diagnoseDocxNotes(archive, document)).map(({ code }) => code),
    ).toEqual(['docx.notes', 'docx.notes.content-controls']);
  });
});

async function diagnosticPackage(
  body: string,
  footnotes: string,
  endnotes: string,
): Promise<{ archive: OoxmlPackage; document: Document }> {
  const zip = new JSZip();
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
  zip.file(
    'word/footnotes.xml',
    `<w:footnotes xmlns:w="${WORD_NAMESPACE}">${footnotes}</w:footnotes>`,
  );
  zip.file(
    'word/endnotes.xml',
    `<w:endnotes xmlns:w="${WORD_NAMESPACE}">${endnotes}</w:endnotes>`,
  );
  const archive = await OoxmlPackage.load(
    await zip.generateAsync({ type: 'arraybuffer' }),
  );
  return { archive, document: await archive.xml('word/document.xml') };
}
