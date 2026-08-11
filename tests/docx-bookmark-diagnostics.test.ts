import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { diagnoseDocxBookmarksAndLinks } from '../src/internal/features/work/work-docx-bookmark-diagnostics';
import { OoxmlPackage } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const HYPERLINK_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

describe('DOCX bookmark and hyperlink diagnostics', () => {
  test('reports native bookmark and supported link preservation without warnings', async () => {
    const { archive, document } = await diagnosticPackage(
      [
        '<w:p><w:bookmarkStart w:id="7" w:name="Target"/>',
        '<w:r><w:t>Target</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>',
        '<w:p><w:hyperlink w:anchor="Target"><w:r><w:t>Jump</w:t></w:r></w:hyperlink>',
        '<w:hyperlink r:id="rId7"><w:r><w:t>Site</w:t></w:r></w:hyperlink></w:p>',
      ].join(''),
      relationship('rId7', 'https://a3s.dev/office'),
    );

    expect(await diagnoseDocxBookmarksAndLinks(archive, document)).toEqual([
      expect.objectContaining({
        code: 'docx.bookmarks-links',
        severity: 'info',
      }),
    ]);
  });

  test('separates structural, identity, range, target, destination, and metadata warnings', async () => {
    const { archive, document } = await diagnosticPackage(
      [
        '<w:p><w:bookmarkStart w:id="1" w:name="Bad name" w:colFirst="0" w:colLast="1"/>',
        '<w:r><w:t>Broken</w:t></w:r></w:p>',
        '<w:p><w:bookmarkStart w:id="2" w:name="Target"/>',
        '<w:r><w:t>First</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>',
        '<w:p><w:bookmarkStart w:id="3" w:name="target"/>',
        '<w:r><w:t>Second</w:t></w:r><w:bookmarkEnd w:id="3"/></w:p>',
        '<w:p><w:hyperlink w:anchor="Missing"><w:r><w:t>Jump</w:t></w:r></w:hyperlink>',
        '<w:hyperlink r:id="rId8" w:tooltip="Legacy tip"><w:r><w:t>FTP</w:t></w:r></w:hyperlink></w:p>',
      ].join(''),
      relationship('rId8', 'ftp://example.test/file'),
    );

    const codes = (await diagnoseDocxBookmarksAndLinks(archive, document)).map(
      ({ code }) => code,
    );
    expect(codes).toEqual([
      'docx.bookmarks-links',
      'docx.bookmarks.structure',
      'docx.bookmarks.identity',
      'docx.bookmarks.columns',
      'docx.hyperlinks.missing-target',
      'docx.hyperlinks.external',
      'docx.hyperlinks.metadata',
    ]);
  });
});

async function diagnosticPackage(
  body: string,
  relationships: string,
): Promise<{ archive: OoxmlPackage; document: Document }> {
  const zip = new JSZip();
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
  zip.file(
    'word/_rels/document.xml.rels',
    `<Relationships xmlns="${PACKAGE_RELATIONSHIP_NAMESPACE}">${relationships}</Relationships>`,
  );
  const archive = await OoxmlPackage.load(
    await zip.generateAsync({ type: 'arraybuffer' }),
  );
  return { archive, document: await archive.xml('word/document.xml') };
}

function relationship(id: string, target: string): string {
  return `<Relationship Id="${id}" Type="${HYPERLINK_RELATIONSHIP}" Target="${target}" TargetMode="External"/>`;
}
