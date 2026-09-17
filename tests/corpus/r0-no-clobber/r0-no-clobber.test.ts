import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import {
  createArtifactBlob,
  forgetSourceBlob,
  importOfficeFile,
} from '../../../src/core';
import {
  buildActiveContentFailClosedFixture,
  buildBookmarksAndLinksFixture,
  buildContractTableFixture,
  buildDuplicateBookmarkFixture,
  buildEndnoteFixture,
  buildFootnoteFixture,
  buildHeaderFooterFixture,
  buildInternalBookmarkLinkFixture,
  buildReviewCommentsFixture,
  buildReviewTrackChangesFixture,
  buildTextContentControlFixture,
} from './fixtures';
import {
  expectIssueCodes,
  extractDocumentIdentities,
  issueCodes,
  roundTripDocx,
} from './harness';

/**
 * R0 no-clobber corpus.
 *
 * Permanent gate: representative DOCX fixtures reopen without silent identity
 * loss; intentional normalizations appear in compatibility diagnostics;
 * active content stays fail-closed on export.
 */
describe('R0 no-clobber corpus', () => {
  test('bookmarks and external links survive import → export → reopen', async () => {
    const bytes = await buildBookmarksAndLinksFixture();
    const result = await roundTripDocx(bytes, 'report-bookmarks-links.docx');

    expect(result.identities.bookmarkNames).toContain('Architecture');
    expect(result.identities.hrefs).toContain('https://a3s.dev/office');
    expect(result.identities.plainText).toContain('Architecture overview');
    expect(result.identities.plainText).toContain('Product site');
    expectIssueCodes(result.firstPassIssues, ['docx.bookmarks-links']);
  });

  test('internal hyperlink anchors keep bookmark identity across round trip', async () => {
    const bytes = await buildInternalBookmarkLinkFixture();
    const result = await roundTripDocx(bytes, 'report-internal-link.docx');

    expect(result.identities.bookmarkNames).toContain('Obligations');
    expect(result.identities.hrefs).toContain('#Obligations');
    expect(result.identities.plainText).toContain('Obligations section');
    expect(result.identities.plainText).toContain('See obligations');
    expectIssueCodes(result.firstPassIssues, ['docx.bookmarks-links']);
  });

  test('review track-change authors and texts survive round trip', async () => {
    const bytes = await buildReviewTrackChangesFixture();
    const result = await roundTripDocx(bytes, 'review-track-changes.docx');

    expect(result.identities.changeAuthors).toEqual(['Ada Reviewer']);
    expect(result.identities.changeTexts).toEqual(
      expect.arrayContaining(['Added warranty', 'Remove liability']),
    );
    expect(result.identities.plainText).toContain('Keep this clause');
  });

  test('review comment authors and texts survive round trip', async () => {
    const bytes = await buildReviewCommentsFixture();
    const result = await roundTripDocx(bytes, 'review-comments.docx');

    expect(result.identities.commentAuthors).toEqual(['Bea Counsel']);
    expect(result.identities.commentTexts).toEqual(['Clarify liability cap']);
    expect(result.identities.plainText).toContain('Liability clause');
    expect(result.exportedParts).toContain('word/comments.xml');
    expectIssueCodes(result.firstPassIssues, ['docx.comments']);
  });

  test('default header and footer text survive round trip', async () => {
    const bytes = await buildHeaderFooterFixture();
    const result = await roundTripDocx(bytes, 'report-header-footer.docx');

    expect(result.identities.headerTexts).toEqual(['Acme Report Header']);
    expect(result.identities.footerTexts).toEqual(['Confidential Footer']);
    expect(result.identities.plainText).toContain('Report body clause');
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expectIssueCodes(result.firstPassIssues, ['docx.headers']);
  });

  test('footnote reference and note body text survive round trip', async () => {
    const bytes = await buildFootnoteFixture();
    const result = await roundTripDocx(bytes, 'report-footnote.docx');

    expect(result.identities.footnoteTexts).toEqual([
      'Cite the warranty clause',
    ]);
    expect(result.identities.plainText).toContain('Body clause');
    expect(result.identities.plainText).toContain('Cite the warranty clause');
    expect(result.exportedParts).toContain('word/footnotes.xml');
    expectIssueCodes(result.firstPassIssues, ['docx.notes']);
  });

  test('endnote reference and note body text survive round trip', async () => {
    const bytes = await buildEndnoteFixture();
    const result = await roundTripDocx(bytes, 'report-endnote.docx');

    expect(result.identities.endnoteTexts).toEqual(['See appendix A']);
    expect(result.identities.plainText).toContain('Body with endnote');
    expect(result.identities.plainText).toContain('See appendix A');
    expect(result.exportedParts).toContain('word/endnotes.xml');
    expectIssueCodes(result.firstPassIssues, ['docx.notes']);
  });

  test('text content control alias, tag, and text survive round trip', async () => {
    const bytes = await buildTextContentControlFixture();
    const result = await roundTripDocx(bytes, 'contract-content-control.docx');

    expect(result.identities.contentControlAliases).toEqual(['PartyName']);
    expect(result.identities.contentControlTags).toEqual(['party_name']);
    expect(result.identities.contentControlTexts).toEqual(['Acme Corp']);
    expect(result.identities.plainText).toContain('Acme Corp');
    expectIssueCodes(result.firstPassIssues, ['docx.content-controls']);
  });

  test('contract table cell identities survive round trip', async () => {
    const bytes = await buildContractTableFixture();
    const result = await roundTripDocx(bytes, 'contract-table.docx');

    expect(result.identities.tableCellTexts).toEqual([
      'Obligation',
      'Party A',
      'Party B',
      'Payment term',
    ]);
  });

  test('duplicate bookmark names diagnose identity normalization without silent clobber', async () => {
    const bytes = await buildDuplicateBookmarkFixture();
    const imported = await importOfficeFile(
      new File([bytes], 'duplicate-bookmarks.docx', {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }),
    );
    try {
      if (imported.content.type !== 'document') {
        throw new Error('Expected a document artifact.');
      }
      expectIssueCodes(imported.compatibility?.issues ?? [], [
        'docx.bookmarks-links',
        'docx.bookmarks.identity',
      ]);
      const names = extractDocumentIdentities(
        imported.content.html,
      ).bookmarkNames;
      expect(names.length).toBeGreaterThanOrEqual(2);
      expect(new Set(names.map((name) => name.toLowerCase())).size).toBe(
        names.length,
      );
      expect(imported.content.html).toContain('First');
      expect(imported.content.html).toContain('Second');
    } finally {
      forgetSourceBlob(imported.id);
    }
  });

  test('active content stays fail-closed while safe vendor parts can survive', async () => {
    const { bytes, vendorPayload } =
      await buildActiveContentFailClosedFixture();
    const imported = await importOfficeFile(
      new File([bytes], 'active-content-boundary.docx', {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }),
    );
    try {
      expect(issueCodes(imported.compatibility?.issues ?? [])).toEqual(
        expect.arrayContaining(['docx.active-content.removed']),
      );
      if (imported.content.type !== 'document') {
        throw new Error('Expected a document artifact.');
      }
      imported.content = {
        ...imported.content,
        html: imported.content.html.replace(
          'Original active-content boundary',
          'Edited active-content boundary',
        ),
        model: undefined,
      };
      const exported = await createArtifactBlob(imported);
      const archive = await JSZip.loadAsync(await exported.arrayBuffer());
      expect(archive.file('word/vbaProject.bin')).toBeNull();
      expect(archive.file('_xmlsignatures/sig1.xml')).toBeNull();
      await expect(
        archive.file('word/vendorData/payload.bin')?.async('uint8array'),
      ).resolves.toEqual(vendorPayload);
      await expect(
        archive.file('word/document.xml')?.async('text'),
      ).resolves.toContain('Edited active-content boundary');
    } finally {
      forgetSourceBlob(imported.id);
    }
  });
});
