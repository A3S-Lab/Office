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
  buildInternalBookmarkLinkFixture,
  buildReviewCommentsFixture,
  buildReviewTrackChangesFixture,
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
