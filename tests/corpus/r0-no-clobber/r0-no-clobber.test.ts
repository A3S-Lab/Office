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
  buildCaptionRefFixture,
  buildContractTableFixture,
  buildDateFieldFixture,
  buildDateFormatFieldFixture,
  buildDuplicateBookmarkFixture,
  buildEndnoteFixture,
  buildFigureCaptionFixture,
  buildFootnoteFixture,
  buildHeaderFooterComplexFieldFixture,
  buildHeaderFooterFieldFixture,
  buildHeaderFooterFixture,
  buildIndexEntryFixture,
  buildIndexFieldFixture,
  buildInternalBookmarkLinkFixture,
  buildNumCharsFieldFixture,
  buildNumPagesFieldFixture,
  buildNumWordsFieldFixture,
  buildPageFieldFixture,
  buildPageRefFieldFixture,
  buildReviewCommentsFixture,
  buildReviewTrackChangesFixture,
  buildRichTextContentControlFixture,
  buildSectionFieldFixture,
  buildSectionPagesFieldFixture,
  buildTableCaptionFixture,
  buildTableOfContentsFixture,
  buildTextContentControlFixture,
  buildTimeFieldFixture,
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

  test('header PAGE and footer NUMPAGES fields survive round trip', async () => {
    const bytes = await buildHeaderFooterFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-header-footer-fields.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['numPages', 'page']);
    expect(result.identities.fieldInstructions).toEqual(['NUMPAGES', 'PAGE']);
    expect(result.identities.headerTexts.join(' ')).toContain('Page');
    expect(result.identities.footerTexts.join(' ')).toContain('of');
    expect(result.identities.plainText).toContain(
      'Report body with live chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('complex header PAGE and footer NUMPAGES fields survive round trip', async () => {
    const bytes = await buildHeaderFooterComplexFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-header-footer-complex-fields.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['numPages', 'page']);
    expect(result.identities.fieldInstructions).toEqual(['NUMPAGES', 'PAGE']);
    expect(result.identities.headerTexts.join(' ')).toContain('Page');
    expect(result.identities.footerTexts.join(' ')).toContain('of');
    expect(result.identities.plainText).toContain(
      'Report body with complex chrome fields',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
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
    expect(result.identities.contentControlTypes).toEqual(['text']);
    expect(result.identities.plainText).toContain('Acme Corp');
    expectIssueCodes(result.firstPassIssues, ['docx.content-controls']);
  });

  test('rich-text content control alias, tag, type, and text survive round trip', async () => {
    const bytes = await buildRichTextContentControlFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-richtext-content-control.docx',
    );

    expect(result.identities.contentControlAliases).toEqual(['ClauseBody']);
    expect(result.identities.contentControlTags).toEqual(['clause_body']);
    expect(result.identities.contentControlTexts).toEqual([
      'Indemnity survives termination.',
    ]);
    expect(result.identities.contentControlTypes).toEqual(['richText']);
    expect(result.identities.plainText).toContain(
      'Indemnity survives termination.',
    );
    expectIssueCodes(result.firstPassIssues, ['docx.content-controls']);
  });

  test('PAGE field kind and instruction survive round trip', async () => {
    const bytes = await buildPageFieldFixture();
    const result = await roundTripDocx(bytes, 'report-page-field.docx');

    expect(result.identities.fieldKinds).toEqual(['page']);
    expect(result.identities.fieldInstructions).toEqual(['PAGE']);
    expect(result.identities.plainText).toContain('Page');
    expectIssueCodes(result.firstPassIssues, ['docx.fields.body']);
  });

  test('bookmark-backed PAGEREF kind, instruction, and target survive round trip', async () => {
    const bytes = await buildPageRefFieldFixture();
    const result = await roundTripDocx(bytes, 'report-pageref-field.docx');

    expect(result.identities.bookmarkNames).toContain('Warranty');
    expect(result.identities.fieldKinds).toEqual(['pageReference']);
    expect(result.identities.fieldInstructions).toEqual([
      'PAGEREF Warranty \\h',
    ]);
    expect(result.identities.fieldTargetNames).toEqual(['Warranty']);
    expect(result.identities.plainText).toContain('Warranty clause');
    expect(result.identities.plainText).toContain('See page');
    expectIssueCodes(result.firstPassIssues, ['docx.fields.body']);
  });

  test('NUMPAGES field kind and instruction survive round trip', async () => {
    const bytes = await buildNumPagesFieldFixture();
    const result = await roundTripDocx(bytes, 'report-numpages-field.docx');

    expect(result.identities.fieldKinds).toEqual(['numPages']);
    expect(result.identities.fieldInstructions).toEqual(['NUMPAGES']);
    expect(result.identities.plainText).toContain('Pages of');
    expectIssueCodes(result.firstPassIssues, ['docx.fields.body']);
  });

  test('SECTION field kind and instruction survive round trip', async () => {
    const bytes = await buildSectionFieldFixture();
    const result = await roundTripDocx(bytes, 'report-section-field.docx');

    expect(result.identities.fieldKinds).toEqual(['section']);
    expect(result.identities.fieldInstructions).toEqual(['SECTION']);
    expect(result.identities.plainText).toContain('Section');
    expectIssueCodes(result.firstPassIssues, ['docx.fields.body']);
  });

  test('SECTIONPAGES field kind and instruction survive round trip', async () => {
    const bytes = await buildSectionPagesFieldFixture();
    const result = await roundTripDocx(bytes, 'report-sectionpages-field.docx');

    expect(result.identities.fieldKinds).toEqual(['sectionPages']);
    expect(result.identities.fieldInstructions).toEqual(['SECTIONPAGES']);
    expect(result.identities.plainText).toContain('Section pages');
    expectIssueCodes(result.firstPassIssues, ['docx.fields.body']);
  });

  test('NUMWORDS field kind and instruction survive round trip', async () => {
    const bytes = await buildNumWordsFieldFixture();
    const result = await roundTripDocx(bytes, 'report-numwords-field.docx');

    expect(result.identities.fieldKinds).toEqual(['wordCount']);
    expect(result.identities.fieldInstructions).toEqual(['NUMWORDS']);
    expect(result.identities.plainText).toContain('Words');
    expectIssueCodes(result.firstPassIssues, ['docx.fields.body']);
  });

  test('NUMCHARS field kind and instruction survive round trip', async () => {
    const bytes = await buildNumCharsFieldFixture();
    const result = await roundTripDocx(bytes, 'report-numchars-field.docx');

    expect(result.identities.fieldKinds).toEqual(['characterCount']);
    expect(result.identities.fieldInstructions).toEqual(['NUMCHARS']);
    expect(result.identities.plainText).toContain('Chars');
    expectIssueCodes(result.firstPassIssues, ['docx.fields.body']);
  });

  test('DATE field kind and instruction survive round trip', async () => {
    const bytes = await buildDateFieldFixture();
    const result = await roundTripDocx(bytes, 'report-date-field.docx');

    expect(result.identities.fieldKinds).toEqual(['date']);
    expect(result.identities.fieldInstructions).toEqual(['DATE']);
    expect(result.identities.plainText).toContain('Date');
    expectIssueCodes(result.firstPassIssues, ['docx.fields.body']);
  });

  test('TIME field kind and instruction survive round trip', async () => {
    const bytes = await buildTimeFieldFixture();
    const result = await roundTripDocx(bytes, 'report-time-field.docx');

    expect(result.identities.fieldKinds).toEqual(['time']);
    expect(result.identities.fieldInstructions).toEqual(['TIME']);
    expect(result.identities.plainText).toContain('Time');
    expectIssueCodes(result.firstPassIssues, ['docx.fields.body']);
  });

  test('DATE format-switch instruction survives round trip', async () => {
    const bytes = await buildDateFormatFieldFixture();
    const result = await roundTripDocx(bytes, 'report-date-format-field.docx');

    expect(result.identities.fieldKinds).toEqual(['date']);
    expect(result.identities.fieldInstructions).toEqual([
      'DATE \\@ "yyyy-MM-dd"',
    ]);
    expect(result.identities.plainText).toContain('Formatted');
    expectIssueCodes(result.firstPassIssues, ['docx.fields.body']);
  });

  test('figure SEQ caption kind and id survive round trip', async () => {
    const bytes = await buildFigureCaptionFixture();
    const result = await roundTripDocx(bytes, 'report-figure-caption.docx');

    expect(result.identities.captionKinds).toEqual(['figure']);
    expect(result.identities.captionIds.length).toBeGreaterThanOrEqual(1);
    expect(result.identities.captionIds[0]).toMatch(
      /^docx-(?:caption-|figure-caption-)/,
    );
    expect(result.identities.plainText).toContain('Architecture');
    expectIssueCodes(result.firstPassIssues, ['docx.captions']);
  });

  test('caption REF target identity survives round trip', async () => {
    const bytes = await buildCaptionRefFixture();
    const result = await roundTripDocx(bytes, 'report-caption-ref.docx');

    expect(result.identities.captionKinds).toEqual(['figure']);
    expect(result.identities.captionIds.length).toBeGreaterThanOrEqual(1);
    expect(result.identities.crossReferenceTargetIds).toEqual(
      result.identities.captionIds,
    );
    expect(result.identities.plainText).toContain('Runtime');
    expect(result.identities.plainText).toContain('See');
    expectIssueCodes(result.firstPassIssues, ['docx.captions']);
  });

  test('table SEQ caption kind and id survive round trip', async () => {
    const bytes = await buildTableCaptionFixture();
    const result = await roundTripDocx(bytes, 'contract-table-caption.docx');

    expect(result.identities.captionKinds).toEqual(['table']);
    expect(result.identities.captionIds.length).toBeGreaterThanOrEqual(1);
    expect(result.identities.captionIds[0]).toMatch(
      /^docx-(?:caption-|table-caption-)/,
    );
    expect(result.identities.plainText).toContain('Schedule');
    expectIssueCodes(result.firstPassIssues, ['docx.captions']);
  });

  test('XE index entry main and sub-entry survive round trip', async () => {
    const bytes = await buildIndexEntryFixture();
    const result = await roundTripDocx(bytes, 'academic-index-entry.docx');

    expect(result.identities.indexMainEntries).toContain('Architecture');
    expect(result.identities.indexSubEntries).toContain('Runtime');
    expect(result.identities.plainText).toContain('Runtime');
    expectIssueCodes(result.firstPassIssues, ['docx.index']);
  });

  test('INDEX field columns survive round trip', async () => {
    const bytes = await buildIndexFieldFixture();
    const result = await roundTripDocx(bytes, 'academic-index-field.docx');

    expect(result.identities.indexColumns).toEqual(['2']);
    expect(result.identities.indexMainEntries).toContain('Architecture');
    expect(result.identities.plainText).toContain('Architecture');
    expectIssueCodes(result.firstPassIssues, ['docx.index']);
  });

  test('TOC options and cached entry title survive round trip', async () => {
    const bytes = await buildTableOfContentsFixture();
    const result = await roundTripDocx(bytes, 'report-table-of-contents.docx');

    expect(result.identities.tocMinLevels).toEqual(['1']);
    expect(result.identities.tocMaxLevels).toEqual(['3']);
    expect(result.identities.tocHyperlinks).toEqual(['true']);
    expect(result.identities.tocEntryTitles).toContain('Overview');
    expect(result.identities.plainText).toContain('Overview');
    expectIssueCodes(result.firstPassIssues, ['docx.tableOfContents']);
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
