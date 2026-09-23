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
  buildBulletNumberingRevisionFixture,
  buildCaptionRefFixture,
  buildCellBordersFormattingRevisionFixture,
  buildCellCnfStyleFormattingRevisionFixture,
  buildCellFillFormattingRevisionFixture,
  buildCellFitTextFormattingRevisionFixture,
  buildCellFormattingRevisionFixture,
  buildCellGridSpanFormattingRevisionFixture,
  buildCellHideMarkFormattingRevisionFixture,
  buildCellHMergeFormattingRevisionFixture,
  buildCellMarginFormattingRevisionFixture,
  buildCellNoWrapFormattingRevisionFixture,
  buildCellTextDirectionFormattingRevisionFixture,
  buildCellVMergeFormattingRevisionFixture,
  buildCellWidthFormattingRevisionFixture,
  buildCharacterFormattingRevisionFixture,
  buildContractTableFixture,
  buildDateFieldFixture,
  buildDateFormatFieldFixture,
  buildDecimalNumberingRevisionFixture,
  buildDuplicateBookmarkFixture,
  buildEndnoteFixture,
  buildEvenPageFooterComplexNumPagesFieldFixture,
  buildEvenPageFooterComplexSectionPagesFieldFixture,
  buildEvenPageFooterComplexTimeFieldFixture,
  buildEvenPageFooterNumPagesFieldFixture,
  buildEvenPageFooterSectionPagesFieldFixture,
  buildEvenPageFooterTimeFieldFixture,
  buildEvenPageHeaderComplexDateFieldFixture,
  buildEvenPageHeaderComplexPageFieldFixture,
  buildEvenPageHeaderComplexSectionFieldFixture,
  buildEvenPageHeaderDateFieldFixture,
  buildEvenPageHeaderFooterFixture,
  buildEvenPageHeaderPageFieldFixture,
  buildEvenPageHeaderSectionFieldFixture,
  buildFigureCaptionFixture,
  buildFirstPageFooterComplexNumPagesFieldFixture,
  buildFirstPageFooterComplexSectionPagesFieldFixture,
  buildFirstPageFooterComplexTimeFieldFixture,
  buildFirstPageFooterNumPagesFieldFixture,
  buildFirstPageFooterSectionPagesFieldFixture,
  buildFirstPageFooterTimeFieldFixture,
  buildFirstPageHeaderComplexDateFieldFixture,
  buildFirstPageHeaderComplexPageFieldFixture,
  buildFirstPageHeaderComplexSectionFieldFixture,
  buildFirstPageHeaderDateFieldFixture,
  buildFirstPageHeaderFooterFixture,
  buildFirstPageHeaderPageFieldFixture,
  buildFirstPageHeaderSectionFieldFixture,
  buildFootnoteFixture,
  buildHeaderFooterComplexDateTimeFieldFixture,
  buildHeaderFooterComplexFieldFixture,
  buildHeaderFooterComplexSectionFieldFixture,
  buildHeaderFooterDateTimeFieldFixture,
  buildHeaderFooterFieldFixture,
  buildHeaderFooterFixture,
  buildHeaderFooterSectionFieldFixture,
  buildIndexEntryFixture,
  buildIndexFieldFixture,
  buildInternalBookmarkLinkFixture,
  buildMultiLevelNumberingRevisionFixture,
  buildNumCharsFieldFixture,
  buildNumPagesFieldFixture,
  buildNumWordsFieldFixture,
  buildPageFieldFixture,
  buildPageRefFieldFixture,
  buildParagraphBreakSplitRevisionFixture,
  buildParagraphFormattingRevisionFixture,
  buildParagraphMarkInsertionRevisionFixture,
  buildReviewCommentsFixture,
  buildReviewTrackChangesFixture,
  buildRichTextContentControlFixture,
  buildRowAlignmentFormattingRevisionFixture,
  buildRowCellSpacingFormattingRevisionFixture,
  buildRowCnfStyleFormattingRevisionFixture,
  buildRowDivIdFormattingRevisionFixture,
  buildRowFormattingRevisionFixture,
  buildRowGridAfterFormattingRevisionFixture,
  buildRowGridBeforeFormattingRevisionFixture,
  buildRowHeaderFormattingRevisionFixture,
  buildRowHeightFormattingRevisionFixture,
  buildRowHiddenFormattingRevisionFixture,
  buildRowWidthAfterFormattingRevisionFixture,
  buildRowWidthBeforeFormattingRevisionFixture,
  buildSectionBidiFormattingRevisionFixture,
  buildSectionDocGridFormattingRevisionFixture,
  buildSectionEndnotePrFormattingRevisionFixture,
  buildSectionEqualColsFormattingRevisionFixture,
  buildSectionFieldFixture,
  buildSectionFootnotePrFormattingRevisionFixture,
  buildSectionFormattingRevisionFixture,
  buildSectionFormProtFormattingRevisionFixture,
  buildSectionLnNumTypeFormattingRevisionFixture,
  buildSectionMarginFormattingRevisionFixture,
  buildSectionNoEndnoteFormattingRevisionFixture,
  buildSectionPageBordersFormattingRevisionFixture,
  buildSectionPageGeometryFormattingRevisionFixture,
  buildSectionPagesFieldFixture,
  buildSectionPaperSourceFormattingRevisionFixture,
  buildSectionPgNumTypeFormattingRevisionFixture,
  buildSectionRtlGutterFormattingRevisionFixture,
  buildSectionTextDirectionFormattingRevisionFixture,
  buildSectionTitlePageFormattingRevisionFixture,
  buildSectionTypeFormattingRevisionFixture,
  buildSectionUnequalColsFormattingRevisionFixture,
  buildSectionVerticalAlignFormattingRevisionFixture,
  buildTableBidiVisualFormattingRevisionFixture,
  buildTableBordersFormattingRevisionFixture,
  buildTableCaptionFixture,
  buildTableCaptionFormattingRevisionFixture,
  buildTableCellMarginFormattingRevisionFixture,
  buildTableCellSpacingFormattingRevisionFixture,
  buildTableColBandSizeFormattingRevisionFixture,
  buildTableDescriptionFormattingRevisionFixture,
  buildTableFillFormattingRevisionFixture,
  buildTableFloatFormattingRevisionFixture,
  buildTableFormattingRevisionFixture,
  buildTableIndentFormattingRevisionFixture,
  buildTableLayoutFormattingRevisionFixture,
  buildTableLookFormattingRevisionFixture,
  buildTableOfContentsFixture,
  buildTableOverlapFormattingRevisionFixture,
  buildTableRowBandSizeFormattingRevisionFixture,
  buildTableStyleFormattingRevisionFixture,
  buildTableWidthFormattingRevisionFixture,
  buildTextContentControlFixture,
  buildTextMoveRangeCompanionRevisionFixture,
  buildTextMoveRevisionFixture,
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

  test('character-formatting revision author and kind survive round trip', async () => {
    const bytes = await buildCharacterFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-character-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['formatting']);
    expect(result.identities.plainText).toContain('Formatted clause text');
    expectIssueCodes(result.firstPassIssues, ['docx.revisions.formatting']);
  });

  test('paragraph-formatting revision author and kind survive round trip', async () => {
    const bytes = await buildParagraphFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-paragraph-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['paragraph-formatting']);
    expect(result.identities.plainText).toContain('Aligned paragraph clause');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.paragraph-formatting',
    ]);
  });

  test('paragraph-mark insertion revision author and kind survive round trip', async () => {
    const bytes = await buildParagraphMarkInsertionRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-paragraph-mark-insertion-revision.docx',
    );

    expect(result.identities.blockChangeAuthors).toEqual(['Reviewer']);
    expect(result.identities.blockChangeKinds).toEqual(['insertion']);
    expect(result.identities.plainText).toContain('Added whole paragraph');
    expectIssueCodes(result.firstPassIssues, ['docx.revisions.paragraph-mark']);
  });

  test('paragraph-break split revision author and kind survive round trip', async () => {
    const bytes = await buildParagraphBreakSplitRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-paragraph-break-split-revision.docx',
    );

    expect(result.identities.paragraphBreakAuthors).toEqual(['Reviewer']);
    expect(result.identities.paragraphBreakKinds).toEqual(['split']);
    expect(result.identities.plainText).toContain('Before split clause');
    expect(result.identities.plainText).toContain('After split clause');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.paragraph-break',
    ]);
  });

  test('text-move revision author and kind survive round trip', async () => {
    const bytes = await buildTextMoveRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-text-move-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['move']);
    expect(result.identities.plainText).toContain('Moved clause');
    expect(result.identities.plainText).toContain('middle');
    expectIssueCodes(result.firstPassIssues, ['docx.revisions.move']);
  });

  test('text-move range companion name survives round trip', async () => {
    const bytes = await buildTextMoveRangeCompanionRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-text-move-range-companion-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['move']);
    expect(result.identities.changeMoveRangeNames).toEqual(['move0']);
    expect(result.identities.plainText).toContain('Ranged move clause');
    expect(result.identities.plainText).toContain('spacer');
    expectIssueCodes(result.firstPassIssues, ['docx.revisions.move']);
  });

  test('table-formatting revision author and kind survive round trip', async () => {
    const bytes = await buildTableFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-table-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['table-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Party obligation']);
    expect(result.identities.plainText).toContain('Party obligation');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.table-formatting',
    ]);
  });

  test('table-width formatting revision author and kind survive round trip', async () => {
    const bytes = await buildTableWidthFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-table-width-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['table-formatting']);
    expect(result.identities.tableCellTexts).toEqual([
      'Preferred width clause',
    ]);
    expect(result.identities.plainText).toContain('Preferred width clause');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.table-formatting',
    ]);
  });

  test('table-indent formatting revision author and kind survive round trip', async () => {
    const bytes = await buildTableIndentFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-table-indent-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['table-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Indented table clause']);
    expect(result.identities.plainText).toContain('Indented table clause');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.table-formatting',
    ]);
  });

  test('table-cell-margin formatting revision author and kind survive round trip', async () => {
    const bytes = await buildTableCellMarginFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-table-cell-margin-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['table-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Padded table clause']);
    expect(result.identities.plainText).toContain('Padded table clause');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.table-formatting',
    ]);
  });

  test('table-layout formatting revision author and kind survive round trip', async () => {
    const bytes = await buildTableLayoutFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-table-layout-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['table-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Fixed layout clause']);
    expect(result.identities.plainText).toContain('Fixed layout clause');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.table-formatting',
    ]);
  });

  test('table-bidi-visual formatting revision author and kind survive round trip', async () => {
    const bytes = await buildTableBidiVisualFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-table-bidi-visual-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['table-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Bidi visual clause']);
    expect(result.identities.plainText).toContain('Bidi visual clause');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.table-formatting',
    ]);
  });

  test('table-fill formatting revision author and kind survive round trip', async () => {
    const bytes = await buildTableFillFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-table-fill-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['table-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Filled table clause']);
    expect(result.identities.plainText).toContain('Filled table clause');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.table-formatting',
    ]);
  });

  test('table-overlap formatting revision author and kind survive round trip', async () => {
    const bytes = await buildTableOverlapFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-table-overlap-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['table-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Overlap table clause']);
    expect(result.identities.plainText).toContain('Overlap table clause');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.table-formatting',
    ]);
  });

  test('table-look formatting revision author and kind survive round trip', async () => {
    const bytes = await buildTableLookFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-table-look-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['table-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Look table clause']);
    expect(result.identities.plainText).toContain('Look table clause');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.table-formatting',
    ]);
  });

  test('table-style formatting revision author and kind survive round trip', async () => {
    const bytes = await buildTableStyleFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-table-style-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['table-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Styled table clause']);
    expect(result.identities.plainText).toContain('Styled table clause');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.table-formatting',
    ]);
  });

  test('table-cell-spacing formatting revision author and kind survive round trip', async () => {
    const bytes = await buildTableCellSpacingFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-table-cell-spacing-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['table-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Spaced table clause']);
    expect(result.identities.plainText).toContain('Spaced table clause');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.table-formatting',
    ]);
  });

  test('table-borders formatting revision author and kind survive round trip', async () => {
    const bytes = await buildTableBordersFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-table-borders-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['table-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Bordered table clause']);
    expect(result.identities.plainText).toContain('Bordered table clause');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.table-formatting',
    ]);
  });

  test('table-col-band-size formatting revision author and kind survive round trip', async () => {
    const bytes = await buildTableColBandSizeFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-table-col-band-size-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['table-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Col-band table clause']);
    expect(result.identities.plainText).toContain('Col-band table clause');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.table-formatting',
    ]);
  });

  test('table-row-band-size formatting revision author and kind survive round trip', async () => {
    const bytes = await buildTableRowBandSizeFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-table-row-band-size-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['table-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Row-band table clause']);
    expect(result.identities.plainText).toContain('Row-band table clause');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.table-formatting',
    ]);
  });

  test('table-caption formatting revision author and kind survive round trip', async () => {
    const bytes = await buildTableCaptionFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-table-caption-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['table-formatting']);
    expect(result.identities.tableCellTexts).toEqual([
      'Captioned table clause',
    ]);
    expect(result.identities.plainText).toContain('Captioned table clause');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.table-formatting',
    ]);
  });

  test('table-description formatting revision author and kind survive round trip', async () => {
    const bytes = await buildTableDescriptionFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-table-description-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['table-formatting']);
    expect(result.identities.tableCellTexts).toEqual([
      'Described table clause',
    ]);
    expect(result.identities.plainText).toContain('Described table clause');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.table-formatting',
    ]);
  });

  test('table-float formatting revision author and kind survive round trip', async () => {
    const bytes = await buildTableFloatFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-table-float-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['table-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Floating table clause']);
    expect(result.identities.plainText).toContain('Floating table clause');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.table-formatting',
    ]);
  });

  test('section-formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with section orientation revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('section-margin formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionMarginFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-margin-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with section margin revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('section-title-page formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionTitlePageFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-title-page-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with title-page revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('section-equal-cols formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionEqualColsFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-equal-cols-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with equal columns revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('section-rtl-gutter formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionRtlGutterFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-rtl-gutter-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with RTL gutter revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('section-paper-source formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionPaperSourceFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-paper-source-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with paper-source revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('section-vertical-align formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionVerticalAlignFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-vertical-align-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with section vertical-align revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('section-doc-grid formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionDocGridFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-doc-grid-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with document-grid revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('section-bidi formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionBidiFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-bidi-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with section bidi revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('section-type formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionTypeFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-type-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with section-break type revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('section-page-geometry formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionPageGeometryFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-page-geometry-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with page-geometry revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('section-unequal-cols formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionUnequalColsFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-unequal-cols-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with unequal columns revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('section-ln-num-type formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionLnNumTypeFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-ln-num-type-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with line-numbering revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('section-pg-num-type formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionPgNumTypeFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-pg-num-type-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with page-number format revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('section-form-prot formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionFormProtFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-form-prot-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with form-protection revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('section-no-endnote formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionNoEndnoteFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-no-endnote-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with no-endnote revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('section-text-direction formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionTextDirectionFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-text-direction-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with section text-direction revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('section-footnote-pr formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionFootnotePrFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-footnote-pr-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with footnote-properties revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('section-endnote-pr formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionEndnotePrFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-endnote-pr-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with endnote-properties revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('section-page-borders formatting revision author and kind survive round trip', async () => {
    const bytes = await buildSectionPageBordersFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'report-section-page-borders-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['section-formatting']);
    expect(result.identities.plainText).toContain(
      'Report body with page-borders revision',
    );
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.section-formatting',
    ]);
  });

  test('row-formatting revision author and kind survive round trip', async () => {
    const bytes = await buildRowFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-row-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['row-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Keep row together']);
    expect(result.identities.plainText).toContain('Keep row together');
    expectIssueCodes(result.firstPassIssues, ['docx.revisions.row-formatting']);
  });

  test('row-header formatting revision author and kind survive round trip', async () => {
    const bytes = await buildRowHeaderFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-row-header-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['row-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Repeat header row']);
    expect(result.identities.plainText).toContain('Repeat header row');
    expectIssueCodes(result.firstPassIssues, ['docx.revisions.row-formatting']);
  });

  test('row-height formatting revision author and kind survive round trip', async () => {
    const bytes = await buildRowHeightFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-row-height-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['row-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Fixed row height']);
    expect(result.identities.plainText).toContain('Fixed row height');
    expectIssueCodes(result.firstPassIssues, ['docx.revisions.row-formatting']);
  });

  test('row-hidden formatting revision author and kind survive round trip', async () => {
    const bytes = await buildRowHiddenFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-row-hidden-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['row-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Hidden contract row']);
    expect(result.identities.plainText).toContain('Hidden contract row');
    expectIssueCodes(result.firstPassIssues, ['docx.revisions.row-formatting']);
  });

  test('row-alignment formatting revision author and kind survive round trip', async () => {
    const bytes = await buildRowAlignmentFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-row-alignment-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['row-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Aligned contract row']);
    expect(result.identities.plainText).toContain('Aligned contract row');
    expectIssueCodes(result.firstPassIssues, ['docx.revisions.row-formatting']);
  });

  test('row-grid-before formatting revision author and kind survive round trip', async () => {
    const bytes = await buildRowGridBeforeFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-row-grid-before-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['row-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Leading grid skip']);
    expect(result.identities.plainText).toContain('Leading grid skip');
    expectIssueCodes(result.firstPassIssues, ['docx.revisions.row-formatting']);
  });

  test('row-grid-after formatting revision author and kind survive round trip', async () => {
    const bytes = await buildRowGridAfterFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-row-grid-after-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['row-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Trailing grid skip']);
    expect(result.identities.plainText).toContain('Trailing grid skip');
    expectIssueCodes(result.firstPassIssues, ['docx.revisions.row-formatting']);
  });

  test('row-width-before formatting revision author and kind survive round trip', async () => {
    const bytes = await buildRowWidthBeforeFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-row-width-before-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['row-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Leading row indent']);
    expect(result.identities.plainText).toContain('Leading row indent');
    expectIssueCodes(result.firstPassIssues, ['docx.revisions.row-formatting']);
  });

  test('row-width-after formatting revision author and kind survive round trip', async () => {
    const bytes = await buildRowWidthAfterFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-row-width-after-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['row-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Trailing row indent']);
    expect(result.identities.plainText).toContain('Trailing row indent');
    expectIssueCodes(result.firstPassIssues, ['docx.revisions.row-formatting']);
  });

  test('row-cnf-style formatting revision author and kind survive round trip', async () => {
    const bytes = await buildRowCnfStyleFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-row-cnf-style-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['row-formatting']);
    expect(result.identities.tableCellTexts).toEqual([
      'Conditional format row',
    ]);
    expect(result.identities.plainText).toContain('Conditional format row');
    expectIssueCodes(result.firstPassIssues, ['docx.revisions.row-formatting']);
  });

  test('row-div-id formatting revision author and kind survive round trip', async () => {
    const bytes = await buildRowDivIdFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-row-div-id-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['row-formatting']);
    expect(result.identities.tableCellTexts).toEqual([
      'Div-mapped contract row',
    ]);
    expect(result.identities.plainText).toContain('Div-mapped contract row');
    expectIssueCodes(result.firstPassIssues, ['docx.revisions.row-formatting']);
  });

  test('row-cell-spacing formatting revision author and kind survive round trip', async () => {
    const bytes = await buildRowCellSpacingFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-row-cell-spacing-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['row-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Spaced contract row']);
    expect(result.identities.plainText).toContain('Spaced contract row');
    expectIssueCodes(result.firstPassIssues, ['docx.revisions.row-formatting']);
  });

  test('cell-formatting revision author and kind survive round trip', async () => {
    const bytes = await buildCellFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-cell-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['cell-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Aligned cell']);
    expect(result.identities.plainText).toContain('Aligned cell');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.cell-formatting',
    ]);
  });

  test('cell-fill formatting revision author and kind survive round trip', async () => {
    const bytes = await buildCellFillFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-cell-fill-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['cell-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Filled contract cell']);
    expect(result.identities.plainText).toContain('Filled contract cell');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.cell-formatting',
    ]);
  });

  test('cell-width formatting revision author and kind survive round trip', async () => {
    const bytes = await buildCellWidthFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-cell-width-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['cell-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Width contract cell']);
    expect(result.identities.plainText).toContain('Width contract cell');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.cell-formatting',
    ]);
  });

  test('cell-margin formatting revision author and kind survive round trip', async () => {
    const bytes = await buildCellMarginFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-cell-margin-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['cell-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['Margin contract cell']);
    expect(result.identities.plainText).toContain('Margin contract cell');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.cell-formatting',
    ]);
  });

  test('cell-no-wrap formatting revision author and kind survive round trip', async () => {
    const bytes = await buildCellNoWrapFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-cell-no-wrap-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['cell-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['No-wrap contract cell']);
    expect(result.identities.plainText).toContain('No-wrap contract cell');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.cell-formatting',
    ]);
  });

  test('cell-text-direction formatting revision author and kind survive round trip', async () => {
    const bytes = await buildCellTextDirectionFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-cell-text-direction-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['cell-formatting']);
    expect(result.identities.tableCellTexts).toEqual([
      'Direction contract cell',
    ]);
    expect(result.identities.plainText).toContain('Direction contract cell');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.cell-formatting',
    ]);
  });

  test('cell-fit-text formatting revision author and kind survive round trip', async () => {
    const bytes = await buildCellFitTextFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-cell-fit-text-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['cell-formatting']);
    expect(result.identities.tableCellTexts).toEqual([
      'Fit-text contract cell',
    ]);
    expect(result.identities.plainText).toContain('Fit-text contract cell');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.cell-formatting',
    ]);
  });

  test('cell-hide-mark formatting revision author and kind survive round trip', async () => {
    const bytes = await buildCellHideMarkFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-cell-hide-mark-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['cell-formatting']);
    expect(result.identities.tableCellTexts).toEqual([
      'Hide-mark contract cell',
    ]);
    expect(result.identities.plainText).toContain('Hide-mark contract cell');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.cell-formatting',
    ]);
  });

  test('cell-cnf-style formatting revision author and kind survive round trip', async () => {
    const bytes = await buildCellCnfStyleFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-cell-cnf-style-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['cell-formatting']);
    expect(result.identities.tableCellTexts).toEqual([
      'Conditional format cell',
    ]);
    expect(result.identities.plainText).toContain('Conditional format cell');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.cell-formatting',
    ]);
  });

  test('cell-h-merge formatting revision author and kind survive round trip', async () => {
    const bytes = await buildCellHMergeFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-cell-h-merge-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['cell-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['H-merge contract cell']);
    expect(result.identities.plainText).toContain('H-merge contract cell');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.cell-formatting',
    ]);
  });

  test('cell-v-merge formatting revision author and kind survive round trip', async () => {
    const bytes = await buildCellVMergeFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-cell-v-merge-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['cell-formatting']);
    expect(result.identities.tableCellTexts).toEqual(['V-merge contract cell']);
    expect(result.identities.plainText).toContain('V-merge contract cell');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.cell-formatting',
    ]);
  });

  test('cell-grid-span formatting revision author and kind survive round trip', async () => {
    const bytes = await buildCellGridSpanFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-cell-grid-span-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['cell-formatting']);
    expect(result.identities.tableCellTexts).toEqual([
      'Grid-span contract cell',
    ]);
    expect(result.identities.plainText).toContain('Grid-span contract cell');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.cell-formatting',
    ]);
  });

  test('cell-borders formatting revision author and kind survive round trip', async () => {
    const bytes = await buildCellBordersFormattingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-cell-borders-formatting-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Reviewer']);
    expect(result.identities.changeKinds).toEqual(['cell-formatting']);
    expect(result.identities.tableCellTexts).toEqual([
      'Bordered contract cell',
    ]);
    expect(result.identities.plainText).toContain('Bordered contract cell');
    expectIssueCodes(result.firstPassIssues, [
      'docx.revisions.cell-formatting',
    ]);
  });

  test('bullet numbering revision author and kind survive round trip', async () => {
    const bytes = await buildBulletNumberingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-bullet-numbering-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Ada Reviewer']);
    expect(result.identities.changeKinds).toEqual(['numbering']);
    expect(result.identities.plainText).toContain('First obligation');
    expect(result.identities.plainText).toContain('Second obligation');
    expect(result.exportedParts).toContain('word/numbering.xml');
    expectIssueCodes(result.firstPassIssues, ['docx.revisions.numbering']);
  });

  test('decimal numbering revision author and kind survive round trip', async () => {
    const bytes = await buildDecimalNumberingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-decimal-numbering-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Ada Reviewer']);
    expect(result.identities.changeKinds).toEqual(['numbering']);
    expect(result.identities.plainText).toContain('First deliverable');
    expect(result.identities.plainText).toContain('Second deliverable');
    expect(result.exportedParts).toContain('word/numbering.xml');
    expectIssueCodes(result.firstPassIssues, ['docx.revisions.numbering']);
  });

  test('multi-level numbering revision author and kind survive round trip', async () => {
    const bytes = await buildMultiLevelNumberingRevisionFixture();
    const result = await roundTripDocx(
      bytes,
      'contract-multi-level-numbering-revision.docx',
    );

    expect(result.identities.changeAuthors).toEqual(['Ada Reviewer']);
    expect(result.identities.changeKinds).toEqual(['numbering']);
    expect(result.identities.plainText).toContain('First milestone');
    expect(result.identities.plainText).toContain('Second milestone');
    expect(result.exportedParts).toContain('word/numbering.xml');
    expectIssueCodes(result.firstPassIssues, ['docx.revisions.numbering']);
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

  test('first-page and default header/footer text survive round trip', async () => {
    const bytes = await buildFirstPageHeaderFooterFixture();
    const result = await roundTripDocx(
      bytes,
      'report-first-page-header-footer.docx',
    );

    expect(result.identities.headerTexts).toEqual([
      'Acme Continuing Header',
      'Acme Title Header',
    ]);
    expect(result.identities.footerTexts).toEqual([
      'Continuing Footer',
      'Title Footer',
    ]);
    expect(result.identities.plainText).toContain(
      'Report body with first-page chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/header2.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.exportedParts).toContain('word/footer2.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('first-page header PAGE field survives round trip', async () => {
    const bytes = await buildFirstPageHeaderPageFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-first-page-header-page-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['page']);
    expect(result.identities.fieldInstructions).toEqual(['PAGE']);
    expect(result.identities.headerTexts).toEqual([
      'Acme Continuing Header',
      'Title page 1',
    ]);
    expect(result.identities.footerTexts).toContain('Continuing Footer');
    expect(result.identities.plainText).toContain(
      'Report body with first-page PAGE chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/header2.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('complex first-page header PAGE field survives round trip', async () => {
    const bytes = await buildFirstPageHeaderComplexPageFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-complex-first-page-header-page-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['page']);
    expect(result.identities.fieldInstructions).toEqual(['PAGE']);
    expect(result.identities.headerTexts).toEqual([
      'Acme Continuing Header',
      'Title page 1',
    ]);
    expect(result.identities.footerTexts).toContain('Continuing Footer');
    expect(result.identities.plainText).toContain(
      'Report body with complex first-page PAGE chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/header2.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('first-page footer NUMPAGES field survives round trip', async () => {
    const bytes = await buildFirstPageFooterNumPagesFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-first-page-footer-numpages-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['numPages']);
    expect(result.identities.fieldInstructions).toEqual(['NUMPAGES']);
    expect(result.identities.headerTexts).toContain('Acme Continuing Header');
    expect(result.identities.footerTexts).toEqual([
      'Continuing Footer',
      'Title of 12',
    ]);
    expect(result.identities.plainText).toContain(
      'Report body with first-page NUMPAGES chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.exportedParts).toContain('word/footer2.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('complex first-page footer NUMPAGES field survives round trip', async () => {
    const bytes = await buildFirstPageFooterComplexNumPagesFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-complex-first-page-footer-numpages-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['numPages']);
    expect(result.identities.fieldInstructions).toEqual(['NUMPAGES']);
    expect(result.identities.headerTexts).toContain('Acme Continuing Header');
    expect(result.identities.footerTexts).toEqual([
      'Continuing Footer',
      'Title of 12',
    ]);
    expect(result.identities.plainText).toContain(
      'Report body with complex first-page NUMPAGES chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.exportedParts).toContain('word/footer2.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('even-page and default header/footer text survive round trip', async () => {
    const bytes = await buildEvenPageHeaderFooterFixture();
    const result = await roundTripDocx(
      bytes,
      'report-even-page-header-footer.docx',
    );

    expect(result.identities.headerTexts).toEqual([
      'Even Page Header',
      'Odd Page Header',
    ]);
    expect(result.identities.footerTexts).toEqual([
      'Even Page Footer',
      'Odd Page Footer',
    ]);
    expect(result.identities.plainText).toContain(
      'Report body with even-page chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/header2.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.exportedParts).toContain('word/footer2.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('even-page header PAGE field survives round trip', async () => {
    const bytes = await buildEvenPageHeaderPageFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-even-page-header-page-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['page']);
    expect(result.identities.fieldInstructions).toEqual(['PAGE']);
    expect(result.identities.headerTexts).toEqual([
      'Even page 2',
      'Odd Continuing Header',
    ]);
    expect(result.identities.footerTexts).toContain('Continuing Footer');
    expect(result.identities.plainText).toContain(
      'Report body with even-page PAGE chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/header2.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('complex even-page header PAGE field survives round trip', async () => {
    const bytes = await buildEvenPageHeaderComplexPageFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-complex-even-page-header-page-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['page']);
    expect(result.identities.fieldInstructions).toEqual(['PAGE']);
    expect(result.identities.headerTexts).toEqual([
      'Even page 2',
      'Odd Continuing Header',
    ]);
    expect(result.identities.footerTexts).toContain('Continuing Footer');
    expect(result.identities.plainText).toContain(
      'Report body with complex even-page PAGE chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/header2.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('even-page footer NUMPAGES field survives round trip', async () => {
    const bytes = await buildEvenPageFooterNumPagesFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-even-page-footer-numpages-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['numPages']);
    expect(result.identities.fieldInstructions).toEqual(['NUMPAGES']);
    expect(result.identities.headerTexts).toContain('Odd Continuing Header');
    expect(result.identities.footerTexts).toEqual([
      'Even of 14',
      'Odd Continuing Footer',
    ]);
    expect(result.identities.plainText).toContain(
      'Report body with even-page NUMPAGES chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.exportedParts).toContain('word/footer2.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('complex even-page footer NUMPAGES field survives round trip', async () => {
    const bytes = await buildEvenPageFooterComplexNumPagesFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-complex-even-page-footer-numpages-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['numPages']);
    expect(result.identities.fieldInstructions).toEqual(['NUMPAGES']);
    expect(result.identities.headerTexts).toContain('Odd Continuing Header');
    expect(result.identities.footerTexts).toEqual([
      'Even of 14',
      'Odd Continuing Footer',
    ]);
    expect(result.identities.plainText).toContain(
      'Report body with complex even-page NUMPAGES chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.exportedParts).toContain('word/footer2.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('first-page header DATE field survives round trip', async () => {
    const bytes = await buildFirstPageHeaderDateFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-first-page-header-date-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['date']);
    expect(result.identities.fieldInstructions).toEqual([
      'DATE \\@ "yyyy-MM-dd"',
    ]);
    expect(result.identities.headerTexts).toEqual([
      'Acme Continuing Header',
      'Printed 2026-09-22',
    ]);
    expect(result.identities.footerTexts).toContain('Continuing Footer');
    expect(result.identities.plainText).toContain(
      'Report body with first-page DATE chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/header2.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('complex first-page header DATE field survives round trip', async () => {
    const bytes = await buildFirstPageHeaderComplexDateFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-complex-first-page-header-date-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['date']);
    expect(result.identities.fieldInstructions).toEqual([
      'DATE \\@ "yyyy-MM-dd"',
    ]);
    expect(result.identities.headerTexts).toEqual([
      'Acme Continuing Header',
      'Printed 2026-09-22',
    ]);
    expect(result.identities.footerTexts).toContain('Continuing Footer');
    expect(result.identities.plainText).toContain(
      'Report body with complex first-page DATE chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/header2.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('first-page footer TIME field survives round trip', async () => {
    const bytes = await buildFirstPageFooterTimeFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-first-page-footer-time-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['time']);
    expect(result.identities.fieldInstructions).toEqual(['TIME']);
    expect(result.identities.headerTexts).toContain('Acme Continuing Header');
    expect(result.identities.footerTexts).toEqual([
      'at 17:08',
      'Continuing Footer',
    ]);
    expect(result.identities.plainText).toContain(
      'Report body with first-page TIME chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.exportedParts).toContain('word/footer2.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('complex first-page footer TIME field survives round trip', async () => {
    const bytes = await buildFirstPageFooterComplexTimeFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-complex-first-page-footer-time-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['time']);
    expect(result.identities.fieldInstructions).toEqual(['TIME']);
    expect(result.identities.headerTexts).toContain('Acme Continuing Header');
    expect(result.identities.footerTexts).toEqual([
      'at 17:08',
      'Continuing Footer',
    ]);
    expect(result.identities.plainText).toContain(
      'Report body with complex first-page TIME chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.exportedParts).toContain('word/footer2.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('even-page header DATE field survives round trip', async () => {
    const bytes = await buildEvenPageHeaderDateFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-even-page-header-date-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['date']);
    expect(result.identities.fieldInstructions).toEqual([
      'DATE \\@ "yyyy-MM-dd"',
    ]);
    expect(result.identities.headerTexts).toEqual([
      'Odd Continuing Header',
      'Printed 2026-09-22',
    ]);
    expect(result.identities.footerTexts).toContain('Continuing Footer');
    expect(result.identities.plainText).toContain(
      'Report body with even-page DATE chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/header2.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('complex even-page header DATE field survives round trip', async () => {
    const bytes = await buildEvenPageHeaderComplexDateFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-complex-even-page-header-date-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['date']);
    expect(result.identities.fieldInstructions).toEqual([
      'DATE \\@ "yyyy-MM-dd"',
    ]);
    expect(result.identities.headerTexts).toEqual([
      'Odd Continuing Header',
      'Printed 2026-09-22',
    ]);
    expect(result.identities.footerTexts).toContain('Continuing Footer');
    expect(result.identities.plainText).toContain(
      'Report body with complex even-page DATE chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/header2.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('even-page footer TIME field survives round trip', async () => {
    const bytes = await buildEvenPageFooterTimeFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-even-page-footer-time-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['time']);
    expect(result.identities.fieldInstructions).toEqual(['TIME']);
    expect(result.identities.headerTexts).toContain('Odd Continuing Header');
    expect(result.identities.footerTexts).toEqual([
      'at 17:18',
      'Odd Continuing Footer',
    ]);
    expect(result.identities.plainText).toContain(
      'Report body with even-page TIME chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.exportedParts).toContain('word/footer2.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('complex even-page footer TIME field survives round trip', async () => {
    const bytes = await buildEvenPageFooterComplexTimeFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-complex-even-page-footer-time-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['time']);
    expect(result.identities.fieldInstructions).toEqual(['TIME']);
    expect(result.identities.headerTexts).toContain('Odd Continuing Header');
    expect(result.identities.footerTexts).toEqual([
      'at 17:18',
      'Odd Continuing Footer',
    ]);
    expect(result.identities.plainText).toContain(
      'Report body with complex even-page TIME chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.exportedParts).toContain('word/footer2.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('first-page header SECTION field survives round trip', async () => {
    const bytes = await buildFirstPageHeaderSectionFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-first-page-header-section-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['section']);
    expect(result.identities.fieldInstructions).toEqual(['SECTION']);
    expect(result.identities.headerTexts).toEqual([
      'Acme Continuing Header',
      'Section 2',
    ]);
    expect(result.identities.footerTexts).toContain('Continuing Footer');
    expect(result.identities.plainText).toContain(
      'Report body with first-page SECTION chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/header2.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('complex first-page header SECTION field survives round trip', async () => {
    const bytes = await buildFirstPageHeaderComplexSectionFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-complex-first-page-header-section-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['section']);
    expect(result.identities.fieldInstructions).toEqual(['SECTION']);
    expect(result.identities.headerTexts).toEqual([
      'Acme Continuing Header',
      'Section 2',
    ]);
    expect(result.identities.footerTexts).toContain('Continuing Footer');
    expect(result.identities.plainText).toContain(
      'Report body with complex first-page SECTION chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/header2.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('first-page footer SECTIONPAGES field survives round trip', async () => {
    const bytes = await buildFirstPageFooterSectionPagesFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-first-page-footer-sectionpages-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['sectionPages']);
    expect(result.identities.fieldInstructions).toEqual(['SECTIONPAGES']);
    expect(result.identities.headerTexts).toContain('Acme Continuing Header');
    expect(result.identities.footerTexts).toEqual([
      'Continuing Footer',
      'of 4',
    ]);
    expect(result.identities.plainText).toContain(
      'Report body with first-page SECTIONPAGES chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.exportedParts).toContain('word/footer2.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('complex first-page footer SECTIONPAGES field survives round trip', async () => {
    const bytes = await buildFirstPageFooterComplexSectionPagesFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-complex-first-page-footer-sectionpages-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['sectionPages']);
    expect(result.identities.fieldInstructions).toEqual(['SECTIONPAGES']);
    expect(result.identities.headerTexts).toContain('Acme Continuing Header');
    expect(result.identities.footerTexts).toEqual([
      'Continuing Footer',
      'of 4',
    ]);
    expect(result.identities.plainText).toContain(
      'Report body with complex first-page SECTIONPAGES chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.exportedParts).toContain('word/footer2.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('even-page header SECTION field survives round trip', async () => {
    const bytes = await buildEvenPageHeaderSectionFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-even-page-header-section-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['section']);
    expect(result.identities.fieldInstructions).toEqual(['SECTION']);
    expect(result.identities.headerTexts).toEqual([
      'Odd Continuing Header',
      'Section 2',
    ]);
    expect(result.identities.footerTexts).toContain('Continuing Footer');
    expect(result.identities.plainText).toContain(
      'Report body with even-page SECTION chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/header2.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('complex even-page header SECTION field survives round trip', async () => {
    const bytes = await buildEvenPageHeaderComplexSectionFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-complex-even-page-header-section-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['section']);
    expect(result.identities.fieldInstructions).toEqual(['SECTION']);
    expect(result.identities.headerTexts).toEqual([
      'Odd Continuing Header',
      'Section 2',
    ]);
    expect(result.identities.footerTexts).toContain('Continuing Footer');
    expect(result.identities.plainText).toContain(
      'Report body with complex even-page SECTION chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/header2.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('even-page footer SECTIONPAGES field survives round trip', async () => {
    const bytes = await buildEvenPageFooterSectionPagesFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-even-page-footer-sectionpages-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['sectionPages']);
    expect(result.identities.fieldInstructions).toEqual(['SECTIONPAGES']);
    expect(result.identities.headerTexts).toContain('Odd Continuing Header');
    expect(result.identities.footerTexts).toEqual([
      'Odd Continuing Footer',
      'of 4',
    ]);
    expect(result.identities.plainText).toContain(
      'Report body with even-page SECTIONPAGES chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.exportedParts).toContain('word/footer2.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('complex even-page footer SECTIONPAGES field survives round trip', async () => {
    const bytes = await buildEvenPageFooterComplexSectionPagesFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-complex-even-page-footer-sectionpages-field.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['sectionPages']);
    expect(result.identities.fieldInstructions).toEqual(['SECTIONPAGES']);
    expect(result.identities.headerTexts).toContain('Odd Continuing Header');
    expect(result.identities.footerTexts).toEqual([
      'Odd Continuing Footer',
      'of 4',
    ]);
    expect(result.identities.plainText).toContain(
      'Report body with complex even-page SECTIONPAGES chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.exportedParts).toContain('word/footer2.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
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

  test('header DATE and footer TIME fields survive round trip', async () => {
    const bytes = await buildHeaderFooterDateTimeFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-header-footer-date-time-fields.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['date', 'time']);
    expect(result.identities.fieldInstructions).toEqual([
      'DATE \\@ "yyyy-MM-dd"',
      'TIME',
    ]);
    expect(result.identities.headerTexts.join(' ')).toContain('Printed');
    expect(result.identities.footerTexts.join(' ')).toContain('at');
    expect(result.identities.plainText).toContain(
      'Report body with date/time chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('complex header DATE and footer TIME fields survive round trip', async () => {
    const bytes = await buildHeaderFooterComplexDateTimeFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-header-footer-complex-date-time-fields.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['date', 'time']);
    expect(result.identities.fieldInstructions).toEqual([
      'DATE \\@ "yyyy-MM-dd"',
      'TIME',
    ]);
    expect(result.identities.headerTexts.join(' ')).toContain('Printed');
    expect(result.identities.footerTexts.join(' ')).toContain('at');
    expect(result.identities.plainText).toContain(
      'Report body with complex date/time chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('header SECTION and footer SECTIONPAGES fields survive round trip', async () => {
    const bytes = await buildHeaderFooterSectionFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-header-footer-section-fields.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['section', 'sectionPages']);
    expect(result.identities.fieldInstructions).toEqual([
      'SECTION',
      'SECTIONPAGES',
    ]);
    expect(result.identities.headerTexts.join(' ')).toContain('Section');
    expect(result.identities.footerTexts.join(' ')).toContain('of');
    expect(result.identities.plainText).toContain(
      'Report body with section chrome',
    );
    expect(result.exportedParts).toContain('word/header1.xml');
    expect(result.exportedParts).toContain('word/footer1.xml');
    expect(result.firstPassIssues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(['docx.headers']),
    );
  });

  test('complex header SECTION and footer SECTIONPAGES fields survive round trip', async () => {
    const bytes = await buildHeaderFooterComplexSectionFieldFixture();
    const result = await roundTripDocx(
      bytes,
      'report-header-footer-complex-section-fields.docx',
    );

    expect(result.identities.fieldKinds).toEqual(['section', 'sectionPages']);
    expect(result.identities.fieldInstructions).toEqual([
      'SECTION',
      'SECTIONPAGES',
    ]);
    expect(result.identities.headerTexts.join(' ')).toContain('Section');
    expect(result.identities.footerTexts.join(' ')).toContain('of');
    expect(result.identities.plainText).toContain(
      'Report body with complex section chrome',
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
