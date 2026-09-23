import JSZip from 'jszip';
import { createArtifact, createArtifactBlob } from '../../../src/core';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CONTENT_TYPES_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/content-types';
const RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Minimal report-shaped DOCX: bookmark, external link, stable prose. */
export async function buildBookmarksAndLinksFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:bookmarkStart w:id="1" w:name="Architecture"/><w:r><w:t>Architecture overview</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p><w:p><w:hyperlink r:id="rId5"><w:r><w:t>Product site</w:t></w:r></w:hyperlink></w:p><w:sectPr/></w:body></w:document>`,
    relationships: [
      [
        'rId5',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/hyperlink`,
        'https://a3s.dev/office',
        'External',
      ],
    ],
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/** Review-shaped DOCX: insertion and deletion with a stable author. */
export async function buildReviewTrackChangesFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Keep this clause. </w:t></w:r><w:ins w:id="10" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z"><w:r><w:t>Added warranty</w:t></w:r></w:ins><w:del w:id="11" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z"><w:r><w:delText>Remove liability</w:delText></w:r></w:del></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: bounded w:rPrChange character-formatting revision identity.
 */
export async function buildCharacterFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="70" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>Formatted clause text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: jc-only w:pPrChange paragraph-formatting revision identity.
 */
export async function buildParagraphFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="71" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>Aligned paragraph clause</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: whole-paragraph mark insertion revision identity.
 * Built through the native HTML→DOCX path so import→export→reopen matches the
 * browser paragraph-mark body contract (tracked wraps + block metadata).
 */
export async function buildParagraphMarkInsertionRevisionFixture(): Promise<Uint8Array> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  artifact.content.html = [
    '<section data-document-section="true">',
    '<p data-document-block-change="true" data-block-change-kind="insertion" data-block-change-id="block-insertion" data-block-change-author="Reviewer" data-block-change-date="2026-09-08T00:00:00.000Z">',
    '<ins data-document-change="true" data-change-kind="insertion" data-change-id="text-insertion" data-change-author="Reviewer" data-change-date="2026-09-08T00:00:00.000Z">Added whole paragraph</ins>',
    '</p><p>Stable paragraph</p></section>',
  ].join('');
  artifact.content.trackChanges = true;
  const blob = await createArtifactBlob(artifact);
  return new Uint8Array(await blob.arrayBuffer());
}

/**
 * Contract-shaped DOCX: reviewable paragraph-break split revision identity.
 */
export async function buildParagraphBreakSplitRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Before split clause</w:t></w:r></w:p><w:p><w:pPr><w:rPr><w:ins w:id="82" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"/></w:rPr></w:pPr><w:r><w:t>After split clause</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: paired w:moveFrom/w:moveTo text-move revision identity.
 */
export async function buildTextMoveRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:moveFrom w:id="83" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:r><w:delText>Moved clause</w:delText></w:r></w:moveFrom><w:r><w:t> middle </w:t></w:r><w:moveTo w:id="83" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:r><w:t>Moved clause</w:t></w:r></w:moveTo></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: paired move wrappers plus companion move-range bookmarks.
 */
export async function buildTextMoveRangeCompanionRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:moveFromRangeStart w:id="0" w:author="Reviewer" w:date="2026-09-08T00:00:00Z" w:name="move0"/><w:moveFrom w:id="84" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:r><w:delText>Ranged move clause</w:delText></w:r></w:moveFrom><w:moveFromRangeEnd w:id="0"/><w:r><w:t> spacer </w:t></w:r><w:moveToRangeStart w:id="0" w:author="Reviewer" w:date="2026-09-08T00:00:00Z" w:name="move0"/><w:moveTo w:id="84" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:r><w:t>Ranged move clause</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="0"/></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: jc-only w:tblPrChange table-formatting revision identity.
 */
export async function buildTableFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tblPr><w:jc w:val="right"/><w:tblPrChange w:id="11" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tblPr><w:jc w:val="center"/></w:tblPr></w:tblPrChange></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Party obligation</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tblW-only w:tblPrChange preferred-width revision identity.
 */
export async function buildTableWidthFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tblPr><w:tblW w:type="pct" w:w="5000"/><w:tblPrChange w:id="12" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tblPr><w:tblW w:type="pct" w:w="2500"/></w:tblPr></w:tblPrChange></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Preferred width clause</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tblInd-only w:tblPrChange table-indent revision identity.
 */
export async function buildTableIndentFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tblPr><w:tblInd w:type="dxa" w:w="1440"/><w:tblPrChange w:id="14" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tblPr><w:tblInd w:type="dxa" w:w="720"/></w:tblPr></w:tblPrChange></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Indented table clause</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tblCellMar-only w:tblPrChange table cell-margin revision identity.
 */
export async function buildTableCellMarginFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tblPr><w:tblCellMar><w:top w:type="dxa" w:w="120"/><w:left w:type="dxa" w:w="120"/><w:bottom w:type="dxa" w:w="120"/><w:right w:type="dxa" w:w="120"/></w:tblCellMar><w:tblPrChange w:id="15" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tblPr><w:tblCellMar><w:top w:type="dxa" w:w="0"/><w:left w:type="dxa" w:w="0"/><w:bottom w:type="dxa" w:w="0"/><w:right w:type="dxa" w:w="0"/></w:tblCellMar></w:tblPr></w:tblPrChange></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Padded table clause</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tblLayout-only w:tblPrChange table-layout revision identity.
 */
export async function buildTableLayoutFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tblPr><w:tblLayout w:type="fixed"/><w:tblPrChange w:id="16" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tblPr><w:tblLayout w:type="autofit"/></w:tblPr></w:tblPrChange></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Fixed layout clause</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: bidiVisual-only w:tblPrChange RTL visual revision identity.
 */
export async function buildTableBidiVisualFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tblPr><w:bidiVisual/><w:tblPrChange w:id="17" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tblPr><w:bidiVisual w:val="0"/></w:tblPr></w:tblPrChange></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Bidi visual clause</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: solid-shd-only w:tblPrChange table-fill revision identity.
 */
export async function buildTableFillFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tblPr><w:shd w:val="clear" w:fill="FFFFFF"/><w:tblPrChange w:id="18" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tblPr><w:shd w:val="clear" w:fill="FFCC00"/></w:tblPr></w:tblPrChange></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Filled table clause</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tblOverlap-only w:tblPrChange overlap revision identity.
 */
export async function buildTableOverlapFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tblPr><w:tblOverlap w:val="overlap"/><w:tblPrChange w:id="19" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tblPr><w:tblOverlap w:val="never"/></w:tblPr></w:tblPrChange></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Overlap table clause</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tblLook-only w:tblPrChange table-look revision identity.
 */
export async function buildTableLookFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tblPr><w:tblLook w:val="00A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/><w:tblPrChange w:id="20" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tblPr><w:tblLook w:val="0020" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/></w:tblPr></w:tblPrChange></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Look table clause</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tblStyle-only w:tblPrChange style-id revision identity.
 */
export async function buildTableStyleFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblPrChange w:id="21" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tblPr><w:tblStyle w:val="TableNormal"/></w:tblPr></w:tblPrChange></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Styled table clause</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
    stylesXml: `<w:styles xmlns:w="${WORD_NAMESPACE}"><w:style w:type="table" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style></w:styles>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tblCellSpacing-only w:tblPrChange cell-spacing revision identity.
 */
export async function buildTableCellSpacingFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tblPr><w:tblCellSpacing w:type="dxa" w:w="240"/><w:tblPrChange w:id="22" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tblPr><w:tblCellSpacing w:type="dxa" w:w="120"/></w:tblPr></w:tblPrChange></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Spaced table clause</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tblBorders-only w:tblPrChange table-border revision identity.
 */
export async function buildTableBordersFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tblPr><w:tblBorders><w:top w:val="double" w:sz="24" w:space="0" w:color="0000FF"/></w:tblBorders><w:tblPrChange w:id="23" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="FF0000"/></w:tblBorders></w:tblPr></w:tblPrChange></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Bordered table clause</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tblStyleColBandSize-only w:tblPrChange column-band revision identity.
 */
export async function buildTableColBandSizeFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tblPr><w:tblStyleColBandSize w:val="2"/><w:tblPrChange w:id="24" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tblPr><w:tblStyleColBandSize w:val="1"/></w:tblPr></w:tblPrChange></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Col-band table clause</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tblStyleRowBandSize-only w:tblPrChange row-band revision identity.
 */
export async function buildTableRowBandSizeFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tblPr><w:tblStyleRowBandSize w:val="2"/><w:tblPrChange w:id="25" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tblPr><w:tblStyleRowBandSize w:val="1"/></w:tblPr></w:tblPrChange></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Row-band table clause</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tblCaption-only w:tblPrChange caption revision identity.
 */
export async function buildTableCaptionFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tblPr><w:tblCaption w:val="Annual summary"/><w:tblPrChange w:id="26" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tblPr><w:tblCaption w:val="Quarterly summary"/></w:tblPr></w:tblPrChange></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Captioned table clause</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tblDescription-only w:tblPrChange description revision identity.
 */
export async function buildTableDescriptionFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tblPr><w:tblDescription w:val="Annual summary"/><w:tblPrChange w:id="27" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tblPr><w:tblDescription w:val="Quarterly summary"/></w:tblPr></w:tblPrChange></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Described table clause</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tblpPr-only w:tblPrChange float-position revision identity.
 */
export async function buildTableFloatFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tblPr><w:tblpPr w:horzAnchor="margin" w:vertAnchor="page" w:tblpX="1440" w:tblpY="720"/><w:tblPrChange w:id="28" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tblPr><w:tblpPr w:horzAnchor="page" w:vertAnchor="text" w:tblpX="720" w:tblpY="360" w:leftFromText="120" w:rightFromText="120" w:topFromText="0" w:bottomFromText="0"/></w:tblPr></w:tblPrChange></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Floating table clause</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: orientation-only w:sectPrChange section-formatting revision identity.
 */
export async function buildSectionFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with section orientation revision</w:t></w:r></w:p><w:sectPr><w:pgSz w:orient="portrait"/><w:sectPrChange w:id="41" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:pgSz w:orient="landscape"/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: pgMar-only w:sectPrChange section-margin revision identity.
 */
export async function buildSectionMarginFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with section margin revision</w:t></w:r></w:p><w:sectPr><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="708" w:footer="708" w:gutter="0"/><w:sectPrChange w:id="42" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: titlePg-only w:sectPrChange title-page revision identity.
 */
export async function buildSectionTitlePageFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with title-page revision</w:t></w:r></w:p><w:sectPr><w:titlePg/><w:sectPrChange w:id="43" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:titlePg w:val="0"/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: equal-width cols-only w:sectPrChange column revision identity.
 */
export async function buildSectionEqualColsFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with equal columns revision</w:t></w:r></w:p><w:sectPr><w:cols w:num="2" w:space="720"/><w:sectPrChange w:id="44" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:cols w:num="1" w:space="720"/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: rtlGutter-only w:sectPrChange RTL gutter revision identity.
 */
export async function buildSectionRtlGutterFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with RTL gutter revision</w:t></w:r></w:p><w:sectPr><w:rtlGutter/><w:sectPrChange w:id="45" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:rtlGutter w:val="0"/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: paperSrc-only w:sectPrChange paper-source revision identity.
 */
export async function buildSectionPaperSourceFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with paper-source revision</w:t></w:r></w:p><w:sectPr><w:paperSrc w:first="5" w:other="6"/><w:sectPrChange w:id="46" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:paperSrc w:first="1" w:other="2"/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: vAlign-only w:sectPrChange vertical-align revision identity.
 */
export async function buildSectionVerticalAlignFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with section vertical-align revision</w:t></w:r></w:p><w:sectPr><w:vAlign w:val="center"/><w:sectPrChange w:id="47" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:vAlign w:val="top"/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: docGrid-only w:sectPrChange document-grid revision identity.
 */
export async function buildSectionDocGridFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with document-grid revision</w:t></w:r></w:p><w:sectPr><w:docGrid w:type="linesAndChars" w:linePitch="480"/><w:sectPrChange w:id="48" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: bidi-only w:sectPrChange section bidi revision identity.
 */
export async function buildSectionBidiFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with section bidi revision</w:t></w:r></w:p><w:sectPr><w:bidi w:val="0"/><w:sectPrChange w:id="49" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:bidi/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: type-only w:sectPrChange section-break revision identity.
 */
export async function buildSectionTypeFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with section-break type revision</w:t></w:r></w:p><w:sectPr><w:type w:val="nextPage"/><w:sectPrChange w:id="50" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:type w:val="continuous"/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: complete pgSz geometry-only w:sectPrChange page-size revision identity.
 */
export async function buildSectionPageGeometryFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with page-geometry revision</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:sectPrChange w:id="51" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: unequal-width cols-only w:sectPrChange column revision identity.
 */
export async function buildSectionUnequalColsFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with unequal columns revision</w:t></w:r></w:p><w:sectPr><w:cols w:num="2" w:equalWidth="0" w:sep="1"><w:col w:w="3600" w:space="720"/><w:col w:w="7200"/></w:cols><w:sectPrChange w:id="52" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:cols w:num="2" w:equalWidth="0"><w:col w:w="2880" w:space="720"/><w:col w:w="5760"/></w:cols></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: lnNumType-only w:sectPrChange line-numbering revision identity.
 */
export async function buildSectionLnNumTypeFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with line-numbering revision</w:t></w:r></w:p><w:sectPr><w:lnNumType w:countBy="2" w:start="5" w:distance="720" w:restart="newSection"/><w:sectPrChange w:id="53" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:lnNumType w:countBy="1" w:start="1" w:restart="newPage"/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: pgNumType-only w:sectPrChange page-number format revision identity.
 */
export async function buildSectionPgNumTypeFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with page-number format revision</w:t></w:r></w:p><w:sectPr><w:pgNumType w:fmt="upperRoman" w:start="5"/><w:sectPrChange w:id="54" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:pgNumType w:fmt="decimal" w:start="1"/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: formProt-only w:sectPrChange form-protection revision identity.
 */
export async function buildSectionFormProtFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with form-protection revision</w:t></w:r></w:p><w:sectPr><w:formProt/><w:sectPrChange w:id="55" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:formProt w:val="0"/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: noEndnote-only w:sectPrChange endnote-suppression revision identity.
 */
export async function buildSectionNoEndnoteFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with no-endnote revision</w:t></w:r></w:p><w:sectPr><w:noEndnote/><w:sectPrChange w:id="56" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:noEndnote w:val="0"/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: textDirection-only w:sectPrChange section text-direction revision identity.
 */
export async function buildSectionTextDirectionFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with section text-direction revision</w:t></w:r></w:p><w:sectPr><w:textDirection w:val="tbRl"/><w:sectPrChange w:id="57" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:textDirection w:val="lrTb"/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: footnotePr-only w:sectPrChange footnote-properties revision identity.
 */
export async function buildSectionFootnotePrFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with footnote-properties revision</w:t></w:r></w:p><w:sectPr><w:footnotePr><w:pos w:val="beneathText"/><w:numFmt w:val="upperRoman"/><w:numStart w:val="5"/><w:numRestart w:val="continuous"/></w:footnotePr><w:sectPrChange w:id="58" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:footnotePr><w:pos w:val="pageBottom"/><w:numFmt w:val="decimal"/><w:numStart w:val="1"/><w:numRestart w:val="eachSect"/></w:footnotePr></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: endnotePr-only w:sectPrChange endnote-properties revision identity.
 */
export async function buildSectionEndnotePrFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with endnote-properties revision</w:t></w:r></w:p><w:sectPr><w:endnotePr><w:pos w:val="docEnd"/><w:numFmt w:val="upperRoman"/><w:numStart w:val="5"/><w:numRestart w:val="continuous"/></w:endnotePr><w:sectPrChange w:id="59" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:endnotePr><w:pos w:val="sectEnd"/><w:numFmt w:val="decimal"/><w:numStart w:val="1"/><w:numRestart w:val="eachSect"/></w:endnotePr></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: pgBorders-only w:sectPrChange page-border revision identity.
 */
export async function buildSectionPageBordersFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with page-borders revision</w:t></w:r></w:p><w:sectPr><w:pgBorders w:display="firstPage" w:offsetFrom="text" w:zOrder="back"><w:left w:val="double" w:sz="12" w:space="4" w:color="0000FF"/><w:right w:val="double" w:sz="12" w:space="4" w:color="0000FF"/></w:pgBorders><w:sectPrChange w:id="60" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:sectPr><w:pgBorders w:display="allPages" w:offsetFrom="page" w:zOrder="front"><w:top w:val="single" w:sz="24" w:space="24" w:color="FF0000"/><w:bottom w:val="single" w:sz="24" w:space="24" w:color="FF0000"/></w:pgBorders></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: cantSplit-only w:trPrChange row-formatting revision identity.
 */
export async function buildRowFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:trPr><w:cantSplit/><w:trPrChange w:id="21" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:trPr><w:cantSplit w:val="0"/></w:trPr></w:trPrChange></w:trPr><w:tc><w:p><w:r><w:t>Keep row together</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tblHeader-only w:trPrChange header-row revision identity.
 */
export async function buildRowHeaderFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:trPr><w:tblHeader/><w:trPrChange w:id="22" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:trPr><w:tblHeader w:val="0"/></w:trPr></w:trPrChange></w:trPr><w:tc><w:p><w:r><w:t>Repeat header row</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: trHeight-only w:trPrChange row-height revision identity.
 */
export async function buildRowHeightFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:trPr><w:trHeight w:val="720" w:hRule="atLeast"/><w:trPrChange w:id="23" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:trPr><w:trHeight w:val="480" w:hRule="exact"/></w:trPr></w:trPrChange></w:trPr><w:tc><w:p><w:r><w:t>Fixed row height</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: hidden-only w:trPrChange row-visibility revision identity.
 */
export async function buildRowHiddenFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:trPr><w:hidden/><w:trPrChange w:id="24" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:trPr><w:hidden w:val="0"/></w:trPr></w:trPrChange></w:trPr><w:tc><w:p><w:r><w:t>Hidden contract row</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: jc-only w:trPrChange row-alignment revision identity.
 */
export async function buildRowAlignmentFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:trPr><w:jc w:val="left"/><w:trPrChange w:id="25" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:trPr><w:jc w:val="center"/></w:trPr></w:trPrChange></w:trPr><w:tc><w:p><w:r><w:t>Aligned contract row</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: gridBefore-only w:trPrChange leading-grid revision identity.
 */
export async function buildRowGridBeforeFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:trPr><w:gridBefore w:val="1"/><w:trPrChange w:id="26" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:trPr><w:gridBefore w:val="2"/></w:trPr></w:trPrChange></w:trPr><w:tc><w:p><w:r><w:t>Leading grid skip</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: gridAfter-only w:trPrChange trailing-grid revision identity.
 */
export async function buildRowGridAfterFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:trPr><w:gridAfter w:val="1"/><w:trPrChange w:id="27" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:trPr><w:gridAfter w:val="2"/></w:trPr></w:trPrChange></w:trPr><w:tc><w:p><w:r><w:t>Trailing grid skip</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: wBefore-only w:trPrChange leading-width revision identity.
 */
export async function buildRowWidthBeforeFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:trPr><w:wBefore w:type="dxa" w:w="144"/><w:trPrChange w:id="28" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:trPr><w:wBefore w:type="dxa" w:w="288"/></w:trPr></w:trPrChange></w:trPr><w:tc><w:p><w:r><w:t>Leading row indent</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: wAfter-only w:trPrChange trailing-width revision identity.
 */
export async function buildRowWidthAfterFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:trPr><w:wAfter w:type="dxa" w:w="144"/><w:trPrChange w:id="29" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:trPr><w:wAfter w:type="dxa" w:w="288"/></w:trPr></w:trPrChange></w:trPr><w:tc><w:p><w:r><w:t>Trailing row indent</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: cnfStyle-only w:trPrChange conditional-format revision identity.
 */
export async function buildRowCnfStyleFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:trPr><w:cnfStyle w:val="010000000000"/><w:trPrChange w:id="30" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:trPr><w:cnfStyle w:val="100000000000"/></w:trPr></w:trPrChange></w:trPr><w:tc><w:p><w:r><w:t>Conditional format row</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: divId-only w:trPrChange HTML-div revision identity.
 */
export async function buildRowDivIdFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:trPr><w:divId w:val="87654321"/><w:trPrChange w:id="31" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:trPr><w:divId w:val="12345678"/></w:trPr></w:trPrChange></w:trPr><w:tc><w:p><w:r><w:t>Div-mapped contract row</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tblCellSpacing-only w:trPrChange cell-spacing revision identity.
 */
export async function buildRowCellSpacingFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:trPr><w:tblCellSpacing w:type="dxa" w:w="144"/><w:trPrChange w:id="32" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:trPr><w:tblCellSpacing w:type="dxa" w:w="288"/></w:trPr></w:trPrChange></w:trPr><w:tc><w:p><w:r><w:t>Spaced contract row</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: vAlign-only w:tcPrChange cell-formatting revision identity.
 */
export async function buildCellFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:tc><w:tcPr><w:vAlign w:val="bottom"/><w:tcPrChange w:id="31" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tcPr><w:vAlign w:val="center"/></w:tcPr></w:tcPrChange></w:tcPr><w:p><w:r><w:t>Aligned cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: solid-shd-only w:tcPrChange cell-fill revision identity.
 */
export async function buildCellFillFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:tc><w:tcPr><w:shd w:val="clear" w:fill="FFFFFF"/><w:tcPrChange w:id="33" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tcPr><w:shd w:val="clear" w:fill="FFAA00"/></w:tcPr></w:tcPrChange></w:tcPr><w:p><w:r><w:t>Filled contract cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tcW-only w:tcPrChange preferred-width revision identity.
 */
export async function buildCellWidthFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2880"/><w:tcPrChange w:id="34" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tcPr><w:tcW w:type="dxa" w:w="1440"/></w:tcPr></w:tcPrChange></w:tcPr><w:p><w:r><w:t>Width contract cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tcMar-only w:tcPrChange cell-margin revision identity.
 */
export async function buildCellMarginFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:tc><w:tcPr><w:tcMar><w:top w:type="dxa" w:w="120"/><w:left w:type="dxa" w:w="120"/><w:bottom w:type="dxa" w:w="120"/><w:right w:type="dxa" w:w="120"/></w:tcMar><w:tcPrChange w:id="35" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tcPr><w:tcMar><w:top w:type="dxa" w:w="0"/><w:left w:type="dxa" w:w="0"/><w:bottom w:type="dxa" w:w="0"/><w:right w:type="dxa" w:w="0"/></w:tcMar></w:tcPr></w:tcPrChange></w:tcPr><w:p><w:r><w:t>Margin contract cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: noWrap-only w:tcPrChange cell wrap revision identity.
 */
export async function buildCellNoWrapFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:tc><w:tcPr><w:noWrap/><w:tcPrChange w:id="36" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tcPr><w:noWrap w:val="0"/></w:tcPr></w:tcPrChange></w:tcPr><w:p><w:r><w:t>No-wrap contract cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: textDirection-only w:tcPrChange text-direction revision identity.
 */
export async function buildCellTextDirectionFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:tc><w:tcPr><w:textDirection w:val="lrTb"/><w:tcPrChange w:id="37" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tcPr><w:textDirection w:val="btLr"/></w:tcPr></w:tcPrChange></w:tcPr><w:p><w:r><w:t>Direction contract cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tcFitText-only w:tcPrChange fit-text revision identity.
 */
export async function buildCellFitTextFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:tc><w:tcPr><w:tcFitText/><w:tcPrChange w:id="38" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tcPr><w:tcFitText w:val="0"/></w:tcPr></w:tcPrChange></w:tcPr><w:p><w:r><w:t>Fit-text contract cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: hideMark-only w:tcPrChange hide-mark revision identity.
 */
export async function buildCellHideMarkFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:tc><w:tcPr><w:hideMark/><w:tcPrChange w:id="39" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tcPr><w:hideMark w:val="0"/></w:tcPr></w:tcPrChange></w:tcPr><w:p><w:r><w:t>Hide-mark contract cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: cnfStyle-only w:tcPrChange conditional-format revision identity.
 */
export async function buildCellCnfStyleFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:tc><w:tcPr><w:cnfStyle w:val="010000000000"/><w:tcPrChange w:id="40" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tcPr><w:cnfStyle w:val="100000000000"/></w:tcPr></w:tcPrChange></w:tcPr><w:p><w:r><w:t>Conditional format cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: hMerge-only w:tcPrChange horizontal-merge revision identity.
 */
export async function buildCellHMergeFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:tc><w:tcPr><w:hMerge w:val="continue"/><w:tcPrChange w:id="41" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tcPr><w:hMerge w:val="restart"/></w:tcPr></w:tcPrChange></w:tcPr><w:p><w:r><w:t>H-merge contract cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: vMerge-only w:tcPrChange vertical-merge revision identity.
 */
export async function buildCellVMergeFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:tc><w:tcPr><w:vMerge w:val="continue"/><w:tcPrChange w:id="42" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tcPr><w:vMerge w:val="restart"/></w:tcPr></w:tcPrChange></w:tcPr><w:p><w:r><w:t>V-merge contract cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: gridSpan-only w:tcPrChange column-span revision identity.
 */
export async function buildCellGridSpanFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="3"/><w:tcPrChange w:id="43" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tcPr><w:gridSpan w:val="2"/></w:tcPr></w:tcPrChange></w:tcPr><w:p><w:r><w:t>Grid-span contract cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: tcBorders-only w:tcPrChange cell-border revision identity.
 */
export async function buildCellBordersFormattingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:tc><w:tcPr><w:tcBorders><w:top w:val="double" w:sz="24" w:space="0" w:color="0000FF"/></w:tcBorders><w:tcPrChange w:id="44" w:author="Reviewer" w:date="2026-09-08T00:00:00Z"><w:tcPr><w:tcBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="FF0000"/></w:tcBorders></w:tcPr></w:tcPrChange></w:tcPr><w:p><w:r><w:t>Bordered contract cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: current-level bullet w:numberingChange (nfc 23) identity.
 */
export async function buildBulletNumberingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="55"/><w:numberingChange w:id="41" w:author="Ada Reviewer" w:date="2026-09-01T09:30:00Z" w:original="%1:1:23:•"/></w:numPr></w:pPr><w:r><w:t>First obligation</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="55"/><w:numberingChange w:id="42" w:author="Ada Reviewer" w:date="2026-09-01T09:30:00Z" w:original="%1:2:23:•"/></w:numPr></w:pPr><w:r><w:t>Second obligation</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
    numberingXml: `<w:numbering xmlns:w="${WORD_NAMESPACE}"><w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="○"/></w:lvl></w:abstractNum><w:num w:numId="55"><w:abstractNumId w:val="7"/></w:num></w:numbering>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: current-level decimal w:numberingChange (nfc 0) identity.
 */
export async function buildDecimalNumberingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="42"/><w:numberingChange w:id="17" w:author="Ada Reviewer" w:date="2026-09-01T09:30:00Z" w:original="%1:3:1:."/></w:numPr></w:pPr><w:r><w:t>First deliverable</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="42"/><w:numberingChange w:id="18" w:author="Ada Reviewer" w:date="2026-09-01T09:30:00Z" w:original="%1:4:1:."/></w:numPr></w:pPr><w:r><w:t>Second deliverable</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
    numberingXml: `<w:numbering xmlns:w="${WORD_NAMESPACE}"><w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="42"><w:abstractNumId w:val="7"/></w:num></w:numbering>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: multi-level w:numberingChange with sibling original levels.
 */
export async function buildMultiLevelNumberingRevisionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/><w:numberingChange w:id="21" w:author="Ada Reviewer" w:date="2026-09-01T09:30:00Z" w:original="%1:3:0:.%2:1:4:)"/></w:numPr></w:pPr><w:r><w:t>First milestone</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/><w:numberingChange w:id="22" w:author="Ada Reviewer" w:date="2026-09-01T09:30:00Z" w:original="%1:4:0:.%2:1:4:)"/></w:numPr></w:pPr><w:r><w:t>Second milestone</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
    numberingXml: `<w:numbering xmlns:w="${WORD_NAMESPACE}"><w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl></w:abstractNum><w:num w:numId="9"><w:abstractNumId w:val="7"/></w:num></w:numbering>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/** Contract-shaped DOCX: simple table identities. */
export async function buildContractTableFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Party A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Party B</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>Obligation</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Payment term</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/** Duplicate bookmark names must diagnose identity normalization. */
export async function buildDuplicateBookmarkFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:bookmarkStart w:id="1" w:name="Target"/><w:r><w:t>First</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p><w:p><w:bookmarkStart w:id="2" w:name="Target"/><w:r><w:t>Second</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: internal hyperlink keeps Word anchor semantics to a
 * body bookmark (no external relationship).
 */
export async function buildInternalBookmarkLinkFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:bookmarkStart w:id="7" w:name="Obligations"/><w:r><w:t>Obligations section</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p><w:p><w:hyperlink w:anchor="Obligations" w:tooltip="Jump to obligations"><w:r><w:t>See obligations</w:t></w:r></w:hyperlink></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Review-shaped DOCX: one anchored comment with stable author and text.
 */
export async function buildReviewCommentsFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Liability clause</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p><w:sectPr/></w:body></w:document>`,
    commentsXml: `<w:comments xmlns:w="${WORD_NAMESPACE}"><w:comment w:id="0" w:author="Bea Counsel" w:date="2026-09-17T02:00:00Z" w:initials="BC"><w:p><w:r><w:t>Clarify liability cap</w:t></w:r></w:p></w:comment></w:comments>`,
    documentRelationships: [
      [
        'rIdComments',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/comments`,
        'comments.xml',
      ],
    ],
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: default header and footer text identity.
 */
export async function buildHeaderFooterFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body clause</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Acme Report Header</w:t></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Confidential Footer</w:t></w:r></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: distinct first-page vs default header/footer text.
 */
export async function buildFirstPageHeaderFooterFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with first-page chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:headerReference w:type="first" r:id="rIdFirstHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:footerReference w:type="first" r:id="rIdFirstFooter"/><w:titlePg/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      [
        'rIdFirstHeader',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`,
        'header2.xml',
      ],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
      [
        'rIdFirstFooter',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`,
        'footer2.xml',
      ],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Acme Continuing Header</w:t></w:r></w:p></w:hdr>`,
    firstHeaderXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Acme Title Header</w:t></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Continuing Footer</w:t></w:r></w:p></w:ftr>`,
    firstFooterXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Title Footer</w:t></w:r></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: first-page header PAGE field vs default header text.
 */
export async function buildFirstPageHeaderPageFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with first-page PAGE chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:headerReference w:type="first" r:id="rIdFirstHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:titlePg/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      [
        'rIdFirstHeader',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`,
        'header2.xml',
      ],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Acme Continuing Header</w:t></w:r></w:p></w:hdr>`,
    firstHeaderXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Title page </w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Continuing Footer</w:t></w:r></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: complex (fldChar) first-page header PAGE vs default header.
 */
export async function buildFirstPageHeaderComplexPageFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with complex first-page PAGE chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:headerReference w:type="first" r:id="rIdFirstHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:titlePg/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      [
        'rIdFirstHeader',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`,
        'header2.xml',
      ],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Acme Continuing Header</w:t></w:r></w:p></w:hdr>`,
    firstHeaderXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Title page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Continuing Footer</w:t></w:r></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: first-page footer NUMPAGES field vs default footer text.
 */
export async function buildFirstPageFooterNumPagesFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with first-page NUMPAGES chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:footerReference w:type="first" r:id="rIdFirstFooter"/><w:titlePg/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
      [
        'rIdFirstFooter',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`,
        'footer2.xml',
      ],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Acme Continuing Header</w:t></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Continuing Footer</w:t></w:r></w:p></w:ftr>`,
    firstFooterXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Title of </w:t></w:r><w:fldSimple w:instr="NUMPAGES"><w:r><w:t>12</w:t></w:r></w:fldSimple></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: complex (fldChar) first-page footer NUMPAGES vs default footer.
 */
export async function buildFirstPageFooterComplexNumPagesFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with complex first-page NUMPAGES chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:footerReference w:type="first" r:id="rIdFirstFooter"/><w:titlePg/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
      [
        'rIdFirstFooter',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`,
        'footer2.xml',
      ],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Acme Continuing Header</w:t></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Continuing Footer</w:t></w:r></w:p></w:ftr>`,
    firstFooterXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Title of </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>12</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: even-page header PAGE field vs default header text.
 */
export async function buildEvenPageHeaderPageFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with even-page PAGE chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:headerReference w:type="even" r:id="rIdEvenHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      [
        'rIdEvenHeader',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`,
        'header2.xml',
      ],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Continuing Header</w:t></w:r></w:p></w:hdr>`,
    firstHeaderXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Even page </w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>2</w:t></w:r></w:fldSimple></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Continuing Footer</w:t></w:r></w:p></w:ftr>`,
    settingsXml: `<w:settings xmlns:w="${WORD_NAMESPACE}"><w:evenAndOddHeaders/></w:settings>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: complex (fldChar) even-page header PAGE vs default header.
 */
export async function buildEvenPageHeaderComplexPageFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with complex even-page PAGE chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:headerReference w:type="even" r:id="rIdEvenHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      [
        'rIdEvenHeader',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`,
        'header2.xml',
      ],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Continuing Header</w:t></w:r></w:p></w:hdr>`,
    firstHeaderXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Even page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Continuing Footer</w:t></w:r></w:p></w:ftr>`,
    settingsXml: `<w:settings xmlns:w="${WORD_NAMESPACE}"><w:evenAndOddHeaders/></w:settings>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: even-page footer NUMPAGES field vs default footer text.
 */
export async function buildEvenPageFooterNumPagesFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with even-page NUMPAGES chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:footerReference w:type="even" r:id="rIdEvenFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
      [
        'rIdEvenFooter',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`,
        'footer2.xml',
      ],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Continuing Header</w:t></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Continuing Footer</w:t></w:r></w:p></w:ftr>`,
    firstFooterXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Even of </w:t></w:r><w:fldSimple w:instr="NUMPAGES"><w:r><w:t>14</w:t></w:r></w:fldSimple></w:p></w:ftr>`,
    settingsXml: `<w:settings xmlns:w="${WORD_NAMESPACE}"><w:evenAndOddHeaders/></w:settings>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: complex (fldChar) even-page footer NUMPAGES vs default footer.
 */
export async function buildEvenPageFooterComplexNumPagesFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with complex even-page NUMPAGES chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:footerReference w:type="even" r:id="rIdEvenFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
      [
        'rIdEvenFooter',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`,
        'footer2.xml',
      ],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Continuing Header</w:t></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Continuing Footer</w:t></w:r></w:p></w:ftr>`,
    firstFooterXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Even of </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>14</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`,
    settingsXml: `<w:settings xmlns:w="${WORD_NAMESPACE}"><w:evenAndOddHeaders/></w:settings>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: first-page header DATE (\@ format) vs default header text.
 */
export async function buildFirstPageHeaderDateFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with first-page DATE chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:headerReference w:type="first" r:id="rIdFirstHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:titlePg/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      [
        'rIdFirstHeader',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`,
        'header2.xml',
      ],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Acme Continuing Header</w:t></w:r></w:p></w:hdr>`,
    firstHeaderXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Printed </w:t></w:r><w:fldSimple w:instr="DATE \\@ &quot;yyyy-MM-dd&quot;"><w:r><w:t>2026-09-22</w:t></w:r></w:fldSimple></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Continuing Footer</w:t></w:r></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: complex (fldChar) first-page header DATE (\@ format).
 */
export async function buildFirstPageHeaderComplexDateFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with complex first-page DATE chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:headerReference w:type="first" r:id="rIdFirstHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:titlePg/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      [
        'rIdFirstHeader',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`,
        'header2.xml',
      ],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Acme Continuing Header</w:t></w:r></w:p></w:hdr>`,
    firstHeaderXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Printed </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE \\@ &quot;yyyy-MM-dd&quot; </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2026-09-22</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Continuing Footer</w:t></w:r></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: first-page footer TIME vs default footer text.
 */
export async function buildFirstPageFooterTimeFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with first-page TIME chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:footerReference w:type="first" r:id="rIdFirstFooter"/><w:titlePg/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
      [
        'rIdFirstFooter',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`,
        'footer2.xml',
      ],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Acme Continuing Header</w:t></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Continuing Footer</w:t></w:r></w:p></w:ftr>`,
    firstFooterXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>at </w:t></w:r><w:fldSimple w:instr="TIME"><w:r><w:t>17:08</w:t></w:r></w:fldSimple></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: complex (fldChar) first-page footer TIME vs default footer.
 */
export async function buildFirstPageFooterComplexTimeFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with complex first-page TIME chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:footerReference w:type="first" r:id="rIdFirstFooter"/><w:titlePg/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
      [
        'rIdFirstFooter',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`,
        'footer2.xml',
      ],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Acme Continuing Header</w:t></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Continuing Footer</w:t></w:r></w:p></w:ftr>`,
    firstFooterXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>at </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TIME </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>17:08</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: even-page header DATE (\@ format) vs default header text.
 */
export async function buildEvenPageHeaderDateFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with even-page DATE chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:headerReference w:type="even" r:id="rIdEvenHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      [
        'rIdEvenHeader',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`,
        'header2.xml',
      ],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Continuing Header</w:t></w:r></w:p></w:hdr>`,
    firstHeaderXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Printed </w:t></w:r><w:fldSimple w:instr="DATE \\@ &quot;yyyy-MM-dd&quot;"><w:r><w:t>2026-09-22</w:t></w:r></w:fldSimple></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Continuing Footer</w:t></w:r></w:p></w:ftr>`,
    settingsXml: `<w:settings xmlns:w="${WORD_NAMESPACE}"><w:evenAndOddHeaders/></w:settings>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: complex (fldChar) even-page header DATE (\@ format).
 */
export async function buildEvenPageHeaderComplexDateFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with complex even-page DATE chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:headerReference w:type="even" r:id="rIdEvenHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      [
        'rIdEvenHeader',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`,
        'header2.xml',
      ],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Continuing Header</w:t></w:r></w:p></w:hdr>`,
    firstHeaderXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Printed </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE \\@ &quot;yyyy-MM-dd&quot; </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2026-09-22</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Continuing Footer</w:t></w:r></w:p></w:ftr>`,
    settingsXml: `<w:settings xmlns:w="${WORD_NAMESPACE}"><w:evenAndOddHeaders/></w:settings>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: even-page footer TIME vs default footer text.
 */
export async function buildEvenPageFooterTimeFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with even-page TIME chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:footerReference w:type="even" r:id="rIdEvenFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
      [
        'rIdEvenFooter',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`,
        'footer2.xml',
      ],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Continuing Header</w:t></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Continuing Footer</w:t></w:r></w:p></w:ftr>`,
    firstFooterXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>at </w:t></w:r><w:fldSimple w:instr="TIME"><w:r><w:t>17:18</w:t></w:r></w:fldSimple></w:p></w:ftr>`,
    settingsXml: `<w:settings xmlns:w="${WORD_NAMESPACE}"><w:evenAndOddHeaders/></w:settings>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: complex (fldChar) even-page footer TIME vs default footer.
 */
export async function buildEvenPageFooterComplexTimeFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with complex even-page TIME chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:footerReference w:type="even" r:id="rIdEvenFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
      [
        'rIdEvenFooter',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`,
        'footer2.xml',
      ],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Continuing Header</w:t></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Continuing Footer</w:t></w:r></w:p></w:ftr>`,
    firstFooterXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>at </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TIME </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>17:18</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`,
    settingsXml: `<w:settings xmlns:w="${WORD_NAMESPACE}"><w:evenAndOddHeaders/></w:settings>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: first-page header SECTION vs default header text.
 */
export async function buildFirstPageHeaderSectionFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with first-page SECTION chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:headerReference w:type="first" r:id="rIdFirstHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:titlePg/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      [
        'rIdFirstHeader',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`,
        'header2.xml',
      ],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Acme Continuing Header</w:t></w:r></w:p></w:hdr>`,
    firstHeaderXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Section </w:t></w:r><w:fldSimple w:instr="SECTION"><w:r><w:t>2</w:t></w:r></w:fldSimple></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Continuing Footer</w:t></w:r></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: complex (fldChar) first-page header SECTION.
 */
export async function buildFirstPageHeaderComplexSectionFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with complex first-page SECTION chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:headerReference w:type="first" r:id="rIdFirstHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:titlePg/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      [
        'rIdFirstHeader',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`,
        'header2.xml',
      ],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Acme Continuing Header</w:t></w:r></w:p></w:hdr>`,
    firstHeaderXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Section </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SECTION </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Continuing Footer</w:t></w:r></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: first-page footer SECTIONPAGES vs default footer text.
 */
export async function buildFirstPageFooterSectionPagesFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with first-page SECTIONPAGES chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:footerReference w:type="first" r:id="rIdFirstFooter"/><w:titlePg/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
      [
        'rIdFirstFooter',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`,
        'footer2.xml',
      ],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Acme Continuing Header</w:t></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Continuing Footer</w:t></w:r></w:p></w:ftr>`,
    firstFooterXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>of </w:t></w:r><w:fldSimple w:instr="SECTIONPAGES"><w:r><w:t>4</w:t></w:r></w:fldSimple></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: complex (fldChar) first-page footer SECTIONPAGES.
 */
export async function buildFirstPageFooterComplexSectionPagesFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with complex first-page SECTIONPAGES chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:footerReference w:type="first" r:id="rIdFirstFooter"/><w:titlePg/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
      [
        'rIdFirstFooter',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`,
        'footer2.xml',
      ],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Acme Continuing Header</w:t></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Continuing Footer</w:t></w:r></w:p></w:ftr>`,
    firstFooterXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>of </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SECTIONPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>4</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: even-page header SECTION vs default header text.
 */
export async function buildEvenPageHeaderSectionFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with even-page SECTION chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:headerReference w:type="even" r:id="rIdEvenHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      [
        'rIdEvenHeader',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`,
        'header2.xml',
      ],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Continuing Header</w:t></w:r></w:p></w:hdr>`,
    firstHeaderXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Section </w:t></w:r><w:fldSimple w:instr="SECTION"><w:r><w:t>2</w:t></w:r></w:fldSimple></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Continuing Footer</w:t></w:r></w:p></w:ftr>`,
    settingsXml: `<w:settings xmlns:w="${WORD_NAMESPACE}"><w:evenAndOddHeaders/></w:settings>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: complex (fldChar) even-page header SECTION.
 */
export async function buildEvenPageHeaderComplexSectionFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with complex even-page SECTION chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:headerReference w:type="even" r:id="rIdEvenHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      [
        'rIdEvenHeader',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`,
        'header2.xml',
      ],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Continuing Header</w:t></w:r></w:p></w:hdr>`,
    firstHeaderXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Section </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SECTION </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Continuing Footer</w:t></w:r></w:p></w:ftr>`,
    settingsXml: `<w:settings xmlns:w="${WORD_NAMESPACE}"><w:evenAndOddHeaders/></w:settings>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: even-page footer SECTIONPAGES vs default footer text.
 */
export async function buildEvenPageFooterSectionPagesFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with even-page SECTIONPAGES chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:footerReference w:type="even" r:id="rIdEvenFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
      [
        'rIdEvenFooter',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`,
        'footer2.xml',
      ],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Continuing Header</w:t></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Continuing Footer</w:t></w:r></w:p></w:ftr>`,
    firstFooterXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>of </w:t></w:r><w:fldSimple w:instr="SECTIONPAGES"><w:r><w:t>4</w:t></w:r></w:fldSimple></w:p></w:ftr>`,
    settingsXml: `<w:settings xmlns:w="${WORD_NAMESPACE}"><w:evenAndOddHeaders/></w:settings>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: complex (fldChar) even-page footer SECTIONPAGES vs default footer.
 */
export async function buildEvenPageFooterComplexSectionPagesFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with complex even-page SECTIONPAGES chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:footerReference w:type="even" r:id="rIdEvenFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
      [
        'rIdEvenFooter',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`,
        'footer2.xml',
      ],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Continuing Header</w:t></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Continuing Footer</w:t></w:r></w:p></w:ftr>`,
    firstFooterXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>of </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SECTIONPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>4</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`,
    settingsXml: `<w:settings xmlns:w="${WORD_NAMESPACE}"><w:evenAndOddHeaders/></w:settings>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: distinct even-page vs default header/footer text.
 */
export async function buildEvenPageHeaderFooterFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with even-page chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:headerReference w:type="even" r:id="rIdEvenHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:footerReference w:type="even" r:id="rIdEvenFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      [
        'rIdEvenHeader',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`,
        'header2.xml',
      ],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
      [
        'rIdEvenFooter',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`,
        'footer2.xml',
      ],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Page Header</w:t></w:r></w:p></w:hdr>`,
    firstHeaderXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Even Page Header</w:t></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Odd Page Footer</w:t></w:r></w:p></w:ftr>`,
    firstFooterXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Even Page Footer</w:t></w:r></w:p></w:ftr>`,
    settingsXml: `<w:settings xmlns:w="${WORD_NAMESPACE}"><w:evenAndOddHeaders/></w:settings>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: header PAGE + footer NUMPAGES live field identity.
 */
export async function buildHeaderFooterFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with live chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Page </w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>2</w:t></w:r></w:fldSimple></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>of </w:t></w:r><w:fldSimple w:instr="NUMPAGES"><w:r><w:t>10</w:t></w:r></w:fldSimple></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: complex (fldChar) header PAGE + footer NUMPAGES identity.
 */
export async function buildHeaderFooterComplexFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with complex chrome fields</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>of </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>10</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: header DATE + footer TIME live field identity.
 */
export async function buildHeaderFooterDateTimeFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with date/time chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Printed </w:t></w:r><w:fldSimple w:instr="DATE \\@ &quot;yyyy-MM-dd&quot;"><w:r><w:t>2026-09-17</w:t></w:r></w:fldSimple></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>at </w:t></w:r><w:fldSimple w:instr="TIME"><w:r><w:t>18:54</w:t></w:r></w:fldSimple></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: complex header DATE + footer TIME live field identity.
 */
export async function buildHeaderFooterComplexDateTimeFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with complex date/time chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Printed </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE \\@ &quot;yyyy-MM-dd&quot; </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2026-09-18</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>at </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TIME </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>06:43</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: header SECTION + footer SECTIONPAGES live field identity.
 */
export async function buildHeaderFooterSectionFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with section chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Section </w:t></w:r><w:fldSimple w:instr="SECTION"><w:r><w:t>2</w:t></w:r></w:fldSimple></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>of </w:t></w:r><w:fldSimple w:instr="SECTIONPAGES"><w:r><w:t>4</w:t></w:r></w:fldSimple></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: complex header SECTION + footer SECTIONPAGES identity.
 */
export async function buildHeaderFooterComplexSectionFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><w:body><w:p><w:r><w:t>Report body with complex section chrome</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:body></w:document>`,
    documentRelationships: [
      ['rIdHeader', `${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, 'header1.xml'],
      ['rIdFooter', `${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, 'footer1.xml'],
    ],
    headerXml: `<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>Section </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SECTION </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>`,
    footerXml: `<w:ftr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t>of </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SECTIONPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>4</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Academic/report-shaped DOCX: footnote reference + note body identity.
 */
export async function buildFootnoteFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Body clause</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:sectPr/></w:body></w:document>`,
    documentRelationships: [
      [
        'rIdFootnotes',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/footnotes`,
        'footnotes.xml',
      ],
    ],
    footnotesXml: `<w:footnotes xmlns:w="${WORD_NAMESPACE}"><w:footnote w:type="separator" w:id="-1"/><w:footnote w:type="continuationSeparator" w:id="0"/><w:footnote w:id="1"><w:p><w:r><w:t>Cite the warranty clause</w:t></w:r></w:p></w:footnote></w:footnotes>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Academic/report-shaped DOCX: endnote reference + note body identity.
 */
export async function buildEndnoteFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Body with endnote</w:t></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p><w:sectPr/></w:body></w:document>`,
    documentRelationships: [
      [
        'rIdEndnotes',
        `${OFFICE_RELATIONSHIPS_NAMESPACE}/endnotes`,
        'endnotes.xml',
      ],
    ],
    endnotesXml: `<w:endnotes xmlns:w="${WORD_NAMESPACE}"><w:endnote w:type="separator" w:id="-1"/><w:endnote w:type="continuationSeparator" w:id="0"/><w:endnote w:id="1"><w:p><w:r><w:t>See appendix A</w:t></w:r></w:p></w:endnote></w:endnotes>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: inline text content control alias/tag/text identity.
 */
export async function buildTextContentControlFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:sdt><w:sdtPr><w:alias w:val="PartyName"/><w:tag w:val="party_name"/><w:id w:val="1001"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Acme Corp</w:t></w:r></w:sdtContent></w:sdt></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: rich-text content-control alias/tag/type/text identity.
 */
export async function buildRichTextContentControlFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:sdt><w:sdtPr><w:alias w:val="ClauseBody"/><w:tag w:val="clause_body"/><w:id w:val="1002"/><w:richText/></w:sdtPr><w:sdtContent><w:r><w:t>Indemnity survives termination.</w:t></w:r></w:sdtContent></w:sdt></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: body PAGE field kind/instruction identity.
 */
export async function buildPageFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Page </w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>3</w:t></w:r></w:fldSimple></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: bookmark + PAGEREF target/instruction identity.
 */
export async function buildPageRefFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:bookmarkStart w:id="9" w:name="Warranty"/><w:r><w:t>Warranty clause</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p><w:p><w:r><w:t>See page </w:t></w:r><w:fldSimple w:instr="PAGEREF Warranty \\h"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: body NUMPAGES field kind/instruction identity.
 */
export async function buildNumPagesFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Pages of </w:t></w:r><w:fldSimple w:instr="NUMPAGES"><w:r><w:t>12</w:t></w:r></w:fldSimple></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: body SECTION field kind/instruction identity.
 */
export async function buildSectionFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Section </w:t></w:r><w:fldSimple w:instr="SECTION"><w:r><w:t>2</w:t></w:r></w:fldSimple></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: body SECTIONPAGES field kind/instruction identity.
 */
export async function buildSectionPagesFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Section pages </w:t></w:r><w:fldSimple w:instr="SECTIONPAGES"><w:r><w:t>4</w:t></w:r></w:fldSimple></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: body NUMWORDS field kind/instruction identity.
 */
export async function buildNumWordsFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Words </w:t></w:r><w:fldSimple w:instr="NUMWORDS"><w:r><w:t>128</w:t></w:r></w:fldSimple></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: body NUMCHARS field kind/instruction identity.
 */
export async function buildNumCharsFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Chars </w:t></w:r><w:fldSimple w:instr="NUMCHARS"><w:r><w:t>640</w:t></w:r></w:fldSimple></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: body DATE field kind/instruction identity.
 */
export async function buildDateFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Date </w:t></w:r><w:fldSimple w:instr="DATE"><w:r><w:t>2026-09-17</w:t></w:r></w:fldSimple></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: body TIME field kind/instruction identity.
 */
export async function buildTimeFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Time </w:t></w:r><w:fldSimple w:instr="TIME"><w:r><w:t>16:05</w:t></w:r></w:fldSimple></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: body DATE with \@ format-switch instruction identity.
 */
export async function buildDateFormatFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Formatted </w:t></w:r><w:fldSimple w:instr="DATE \\@ &quot;yyyy-MM-dd&quot;"><w:r><w:t>2026-09-17</w:t></w:r></w:fldSimple></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: figure SEQ caption kind/id identity.
 */
export async function buildFigureCaptionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:bookmarkStart w:id="21" w:name="_RefFigureArchitecture"/><w:r><w:t>Figure </w:t></w:r><w:fldSimple w:instr=" SEQ Figure \\* ARABIC "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t>: Architecture</w:t></w:r><w:bookmarkEnd w:id="21"/></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: figure caption + caption REF target identity.
 */
export async function buildCaptionRefFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:bookmarkStart w:id="22" w:name="_RefFigureRuntime"/><w:r><w:t>Figure </w:t></w:r><w:fldSimple w:instr=" SEQ Figure \\* ARABIC "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t>: Runtime</w:t></w:r><w:bookmarkEnd w:id="22"/></w:p><w:p><w:r><w:t>See </w:t></w:r><w:fldSimple w:instr=" REF _RefFigureRuntime \\h "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Contract-shaped DOCX: table SEQ caption kind/id identity.
 */
export async function buildTableCaptionFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:bookmarkStart w:id="23" w:name="_RefTableSchedule"/><w:r><w:t>Table </w:t></w:r><w:fldSimple w:instr=" SEQ Table \\* ARABIC "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t>: Schedule</w:t></w:r><w:bookmarkEnd w:id="23"/></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Academic-shaped DOCX: XE index entry main/sub-entry identity.
 */
export async function buildIndexEntryFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Runtime</w:t></w:r><w:fldSimple w:instr="XE &quot;Architecture:Runtime&quot; \\b"><w:r><w:t/></w:r></w:fldSimple></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Academic-shaped DOCX: SDT-wrapped INDEX field columns identity.
 */
export async function buildIndexFieldFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Indexes"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> INDEX \\c "2" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Index1"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9360"/></w:tabs></w:pPr><w:r><w:t>Architecture</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:sdtContent></w:sdt><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Report-shaped DOCX: SDT-wrapped TOC options and cached entry identity.
 */
export async function buildTableOfContentsFixture(): Promise<Uint8Array> {
  const archive = new JSZip();
  writePackageSkeleton(archive, {
    documentXml: `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:sdt><w:sdtContent><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p><w:p><w:pPr><w:pStyle w:val="TOC1"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9360"/></w:tabs></w:pPr><w:r><w:t>Overview</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:sdtContent></w:sdt><w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>Overview</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
  });
  return archive.generateAsync({ type: 'uint8array' });
}

/**
 * Active-content fail-closed: safe custom parts may survive, VBA/signatures
 * must not be revived after a light edit + export.
 */
export async function buildActiveContentFailClosedFixture(): Promise<{
  bytes: Uint8Array;
  vendorPayload: Uint8Array;
}> {
  const vendorPayload = new TextEncoder().encode('vendor-safe-payload');
  const archive = new JSZip();
  archive.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="${CONTENT_TYPES_NAMESPACE}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/_xmlsignatures/sig1.xml" ContentType="application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"/></Types>`,
  );
  archive.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${RELATIONSHIPS_NAMESPACE}"><Relationship Id="rId1" Type="${OFFICE_RELATIONSHIPS_NAMESPACE}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  archive.file(
    'word/document.xml',
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>Original active-content boundary</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
  );
  archive.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${RELATIONSHIPS_NAMESPACE}"><Relationship Id="rIdVendor" Type="https://a3s.dev/relationships/vendor-data" Target="vendorData/payload.bin"/></Relationships>`,
  );
  archive.file('word/vendorData/payload.bin', vendorPayload);
  archive.file('word/vbaProject.bin', new TextEncoder().encode('macros'));
  archive.file('_xmlsignatures/sig1.xml', '<Signature/>');
  return {
    bytes: await archive.generateAsync({ type: 'uint8array' }),
    vendorPayload,
  };
}

function writePackageSkeleton(
  archive: JSZip,
  options: {
    documentXml: string;
    relationships?: Array<[string, string, string, string?]>;
    documentRelationships?: Array<[string, string, string]>;
    commentsXml?: string;
    headerXml?: string;
    footerXml?: string;
    firstHeaderXml?: string;
    firstFooterXml?: string;
    settingsXml?: string;
    footnotesXml?: string;
    endnotesXml?: string;
    numberingXml?: string;
    stylesXml?: string;
  },
): void {
  const overrides = [
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
  ];
  if (options.commentsXml) {
    overrides.push(
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>',
    );
  }
  if (options.headerXml) {
    overrides.push(
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
    );
  }
  if (options.firstHeaderXml) {
    overrides.push(
      '<Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
    );
  }
  if (options.footerXml) {
    overrides.push(
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>',
    );
  }
  if (options.firstFooterXml) {
    overrides.push(
      '<Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>',
    );
  }
  if (options.settingsXml) {
    overrides.push(
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>',
    );
  }
  if (options.footnotesXml) {
    overrides.push(
      '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>',
    );
  }
  if (options.endnotesXml) {
    overrides.push(
      '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>',
    );
  }
  if (options.numberingXml) {
    overrides.push(
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>',
    );
  }
  if (options.stylesXml) {
    overrides.push(
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    );
  }
  archive.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="${CONTENT_TYPES_NAMESPACE}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides.join('')}</Types>`,
  );
  archive.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${RELATIONSHIPS_NAMESPACE}"><Relationship Id="rId1" Type="${OFFICE_RELATIONSHIPS_NAMESPACE}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  archive.file('word/document.xml', options.documentXml);
  if (options.commentsXml) {
    archive.file('word/comments.xml', options.commentsXml);
  }
  if (options.headerXml) {
    archive.file('word/header1.xml', options.headerXml);
  }
  if (options.firstHeaderXml) {
    archive.file('word/header2.xml', options.firstHeaderXml);
  }
  if (options.footerXml) {
    archive.file('word/footer1.xml', options.footerXml);
  }
  if (options.firstFooterXml) {
    archive.file('word/footer2.xml', options.firstFooterXml);
  }
  if (options.settingsXml) {
    archive.file('word/settings.xml', options.settingsXml);
  }
  if (options.footnotesXml) {
    archive.file('word/footnotes.xml', options.footnotesXml);
  }
  if (options.endnotesXml) {
    archive.file('word/endnotes.xml', options.endnotesXml);
  }
  if (options.numberingXml) {
    archive.file('word/numbering.xml', options.numberingXml);
  }
  if (options.stylesXml) {
    archive.file('word/styles.xml', options.stylesXml);
  }
  const relationships = [
    ...(options.documentRelationships ?? []),
    ...(options.relationships ?? []),
  ];
  if (options.settingsXml) {
    relationships.unshift([
      'rIdSettings',
      `${OFFICE_RELATIONSHIPS_NAMESPACE}/settings`,
      'settings.xml',
    ]);
  }
  if (options.numberingXml) {
    relationships.unshift([
      'rIdNumbering',
      `${OFFICE_RELATIONSHIPS_NAMESPACE}/numbering`,
      'numbering.xml',
    ]);
  }
  if (options.stylesXml) {
    relationships.unshift([
      'rIdStyles',
      `${OFFICE_RELATIONSHIPS_NAMESPACE}/styles`,
      'styles.xml',
    ]);
  }
  if (relationships.length > 0) {
    archive.file(
      'word/_rels/document.xml.rels',
      `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">${relationships
        .map((entry) => {
          const [id, type, target, mode] = entry;
          return mode
            ? `<Relationship Id="${id}" Type="${type}" Target="${target}" TargetMode="${mode}"/>`
            : `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
        })
        .join('')}</Relationships>`,
    );
  }
}
