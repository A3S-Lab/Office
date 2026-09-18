import JSZip from 'jszip';

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
