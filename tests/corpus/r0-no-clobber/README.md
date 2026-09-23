# R0 no-clobber corpus

Permanent R0 gate for representative DOCX round trips.

## Run

```bash
bun run test:corpus:r0
```

## Contract

Each fixture must prove:

1. Import → light/no-op edit path → export → reopen does not silently drop
   critical identities (bookmarks, revision authors/text, review comments,
   internal and external links, headers/footers (including first-page and even-page variants), footnotes, endnotes, text and
   rich-text content controls, PAGE/PAGEREF/NUMPAGES/SECTION/SECTIONPAGES/NUMWORDS/
   NUMCHARS/DATE/TIME fields (including DATE `\@` format switches), header/footer
   PAGE and NUMPAGES as `fldSimple` or complex `fldChar` live fields, header DATE
   (`\@`) and footer TIME (fldSimple and complex), header SECTION and footer
   SECTIONPAGES live fields (`fldSimple` and complex `fldChar`), figure
   SEQ captions and caption REF targets, table SEQ captions, XE/INDEX index
   identities, TOC options/entries, table cells).
2. Every intentional normalization appears in `compatibility.issues` with a
   stable code.
3. Active content (VBA, package signatures) stays fail-closed on export.

Current fixtures cover report bookmarks/external links, internal
bookmark-anchor hyperlinks, review track changes, reviewable bounded w:rPrChange character-formatting revisions, reviewable jc-only w:pPrChange paragraph-formatting revisions, reviewable whole-paragraph mark insertion revisions, reviewable paragraph-break split revisions, reviewable paired text-move revisions, reviewable companion move-range bookmark text-move revisions, reviewable jc-only w:tblPrChange table-formatting revisions, reviewable tblW-only w:tblPrChange preferred-width table-formatting revisions, reviewable tblInd-only w:tblPrChange table-indent revisions, reviewable tblCellMar-only w:tblPrChange table cell-margin revisions, reviewable tblLayout-only w:tblPrChange table-layout revisions, reviewable bidiVisual-only w:tblPrChange RTL visual revisions, reviewable solid-shd-only w:tblPrChange table-fill revisions, reviewable tblOverlap-only w:tblPrChange overlap revisions, reviewable tblLook-only w:tblPrChange table-look revisions, reviewable tblStyle-only w:tblPrChange table-style revisions, reviewable tblCellSpacing-only w:tblPrChange table cell-spacing revisions, reviewable tblBorders-only w:tblPrChange table-border revisions, reviewable tblStyleColBandSize-only w:tblPrChange column-band revisions, reviewable tblStyleRowBandSize-only w:tblPrChange row-band revisions, reviewable tblCaption-only w:tblPrChange table-caption revisions, reviewable tblDescription-only w:tblPrChange table-description revisions, reviewable tblpPr-only w:tblPrChange table-float revisions, reviewable orientation-only w:sectPrChange section-formatting revisions, reviewable pgMar-only w:sectPrChange section-margin revisions, reviewable titlePg-only w:sectPrChange title-page revisions, reviewable equal-width cols-only w:sectPrChange column revisions, reviewable rtlGutter-only w:sectPrChange RTL gutter revisions, reviewable paperSrc-only w:sectPrChange paper-source revisions, reviewable vAlign-only w:sectPrChange section vertical-align revisions, reviewable docGrid-only w:sectPrChange document-grid revisions, reviewable bidi-only w:sectPrChange section bidi revisions, reviewable type-only w:sectPrChange section-break revisions, reviewable complete pgSz geometry-only w:sectPrChange page-geometry revisions, reviewable unequal-width cols-only w:sectPrChange unequal-column revisions, reviewable lnNumType-only w:sectPrChange line-numbering revisions, reviewable pgNumType-only w:sectPrChange page-number format revisions, reviewable formProt-only w:sectPrChange form-protection revisions, reviewable noEndnote-only w:sectPrChange endnote-suppression revisions, reviewable textDirection-only w:sectPrChange section text-direction revisions, reviewable footnotePr-only w:sectPrChange footnote-properties revisions, reviewable endnotePr-only w:sectPrChange endnote-properties revisions, reviewable pgBorders-only w:sectPrChange page-border revisions, reviewable cantSplit-only w:trPrChange row-formatting revisions, reviewable tblHeader-only w:trPrChange header-row revisions, reviewable trHeight-only w:trPrChange row-height revisions, reviewable hidden-only w:trPrChange row-visibility revisions, reviewable jc-only w:trPrChange row-alignment revisions, reviewable gridBefore-only w:trPrChange leading-grid revisions, reviewable gridAfter-only w:trPrChange trailing-grid revisions, reviewable wBefore-only w:trPrChange leading-width revisions, reviewable wAfter-only w:trPrChange trailing-width revisions, reviewable cnfStyle-only w:trPrChange conditional-format revisions, reviewable divId-only w:trPrChange HTML-div revisions, reviewable tblCellSpacing-only w:trPrChange cell-spacing revisions, reviewable vAlign-only w:tcPrChange cell-formatting revisions, reviewable solid-shd-only w:tcPrChange cell-fill revisions, reviewable tcW-only w:tcPrChange preferred-width cell revisions, reviewable tcMar-only w:tcPrChange cell-margin revisions, reviewable noWrap-only w:tcPrChange cell-wrap revisions, reviewable textDirection-only w:tcPrChange text-direction revisions, reviewable tcFitText-only w:tcPrChange fit-text revisions, reviewable hideMark-only w:tcPrChange hide-mark revisions, reviewable cnfStyle-only w:tcPrChange conditional-format cell revisions, reviewable hMerge-only w:tcPrChange horizontal-merge revisions, reviewable vMerge-only w:tcPrChange vertical-merge revisions, reviewable gridSpan-only w:tcPrChange column-span revisions, reviewable tcBorders-only w:tcPrChange cell-border revisions, reviewable current-level bullet w:numberingChange (nfc 23) numbering revisions, reviewable current-level decimal w:numberingChange numbering revisions, reviewable multi-level w:numberingChange (sibling original levels) numbering revisions, review comments, default
headers/footers, first-page vs default header/footer text, even-page vs default header/footer text, first-page header PAGE live fields (fldSimple and complex fldChar), even-page header PAGE live fields (fldSimple and complex fldChar), first-page footer NUMPAGES live fields (fldSimple and complex fldChar), even-page footer NUMPAGES live fields (fldSimple and complex fldChar), first-page header DATE (`\@` format) live fields (fldSimple and complex fldChar), first-page footer TIME live fields (fldSimple and complex fldChar), even-page header DATE (`\@` format) live fields (fldSimple and complex fldChar), even-page footer TIME live fields (fldSimple and complex fldChar), first-page header SECTION live fields (fldSimple and complex fldChar), first-page footer SECTIONPAGES live fields (fldSimple and complex fldChar), even-page header SECTION live fields (fldSimple and complex fldChar), even-page footer SECTIONPAGES live fields (fldSimple and complex fldChar), footnotes, endnotes, text and rich-text content controls,
body PAGE, bookmark-backed PAGEREF, NUMPAGES, SECTION, SECTIONPAGES, NUMWORDS,
NUMCHARS, DATE (bare and `\@` format), and TIME fields, header PAGE and footer
NUMPAGES live fields (both `fldSimple` and complex `fldChar` forms), header DATE (`\@` format) and footer TIME live fields (fldSimple and complex fldChar), header SECTION and footer SECTIONPAGES live fields (fldSimple and complex fldChar), figure SEQ captions and caption REF targets, table SEQ
captions, XE index entries and INDEX field columns, TOC level/hyperlink options
and cached entry titles, contract tables, duplicate-bookmark identity
diagnostics, and active-content fail-closed export.

Grow fixtures here when a real Traditional Office sample still fails the
contract. Do not grow decorative PDF art-border coverage in place of this gate.
