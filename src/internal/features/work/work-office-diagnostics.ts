import type { WorkBook, WorkSheet } from 'xlsx';
import { diagnoseDocxBookmarksAndLinks } from './work-docx-bookmark-diagnostics';
import { diagnoseDocxCaptions } from './work-docx-caption-diagnostics';
import {
  isSupportedDocxMoveChange,
  supportedDocxMovePairCount,
} from './work-docx-change-import';
import { diagnoseDocxCitations } from './work-docx-citation-diagnostics';
import { inspectDocxConnectors } from './work-docx-connector-diagnostics';
import { inspectDocxContentControls } from './work-docx-content-control-import';
import { diagnoseDocxEmphasisMarks } from './work-docx-emphasis-diagnostics';
import { diagnoseDocxEquations } from './work-docx-equation-diagnostics';
import { diagnoseDocxHiddenText } from './work-docx-hidden-text-diagnostics';
import { readDocxImageTransform } from './work-docx-image-transform';
import { diagnoseDocxKerning } from './work-docx-kerning-diagnostics';
import { diagnoseDocxLegacyTextEffects } from './work-docx-legacy-text-effects-diagnostics';
import { diagnoseDocxNotes } from './work-docx-note-diagnostics';
import { isSupportedDocxNumberingChange } from './work-docx-numbering-change-import';
import { diagnoseDocxOpenTypeTypography } from './work-docx-opentype-diagnostics';
import { diagnoseDocxPageBorders } from './work-docx-page-borders-diagnostics';
import { diagnoseDocxPageChrome } from './work-docx-page-chrome-diagnostics';
import { diagnoseDocxPageMargins } from './work-docx-page-margins-diagnostics';
import { diagnoseDocxPageSize } from './work-docx-page-size-diagnostics';
import { diagnoseDocxParagraphBorders } from './work-docx-paragraph-borders-diagnostics';
import { parseDocxParagraphDefaultCollapsed } from './work-docx-paragraph-default-collapsed';
import { isSupportedDocxParagraphFormattingChange } from './work-docx-paragraph-format-change-import';
import { isSupportedDocxParagraphMarkChange } from './work-docx-paragraph-mark-change-import';
import { diagnoseDocxParagraphShading } from './work-docx-paragraph-shading-diagnostics';
import { diagnoseDocxProofing } from './work-docx-proofing-diagnostics';
import { diagnoseDocxRunBorders } from './work-docx-run-border-diagnostics';
import { diagnoseDocxRunFonts } from './work-docx-run-fonts-diagnostics';
import { isSupportedDocxRunFormattingChange } from './work-docx-run-formatting-import';
import { diagnoseDocxRunShading } from './work-docx-run-shading-diagnostics';
import { inspectDocxTextBoxes } from './work-docx-text-box-import';
import {
  attribute,
  contentTypeForPart,
  descendants,
  directChild,
  directChildren,
  firstDescendant,
  OoxmlPackage,
  parseXml,
  xmlContainsAnyElement,
} from './work-ooxml-package';
import {
  emptySpreadsheetWorksheetCompatibilitySummary,
  type SpreadsheetWorksheetCompatibilitySummary,
  updateSpreadsheetWorksheetCompatibilitySummary,
} from './work-spreadsheet-compatibility-summary';
import {
  parseSpreadsheetPrintTitles,
  stripSpreadsheetSheetQualifier,
} from './work-spreadsheet-ranges';
import type {
  WorkCompatibilityIssue,
  WorkCompatibilityReport,
} from './work-types';
import { diagnoseXlsxCharts } from './work-xlsx-chart-diagnostics';
import { diagnoseXlsxConditionalFormatting } from './work-xlsx-conditional-format-diagnostics';
import { diagnoseXlsxFormulas } from './work-xlsx-formula-diagnostics';
import type { XlsxFormulaFeatures } from './work-xlsx-formulas';
import {
  isSupportedXlsxWorksheetImageContentType,
  MAX_XLSX_WORKSHEET_IMAGE_BYTES,
} from './work-xlsx-images';
import { diagnoseXlsxPageSetup } from './work-xlsx-page-setup-diagnostics';
import { diagnoseXlsxPivots } from './work-xlsx-pivot-diagnostics';
import type { XlsxPivotReadResult } from './work-xlsx-pivots';
import { diagnoseXlsxProtection } from './work-xlsx-protection';
import { xlsxWorksheetCellEntries } from './work-xlsx-worksheet';
import type { XlsxWorksheetXmlScan } from './work-xlsx-worksheet-scan';

interface ConversionMessage {
  type: string;
  message: string;
}

export interface SpreadsheetCompatibilityImportMetadata {
  formulaFeatures?: XlsxFormulaFeatures | null;
  pivotFeatures?: XlsxPivotReadResult | null;
  worksheetScans?: Readonly<Record<string, XlsxWorksheetXmlScan>>;
}

const XLSX_DIAGNOSTIC_WORKSHEET_ELEMENTS = [
  'conditionalFormatting',
  'dataValidation',
  'sheetProtection',
  'protectedRange',
  'pageSetup',
  'pageMargins',
  'printOptions',
  'headerFooter',
  'pageSetUpPr',
  'rowBreaks',
  'colBreaks',
] as const;

export async function analyzeDocxCompatibility(
  file: File,
  messages: ConversionMessage[],
  sourcePackage?: OoxmlPackage | null,
): Promise<WorkCompatibilityReport> {
  const issues: WorkCompatibilityIssue[] = [
    issue(
      'docx.page-layout',
      'Page layout',
      'Per-section paper size, orientation, exact native page margins, one-to-six equal or custom-width columns, section breaks, and explicit page breaks are preserved; exact pagination and line wrapping may normalize.',
    ),
    issue(
      'docx.package-state',
      'OOXML package state',
      'Source-backed export retains safe source-only parts byte-for-byte and reconnects their content types and relationships. Passive relationship-free extensions in settings XML, uniquely matched style or numbering identities, stable picture drawings, unchanged native-identity paragraphs, stable table, row, or cell scopes, and uniquely matched footnote, endnote, comment, or comment-thread records are preserved against the final package graph. Text-stable direct, supported-hyperlink-wrapped, and eligible static-content-control note and comment runs retain paragraph, table, wrapper, run, and run-property metadata; native note tables and DrawingML pictures remain editable, note image relationships and media targets are repaired and validated, bounded structured OMML equations remain atomically updateable in supported Word stories, uniquely matched table-bearing rich-text controls are reconstructed, safe web, mail, and internal note/comment links are rebound to collision-free final relationship IDs, and content-control ID collisions are rewritten. Native note, comment, and note-picture identities survive reorderings, comment durable IDs are rebound to regenerated paragraph IDs, and eligible source font-table metadata is retained. Generated Word semantics remain authoritative; source-only, changed-text or table structure, malformed, duplicate or cross-kind, active, unsupported wrappers, mixed-semantic, unsafe relationship-bound, legacy VML, non-picture drawings, or unsupported rich content may normalize.',
      'info',
    ),
  ];
  for (const message of messages) {
    issues.push(
      issue(
        `docx.converter.${message.type}`,
        'Document conversion',
        message.message,
        message.type === 'error' ? 'error' : 'warning',
      ),
    );
  }

  try {
    const archive =
      sourcePackage === undefined
        ? await OoxmlPackage.load(await file.arrayBuffer())
        : sourcePackage;
    if (!archive) throw new Error('The DOCX package could not be loaded.');
    const packagePaths = archive.paths('');
    const normalizedPackagePaths = packagePaths.map((path) =>
      path.toLowerCase(),
    );
    const document = archive.has('word/document.xml')
      ? await archive.xml('word/document.xml')
      : null;
    if (
      normalizedPackagePaths.some((path) => path.startsWith('_xmlsignatures/'))
    ) {
      issues.push(
        issue(
          'docx.signatures.removed',
          'Digital signatures',
          'Package signatures become invalid after an edit and are deliberately omitted from the exported DOCX.',
        ),
      );
    }
    if (
      normalizedPackagePaths.some(
        (path) =>
          path === 'word/vbaproject.bin' ||
          path === 'word/vbadata.xml' ||
          path.startsWith('word/activex/') ||
          path.startsWith('customui/'),
      )
    ) {
      issues.push(
        issue(
          'docx.active-content.removed',
          'Active content',
          'VBA, ActiveX, and custom-ribbon parts are never executed and are omitted from the macro-free DOCX export.',
        ),
      );
    }
    if (archive.has('word/fontTable.xml')) {
      const fontTable = await archive.xml('word/fontTable.xml');
      const embeddedFontReferences = [
        'embedRegular',
        'embedBold',
        'embedItalic',
        'embedBoldItalic',
      ].reduce(
        (count, localName) => count + descendants(fontTable, localName).length,
        0,
      );
      if (embeddedFontReferences) {
        issues.push(
          issue(
            'docx.embedded-fonts',
            'Embedded fonts',
            `${embeddedFontReferences} embedded font reference(s) were found. Eligible internal obfuscated-font payloads, metadata, relationships, and remapped IDs survive source-backed DOCX export, but the browser editor, preview, and PDF renderer use registered A3S fonts or substitution and may wrap text differently.`,
          ),
        );
      }
    }
    if (archive.has('word/comments.xml')) {
      const comments = await archive.xml('word/comments.xml');
      const commentDefinitions = descendants(comments, 'comment');
      if (commentDefinitions.length) {
        issues.push(
          issue(
            'docx.comments',
            'Comments',
            `${commentDefinitions.length} comment or reply record(s), stable native IDs, text anchors, authors, dates, thread relationships, resolved state, eligible durable IDs, unchanged direct-run formatting, safely rebound web, mail, or internal hyperlinks, and eligible static text content controls are preserved. Comment text remains editable.`,
            'info',
          ),
        );
        if (
          ['tbl', 'drawing', 'pict', 'sdt', 'altChunk'].some(
            (name) => descendants(comments, name).length,
          ) ||
          descendants(comments, 'rPr').length ||
          descendants(comments, 'hyperlink').length
        ) {
          issues.push(
            issue(
              'docx.comments.formatting',
              'Comment formatting',
              'Relationship-free formatting on uniquely matched unchanged direct text runs and eligible reconstructed control runs is retained. Supported HTTP(S), mailto, and internal hyperlink wrappers keep stable formatting, safe tooltips, and collision-free relationship IDs. Eligible static rich-text or plain-text controls retain inline or contiguous block wrappers, aliases, tags, locks, rewritten collision IDs, appearance, color, end formatting, and passive metadata. Editing comment text or using active bindings, placeholder state, form or nested controls, tables, images, math, mixed hyperlinks, unsupported wrappers, malformed relationships, or relationship-bound properties normalizes the affected content to safe plain text.',
            ),
          );
        }
        if (
          archive.has('word/commentsExtensible.xml') ||
          archive.has('word/people.xml')
        ) {
          issues.push(
            issue(
              'docx.comments.modern-sidecars',
              'Modern comment metadata',
              'Durable comment IDs are safely rebound, but follow-up flags, reactions, and people metadata in modern comment sidecars are not editable and are omitted rather than reattached to a stale comment identity.',
            ),
          );
        }
      }
    }
    issues.push(...(await diagnoseDocxNotes(archive, document)));
    issues.push(...(await diagnoseDocxEquations(archive, document)));
    if (document) {
      issues.push(...(await diagnoseDocxKerning(archive, document)));
      issues.push(...(await diagnoseDocxOpenTypeTypography(archive, document)));
      issues.push(...(await diagnoseDocxEmphasisMarks(archive, document)));
      issues.push(...(await diagnoseDocxHiddenText(archive, document)));
      issues.push(...(await diagnoseDocxLegacyTextEffects(archive, document)));
      issues.push(...(await diagnoseDocxRunBorders(archive, document)));
      issues.push(...(await diagnoseDocxRunShading(archive, document)));
      issues.push(...(await diagnoseDocxProofing(archive, document)));
      issues.push(...(await diagnoseDocxRunFonts(archive, document)));
    }
    issues.push(...(await diagnoseDocxPageChrome(archive)));
    if (archive.paths('word/embeddings/').length) {
      issues.push(
        issue(
          'docx.embedded',
          'Embedded objects',
          'Embedded Office and OLE payload parts plus safe relationships are retained in source-backed exports, but Work does not render or edit their object anchors; placement inside regenerated document XML may normalize.',
        ),
      );
    }
    if (document) {
      issues.push(...diagnoseDocxPageSize(document));
      issues.push(...(await diagnoseDocxPageMargins(archive, document)));
      issues.push(...(await diagnoseDocxPageBorders(archive, document)));
      issues.push(...(await diagnoseDocxParagraphBorders(archive, document)));
      issues.push(...(await diagnoseDocxParagraphShading(archive, document)));
      const defaultCollapsedStates = descendants(document, 'pPr')
        .map(parseDocxParagraphDefaultCollapsed)
        .filter((state) => state.status !== 'absent');
      if (defaultCollapsedStates.length) {
        const invalid = defaultCollapsedStates.filter(
          (state) => state.status === 'invalid',
        ).length;
        issues.push(
          issue(
            'docx.headings.default-collapsed',
            'Default-collapsed headings',
            invalid
              ? `Office 2013 default-collapsed heading state is preserved for exact core Word on/off values, but ${invalid} malformed or duplicated paragraph property value(s) are ignored instead of inheriting a stale state.`
              : "Office 2013 default-collapsed heading state is preserved as native paragraph metadata. Browser content remains expanded and editable because the property controls Word's initial document view rather than content visibility.",
            invalid ? 'warning' : 'info',
          ),
        );
      }
      if (descendants(document, 'AlternateContent').length) {
        issues.push(
          issue(
            'docx.markup-compatibility',
            'Markup compatibility',
            '`mc:AlternateContent` branches outside supported stable paragraph, picture-drawing, table, row, or cell scopes can normalize. Safe settings-level and uniquely matched passive branches are preserved separately, but source-backed package preservation does not imply general body-content fallback preservation.',
          ),
        );
      }
      issues.push(...(await diagnoseDocxBookmarksAndLinks(archive, document)));
      issues.push(...(await diagnoseDocxCitations(archive, document)));
      const captionDiagnostics = diagnoseDocxCaptions(document);
      issues.push(...captionDiagnostics.issues);
      const textBoxInspection = inspectDocxTextBoxes(document);
      const connectorInspection = inspectDocxConnectors(document);
      if (textBoxInspection.supported) {
        issues.push(
          issue(
            'docx.text-boxes',
            'Text boxes and shapes',
            `${textBoxInspection.supported} isolated WPS text box or common shape(s) remain editable with bounded DrawingML geometry (rectangle, rounded rectangle, ellipse, diamond, or triangle), inline or floating layout, safe page-relative offsets, fill, outline, padding, and vertical alignment. Mixed paragraphs and connectors remain on the normal DOCX compatibility path.`,
            'info',
          ),
        );
      }
      if (textBoxInspection.unsupported) {
        issues.push(
          issue(
            'docx.text-boxes.unsupported',
            'Text boxes and shapes',
            `${textBoxInspection.unsupported} WPS text box or shape(s) mix a drawing with other paragraph content, use unsupported geometry, or have malformed shape content. Only isolated text-bearing paragraphs in the bounded preset set are converted to editable shapes; connectors and the rest may normalize during browser conversion.`,
          ),
        );
      }
      if (connectorInspection.detected) {
        issues.push(
          issue(
            'docx.connectors',
            'Connectors',
            `${connectorInspection.detected} WPS VML connector(s) were detected. Their endpoints, routing, arrowheads, and floating anchor are not yet represented by the editable Writer model, so they remain on the compatibility path instead of being misclassified as text boxes; editing or regenerating the document may normalize them.`,
          ),
        );
      }
      const contentControlInspection = inspectDocxContentControls(document);
      if (contentControlInspection.supported) {
        issues.push(
          issue(
            'docx.content-controls',
            'Content controls',
            `${contentControlInspection.supported} inline text or rich-text content control(s) remain editable with aliases, tags, bounded locking, multiline text, appearance, color, and native DOCX w:sdt round-tripping. Data bindings, placeholders, repeating sections, form controls, and block or relationship-bound controls remain outside the safe subset.`,
            'info',
          ),
        );
      }
      if (contentControlInspection.unsupported) {
        issues.push(
          issue(
            'docx.content-controls.unsupported',
            'Content controls',
            `${contentControlInspection.unsupported} body content control(s) use active bindings, placeholders, form semantics, unsupported structure, or malformed properties. They are flattened to safe editable text during browser conversion instead of being revived with incomplete behavior.`,
          ),
        );
      }
      if (descendants(document, 'tbl').length) {
        issues.push(
          issue(
            'docx.tables',
            'Tables',
            'Table rows, non-splitting rows, repeated headers, cell shading, vertical alignment, per-edge borders, and paragraph-boundary row continuation remain editable. Native row identities survive body and page-chrome HTML; uniquely matched passive extensions on stable table, row, and cell scopes and their property nodes are retained, while duplicate, cross-kind, relationship-bound, or semantic branches fail closed. Layout algorithm, preferred auto, percentage, or pixel width, alignment, indent, percentage or pixel column widths, table-level cell margins, and cell-level margin overrides are preserved independently across edit, preview, and DOCX export. Percentage column preferences use pixel grid widths as responsive browser fallbacks and remain coherent through merged cells. Nested tables retain independent edit and DOCX geometry, and tall outer rows can paginate at nested-row boundaries. Row-spanning cells continue through every covered physical row with in-cell page breaks and contiguous content ranges; combined row and column spans retain native DOCX merge semantics. Common inherited Word table styles apply whole-table, banded row or column, first or last row or column, corner-cell, border, fill, text-emphasis, paragraph alignment, direction, indents, spacing, line, pagination, and tab-stop rules before direct table, cell, paragraph, and run formatting. Theme tint and shade values resolve to stable RGB for edit and preview, while untouched run colors, run shading, cell fills, and per-edge borders retain their semantic theme references on export; explicit color edits use direct RGB. Less-common conditional paragraph properties may be normalized.',
          ),
        );
      }
      const drawingCount = descendants(document, 'drawing').length;
      const textBoxDrawingCount =
        textBoxInspection.supported + textBoxInspection.unsupported;
      const pictCount = descendants(document, 'pict').length;
      if (
        drawingCount > textBoxDrawingCount ||
        pictCount > connectorInspection.connectorPictContainers
      ) {
        const unsupportedImageTransforms =
          docxUnsupportedImageTransformCount(document);
        issues.push(
          issue(
            'docx.images',
            'Images',
            'Inline images and supported square, tight, through, top-and-bottom, or no-wrap free-floating images remain embedded with editable size, alternative text, alignment, wrap side and distance where applicable, signed horizontal or vertical offsets relative to the column, paragraph, margin, or page, four-edge percentage cropping, 90-degree rotation, horizontal or vertical reflection, and drawing-layer options. Aligned and precise DOCX anchors retain their distinct position semantics; inline and floating crop geometry round-trips as DrawingML source rectangles; tight or through anchors retain their ordered wrap polygons; no-wrap anchors remain outside the body text flow and use behind-text placement to choose their side of the text; and supported floating images preserve relative Z-order, overlap and cell-layout policy, and anchor locking. Inline and floating pictures retain unique drawing-property and anchor IDs together with their edit IDs. Copies receive independent identities while move, delete, undo, and relationship regeneration retain the original object identity. Arbitrary-angle or malformed picture transforms and unsupported drawing types may be normalized.',
          ),
        );
        if (unsupportedImageTransforms) {
          issues.push(
            issue(
              'docx.images.transform',
              'Image transforms',
              `${unsupportedImageTransforms} picture transform(s) use an arbitrary angle or malformed rotation/reflection value. Work imports the editable 90-degree/reflection subset and normalizes the unsupported transform instead of exposing an inexact control.`,
            ),
          );
        }
      }
      const textRevisions = [
        ...descendants(document, 'ins'),
        ...descendants(document, 'del'),
      ];
      const moveFromRevisions = descendants(document, 'moveFrom');
      const moveToRevisions = descendants(document, 'moveTo');
      const moveRevisionCount =
        moveFromRevisions.length + moveToRevisions.length;
      const supportedMovePairCount = supportedDocxMovePairCount(document);
      const supportedMoveRevisionCount = [
        ...moveFromRevisions,
        ...moveToRevisions,
      ].filter(isSupportedDocxMoveChange).length;
      const runFormattingRevisions = descendants(document, 'rPrChange');
      const paragraphFormattingRevisions = descendants(document, 'pPrChange');
      const paragraphMarkRevisions = [
        ...descendants(document, 'ins'),
        ...descendants(document, 'del'),
      ].filter(
        (revision) =>
          revision.parentElement?.localName === 'rPr' &&
          revision.parentElement.parentElement?.localName === 'pPr' &&
          revision.parentElement.parentElement.parentElement?.localName === 'p',
      );
      const numberingRevisions = descendants(document, 'numberingChange');
      const supportedRunFormattingRevisionCount = runFormattingRevisions.filter(
        isSupportedDocxRunFormattingChange,
      ).length;
      const supportedParagraphFormattingRevisionCount =
        paragraphFormattingRevisions.filter(
          isSupportedDocxParagraphFormattingChange,
        ).length;
      const supportedParagraphMarkRevisions = paragraphMarkRevisions.filter(
        isSupportedDocxParagraphMarkChange,
      );
      const supportedParagraphMarkRevisionCount =
        supportedParagraphMarkRevisions.length;
      const supportedNumberingRevisionCount = numberingRevisions.filter(
        isSupportedDocxNumberingChange,
      ).length;
      if (
        textRevisions.some(
          (revision) =>
            descendants(revision, 't').length ||
            descendants(revision, 'delText').length,
        )
      ) {
        issues.push(
          issue(
            'docx.revisions',
            'Tracked changes',
            'Body-text insertions and deletions preserve their author and date, remain reviewable in Work, and round-trip as native DOCX revisions.',
            'info',
          ),
        );
      }
      if (supportedRunFormattingRevisionCount) {
        issues.push(
          issue(
            'docx.revisions.formatting',
            'Character-formatting revisions',
            `${supportedRunFormattingRevisionCount} bounded character-formatting revision(s) preserve author, date, current formatting, and prior bold, italic, underline, strike, text case, hidden text, outline, shadow, emboss, imprint, subscript, superscript, font, size, color, highlight, character borders, character shading, character scale, spacing, kerning threshold, emphasis mark, baseline position, grid state, proofing languages, and explicit proofing state. They remain reviewable in Work and round-trip as native w:rPrChange records.`,
            'info',
          ),
        );
      }
      if (supportedParagraphFormattingRevisionCount) {
        issues.push(
          issue(
            'docx.revisions.paragraph-formatting',
            'Paragraph-formatting revisions',
            `${supportedParagraphFormattingRevisionCount} bounded paragraph-formatting revision(s) preserve author, date, current formatting, and prior alignment, direction, indentation, spacing, pagination, outline, tab-stop, border, shading, and collapsed state. They remain reviewable in Work and round-trip as native w:pPrChange records.`,
            'info',
          ),
        );
      }
      if (supportedNumberingRevisionCount) {
        issues.push(
          issue(
            'docx.revisions.numbering',
            'Numbering revisions',
            `${supportedNumberingRevisionCount} bounded ordered-list numbering revision(s) preserve author, date, prior start, and common decimal, letter, or Roman formats. Contiguous list-item records remain reviewable as one Work change and round-trip as native w:numberingChange records.`,
            'info',
          ),
        );
      }
      if (supportedParagraphMarkRevisionCount) {
        issues.push(
          issue(
            'docx.revisions.paragraph-mark',
            'Paragraph-mark revisions',
            `${supportedParagraphMarkRevisionCount} bounded paragraph-mark insertion/deletion revision(s) preserve author, date, and whole-paragraph accept/reject semantics through Work and native DOCX w:pPr/w:rPr/w:ins or w:del round trips.`,
            'info',
          ),
        );
      }
      if (supportedMovePairCount) {
        issues.push(
          issue(
            'docx.revisions.move',
            'Move revisions',
            `${supportedMovePairCount} bounded text move revision(s) preserve author, date, source and destination text, remain reviewable as one atomic Work change, and round-trip as native w:moveFrom and w:moveTo records. Rich content, range markers, and relationship-bound moves remain on the compatibility path.`,
            'info',
          ),
        );
      }
      if (
        textRevisions.some(
          (revision) =>
            !supportedParagraphMarkRevisions.includes(revision) &&
            !descendants(revision, 't').length &&
            !descendants(revision, 'delText').length,
        ) ||
        supportedRunFormattingRevisionCount !== runFormattingRevisions.length ||
        supportedParagraphFormattingRevisionCount !==
          paragraphFormattingRevisions.length ||
        supportedParagraphMarkRevisionCount !== paragraphMarkRevisions.length ||
        supportedNumberingRevisionCount !== numberingRevisions.length ||
        supportedMoveRevisionCount !== moveRevisionCount ||
        supportedMoveRevisionCount !== supportedMovePairCount * 2 ||
        [
          'moveFromRangeStart',
          'moveFromRangeEnd',
          'moveToRangeStart',
          'moveToRangeEnd',
          'tblPrChange',
          'trPrChange',
          'tcPrChange',
          'sectPrChange',
        ].some((name) => descendants(document, name).length)
      ) {
        issues.push(
          issue(
            'docx.revisions.structural',
            'Structural revisions',
            'Moved content plus unsupported paragraph-break, character formatting, paragraph formatting, numbering, section, row, cell, and table-property revisions may be normalized; Work currently reviews body-text insertions/deletions and bounded whole-paragraph mark, text-move, character-, paragraph-, and ordered-list-numbering subsets.',
          ),
        );
      }
      if (captionDiagnostics.hasUnsupportedFields) {
        issues.push(
          issue(
            'docx.fields',
            'Fields',
            'Fields beyond supported body fields, citations, bibliographies, caption SEQ fields, and bookmark or caption REF fields are converted to their current displayed value.',
          ),
        );
      }
      const sectionProperties = descendants(document, 'sectPr');
      const columnProperties = descendants(document, 'cols');
      if (sectionProperties.length > 1 || columnProperties.length) {
        issues.push(
          issue(
            'docx.sections',
            'Sections and columns',
            `${sectionProperties.length || 1} section(s) and equal or custom-width column settings are preserved, editable, and applied to document PDF output.`,
            'info',
          ),
        );
      }
      const incompleteUnequalColumns = columnProperties.some(
        (columns) =>
          (attribute(columns, 'equalWidth') === '0' ||
            attribute(columns, 'equalWidth') === 'false') &&
          directChildren(columns, 'col').length === 0,
      );
      const excessiveColumns = columnProperties.some(
        (columns) => Number(attribute(columns, 'num')) > 6,
      );
      if (incompleteUnequalColumns || excessiveColumns) {
        issues.push(
          issue(
            'docx.sections.unsupported',
            'Sections and columns',
            `${[
              incompleteUnequalColumns
                ? 'custom columns without explicit width definitions'
                : '',
              excessiveColumns ? 'sections with more than six columns' : '',
            ]
              .filter(Boolean)
              .join(' and ')} are normalized to at most six editable columns.`,
          ),
        );
      }
      if (
        descendants(document, 'type').some(
          (sectionType) =>
            sectionType.parentElement?.localName === 'sectPr' &&
            attribute(sectionType, 'val') === 'nextColumn',
        )
      ) {
        issues.push(
          issue(
            'docx.sections.next-column-preview',
            'Sections and columns',
            'Next-column section breaks survive DOCX round-trips; Work preview and PDF render them as continuous section blocks.',
          ),
        );
      }
    }
  } catch {
    issues.push(
      issue(
        'docx.inspect',
        'Package inspection',
        'Some DOCX features could not be inspected before conversion.',
      ),
    );
  }

  return report(file, 'DOCX', issues);
}

export async function analyzeSpreadsheetCompatibility(
  file: File,
  extension: string,
  workbook: WorkBook,
  sourcePackage?: OoxmlPackage | null,
  worksheetSummaries?: ReadonlyMap<
    string,
    SpreadsheetWorksheetCompatibilitySummary
  >,
  importedMetadata?: SpreadsheetCompatibilityImportMetadata,
): Promise<WorkCompatibilityReport | null> {
  if (extension === 'csv') return null;
  const sourceFormat = extension.toUpperCase();
  const issues: WorkCompatibilityIssue[] = [];

  if (extension === 'xls') {
    issues.push(
      issue(
        'xls.legacy',
        'Legacy workbook',
        'The binary XLS workbook is converted to the modern native sheet model.',
      ),
    );
  } else if (extension === 'ods') {
    issues.push(
      issue(
        'ods.styles',
        'OpenDocument formatting',
        'Advanced ODS formatting and chart behavior may be normalized on export to XLSX.',
      ),
    );
  }

  const definedNames = workbook.Workbook?.Names ?? [];
  const names = definedNames.filter((name) => !/^_xlnm\./i.test(name.Name));
  if (names.length) {
    issues.push(
      issue(
        'sheet.names',
        'Named ranges',
        `${names.length} workbook or worksheet named range(s) are preserved and editable; external references are not refreshed.`,
        'info',
      ),
    );
  }
  const printAreas = definedNames.filter(
    (name) => name.Name.toLowerCase() === '_xlnm.print_area',
  );
  if (printAreas.length) {
    issues.push(
      issue(
        'sheet.print-area',
        'Print areas',
        `${printAreas.length} worksheet print area(s) are preserved and used by PDF export; disjoint regions may be combined into one bounding layout.`,
        'info',
      ),
    );
  }
  const printTitles = definedNames.filter(
    (name) => name.Name.toLowerCase() === '_xlnm.print_titles',
  );
  const validPrintTitles = printTitles.filter((name) => {
    if (typeof name.Sheet !== 'number') return false;
    const sheetName = workbook.SheetNames[name.Sheet];
    return Boolean(
      sheetName &&
        parseSpreadsheetPrintTitles(
          stripSpreadsheetSheetQualifier(name.Ref, sheetName),
        ),
    );
  });
  if (validPrintTitles.length) {
    issues.push(
      issue(
        'sheet.print-titles',
        'Print titles',
        `${validPrintTitles.length} worksheet print-title setting(s) are preserved, editable, and repeated in PDF output.`,
        'info',
      ),
    );
  }
  if (validPrintTitles.length < printTitles.length) {
    issues.push(
      issue(
        'sheet.print-titles.invalid',
        'Print titles',
        'One or more malformed print-title definitions remain in the original workbook only.',
      ),
    );
  }
  for (const [index, name] of workbook.SheetNames.entries()) {
    const worksheet = workbook.Sheets[name];
    inspectWorksheetModel(
      worksheet,
      name,
      issues,
      extension === 'xlsx',
      worksheetSummaries?.get(name),
    );
    if ((workbook.Workbook?.Sheets?.[index]?.Hidden ?? 0) > 0) {
      issues.push(
        issue(
          'sheet.hidden',
          'Hidden worksheets',
          'Hidden worksheet state is preserved but can be changed by the editor.',
          'info',
          name,
        ),
      );
    }
  }

  if (extension === 'xlsx') {
    try {
      const archive =
        sourcePackage === undefined
          ? await OoxmlPackage.load(await file.arrayBuffer())
          : sourcePackage;
      if (!archive) throw new Error('The XLSX package could not be loaded.');
      await inspectXlsxPackage(archive, issues, importedMetadata);
    } catch {
      issues.push(
        issue(
          'xlsx.inspect',
          'Package inspection',
          'Some XLSX features could not be inspected before conversion.',
        ),
      );
    }
  }

  return issues.length ? report(file, sourceFormat, deduplicate(issues)) : null;
}

function inspectWorksheetModel(
  worksheet: WorkSheet | undefined,
  name: string,
  issues: WorkCompatibilityIssue[],
  nativeXlsxFormulaInspection: boolean,
  importedSummary?: SpreadsheetWorksheetCompatibilitySummary,
) {
  if (!worksheet) return;
  if (worksheet['!autofilter']) {
    issues.push(
      issue(
        'sheet.filter',
        'Filters',
        'The auto-filter range and hidden rows are preserved; advanced active criteria may be normalized.',
        'info',
        name,
      ),
    );
  }
  const {
    hasArrayFormulas,
    hasComments,
    hasCommentThreads,
    hasLinks,
    hasRichText,
  } = importedSummary ?? summarizeSpreadsheetWorksheet(worksheet);
  if (hasComments) {
    issues.push(
      issue(
        'sheet.comments',
        'Cell comments',
        'Plain-text cell comments and one author are preserved and editable through the spreadsheet comment tools.',
        'info',
        name,
      ),
    );
  }
  if (hasCommentThreads) {
    issues.push(
      issue(
        'sheet.comment-threads',
        'Comment threads',
        'Multiple comment blocks and threaded replies are flattened into one editable legacy comment.',
        'warning',
        name,
      ),
    );
  }
  if (hasLinks) {
    issues.push(
      issue(
        'sheet.links',
        'Cell hyperlinks',
        'Web and in-workbook hyperlinks are preserved; advanced screen tips may be normalized.',
        'info',
        name,
      ),
    );
  }
  if (hasArrayFormulas && !nativeXlsxFormulaInspection) {
    issues.push(
      issue(
        'sheet.array-formula',
        'Array formulas',
        'Array-formula ranges may be recalculated as ordinary formulas.',
        'warning',
        name,
      ),
    );
  }
  if (hasRichText) {
    issues.push(
      issue(
        'sheet.rich-text',
        'Rich cell text',
        'Rich text inside a cell is converted to uniform text.',
        'warning',
        name,
      ),
    );
  }
}

function summarizeSpreadsheetWorksheet(
  worksheet: WorkSheet,
): SpreadsheetWorksheetCompatibilitySummary {
  const summary = emptySpreadsheetWorksheetCompatibilitySummary();
  for (const { cell } of xlsxWorksheetCellEntries(worksheet)) {
    updateSpreadsheetWorksheetCompatibilitySummary(summary, cell);
  }
  return summary;
}

async function inspectXlsxPackage(
  archive: OoxmlPackage,
  issues: WorkCompatibilityIssue[],
  importedMetadata?: SpreadsheetCompatibilityImportMetadata,
) {
  issues.push(...(await diagnoseXlsxCharts(archive)));
  issues.push(
    ...(await diagnoseXlsxFormulas(
      archive,
      importedMetadata?.formulaFeatures,
      importedMetadata?.worksheetScans,
    )),
  );
  issues.push(
    ...(await diagnoseXlsxPivots(archive, importedMetadata?.pivotFeatures)),
  );
  const drawingParts = archive
    .paths('xl/drawings/')
    .filter((path) => /^xl\/drawings\/drawing\d+\.xml$/i.test(path));
  let preservedImages = 0;
  let hasUnsupportedImages = false;
  let hasImagesBeyondBudget = false;
  let hasNormalizedImageFormatting = false;
  let hasUnsupportedDrawings = false;
  let embeddedImageBytes = 0;
  for (const part of drawingParts) {
    const drawing = await archive.xml(part);
    const relationships = await archive.relationships(part);
    for (const picture of descendants(drawing, 'pic')) {
      const blip = firstDescendant(picture, 'blip');
      const relationshipId = blip
        ? (attribute(blip, 'r:embed') ?? attribute(blip, 'embed'))
        : null;
      const relationship = relationshipId
        ? relationships.get(relationshipId)
        : undefined;
      const contentType = relationship
        ? contentTypeForPart(relationship.target)
        : '';
      if (
        relationship &&
        relationship.targetMode !== 'External' &&
        relationship.type.endsWith('/image') &&
        archive.has(relationship.target) &&
        isSupportedXlsxWorksheetImageContentType(contentType)
      ) {
        const bytes = await archive.bytes(relationship.target);
        if (
          embeddedImageBytes + bytes.byteLength <=
          MAX_XLSX_WORKSHEET_IMAGE_BYTES
        ) {
          embeddedImageBytes += bytes.byteLength;
          preservedImages += 1;
        } else {
          hasImagesBeyondBudget = true;
        }
      } else {
        hasUnsupportedImages = true;
      }
      const sourceRectangle = firstDescendant(picture, 'srcRect');
      const transform = firstDescendant(directChild(picture, 'spPr'), 'xfrm');
      if (
        (sourceRectangle &&
          Array.from(sourceRectangle.attributes).some(
            (item) => Number(item.value) !== 0,
          )) ||
        (transform &&
          (attribute(transform, 'rot') ||
            attribute(transform, 'flipH') === '1' ||
            attribute(transform, 'flipV') === '1'))
      ) {
        hasNormalizedImageFormatting = true;
      }
    }
    hasUnsupportedDrawings ||= ['sp', 'cxnSp', 'grpSp', 'contentPart'].some(
      (name) => descendants(drawing, name).length > 0,
    );
    hasUnsupportedDrawings ||= descendants(drawing, 'graphicFrame').some(
      (frame) => !firstDescendant(frame, 'chart'),
    );
  }
  if (preservedImages) {
    issues.push(
      issue(
        'xlsx.images',
        'Worksheet images',
        `${preservedImages} embedded raster worksheet image(s), positions, sizes, names, and alternative text are preserved and editable.`,
        'info',
      ),
    );
  }
  if (hasUnsupportedImages) {
    issues.push(
      issue(
        'xlsx.images.unsupported',
        'Worksheet images',
        'Linked, missing, or browser-incompatible worksheet images remain in the original XLSX only.',
      ),
    );
  }
  if (hasImagesBeyondBudget) {
    issues.push(
      issue(
        'xlsx.images.limit',
        'Large worksheet images',
        'Worksheet images beyond the 10 MiB editable image budget remain in the original XLSX only.',
      ),
    );
  }
  if (hasNormalizedImageFormatting) {
    issues.push(
      issue(
        'xlsx.images.format',
        'Worksheet image formatting',
        'Worksheet image crop, rotation, or flip settings are normalized to an editable unrotated image.',
      ),
    );
  }
  if (
    hasUnsupportedDrawings ||
    (archive.paths('xl/drawings/').length > 0 && drawingParts.length === 0)
  ) {
    issues.push(
      issue(
        'xlsx.drawings.unsupported',
        'Worksheet drawings',
        'Worksheet shapes, connectors, SmartArt, and non-chart drawing objects remain in the original XLSX only.',
      ),
    );
  }
  if (
    archive.paths('xl/comments').length ||
    archive.paths('xl/threadedComments/').length
  ) {
    issues.push(
      issue(
        'xlsx.comments',
        'Comments',
        'Legacy cell comments are imported and editable; rich formatting and threaded conversations may normalize.',
        'info',
      ),
    );
  }
  if (archive.paths('xl/tables/').length) {
    issues.push(
      issue(
        'xlsx.tables',
        'Structured tables',
        'Table values are preserved, but table names, styles, and totals are normalized.',
      ),
    );
  }
  if (archive.paths('xl/externalLinks/').length) {
    issues.push(
      issue(
        'xlsx.external-links',
        'External links',
        'External workbook links are not refreshed by Work.',
      ),
    );
  }
  if (archive.has('xl/vbaProject.bin')) {
    issues.push(
      issue(
        'xlsx.macros',
        'Macros',
        'VBA macros remain in the original workbook and are never executed.',
      ),
    );
  }
  let styles: Document | null = null;
  if (archive.has('xl/styles.xml')) {
    styles = await archive.xml('xl/styles.xml');
    const cellFormats = firstDescendant(styles, 'cellXfs');
    if (cellFormats && directChildren(cellFormats, 'xf').length > 1) {
      issues.push(
        issue(
          'xlsx.styles',
          'Cell formatting',
          'Common cell formatting is imported; advanced borders and number formats may be normalized.',
        ),
      );
    }
  }
  for (const part of archive.paths('xl/worksheets/')) {
    if (!part.endsWith('.xml')) continue;
    const scan = importedMetadata?.worksheetScans?.[part];
    if (scan && !scan.hasDiagnosticFeatures) continue;
    const source = await archive.text(part);
    if (
      !scan &&
      !xmlContainsAnyElement(source, XLSX_DIAGNOSTIC_WORKSHEET_ELEMENTS)
    ) {
      continue;
    }
    const worksheet = parseXml(source, part);
    if (descendants(worksheet, 'conditionalFormatting').length) {
      for (const diagnostic of diagnoseXlsxConditionalFormatting(
        worksheet,
        styles,
      )) {
        issues.push(
          issue(
            diagnostic.code,
            'Conditional formatting',
            diagnostic.message,
            diagnostic.severity,
          ),
        );
      }
    }
    if (descendants(worksheet, 'dataValidation').length) {
      issues.push(
        issue(
          'xlsx.validation',
          'Data validation',
          'Common list, numeric, date, and text-length rules are editable; custom formulas and very large ranges may be normalized.',
          'info',
        ),
      );
    }
    if (
      descendants(worksheet, 'sheetProtection').length ||
      descendants(worksheet, 'protectedRange').length
    ) {
      for (const diagnostic of diagnoseXlsxProtection(worksheet, styles)) {
        issues.push(
          issue(
            diagnostic.code,
            'Sheet protection',
            diagnostic.message,
            diagnostic.severity,
          ),
        );
      }
    }
    issues.push(...diagnoseXlsxPageSetup(worksheet));
    if (
      descendants(worksheet, 'brk').some((pageBreak) => {
        const manual = attribute(pageBreak, 'man')?.toLowerCase();
        return manual === '1' || manual === 'true';
      })
    ) {
      issues.push(
        issue(
          'xlsx.manual-page-breaks',
          'Manual page breaks',
          'Manual row and column page breaks are preserved, editable, and honored by PDF pagination.',
          'info',
        ),
      );
    }
  }
}

function docxUnsupportedImageTransformCount(document: Document): number {
  const drawings = [
    ...descendants(document, 'anchor'),
    ...descendants(document, 'inline'),
  ];
  return drawings.filter(
    (drawing) =>
      descendants(drawing, 'pic').length > 0 &&
      !readDocxImageTransform(drawing).supported,
  ).length;
}

function report(
  file: File,
  sourceFormat: string,
  issues: WorkCompatibilityIssue[],
): WorkCompatibilityReport {
  return {
    sourceFormat,
    sourceName: file.name,
    assessedAt: Date.now(),
    issues: deduplicate(issues),
  };
}

function issue(
  code: string,
  feature: string,
  message: string,
  severity: WorkCompatibilityIssue['severity'] = 'warning',
  location?: string,
): WorkCompatibilityIssue {
  return { code, feature, message, severity, location };
}

function deduplicate(
  issues: WorkCompatibilityIssue[],
): WorkCompatibilityIssue[] {
  const seen = new Set<string>();
  return issues.filter((item) => {
    const key = `${item.code}:${item.location ?? ''}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}
