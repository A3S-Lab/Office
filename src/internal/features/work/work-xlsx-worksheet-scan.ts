export interface XlsxWorksheetXmlScan {
  hasDirectCellStyles: boolean;
  hasDiagnosticFeatures: boolean;
  hasFormulaFeatures: boolean;
  hasImportedFeatures: boolean;
  hasRichTextCells: boolean;
  requiresSheetJsCellStyles: boolean;
}

const IMPORTED_WORKSHEET_ELEMENTS = new Set([
  'pane',
  'autoFilter',
  'dataValidation',
  'conditionalFormatting',
  'sheetProtection',
  'protectedRange',
  'rowBreaks',
  'colBreaks',
  'pageSetup',
  'pageMargins',
  'printOptions',
  'headerFooter',
  'pageSetUpPr',
  'drawing',
  'tableParts',
]);

const DIAGNOSTIC_WORKSHEET_ELEMENTS = new Set([
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
]);

const WORKSHEET_SCAN_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?(?:(cols|f|pane|autoFilter|dataValidation|conditionalFormatting|sheetProtection|protectedRange|rowBreaks|colBreaks|pageSetup|pageMargins|printOptions|headerFooter|pageSetUpPr|drawing|tableParts)(?=[\s/>])|(row)(?=[\s/>])[^>]*\s(?:collapsed|customFormat|customHeight|hidden|ht|outlineLevel|thickBot|thickTop)\s*=|(c)(?=[\s/>])[^>]*\ss\s*=|(r)(?=[\s/>]))/g;

/**
 * Builds every worksheet gate in one pass over the decompressed XML. The
 * scanner deliberately permits false positives because they only select the
 * existing full parser path; a false negative could discard workbook state.
 */
export function scanXlsxWorksheetXml(source: string): XlsxWorksheetXmlScan {
  let hasDirectCellStyles = false;
  let hasDiagnosticFeatures = false;
  let hasFormulaFeatures = false;
  let hasImportedFeatures = false;
  let hasRichTextCells = false;
  let requiresSheetJsCellStyles = false;
  WORKSHEET_SCAN_PATTERN.lastIndex = 0;

  for (let match = WORKSHEET_SCAN_PATTERN.exec(source); match; ) {
    const element = match[1];
    if (match[3]) hasDirectCellStyles = true;
    if (match[4]) hasRichTextCells = true;
    if (match[2] || match[3] || element === 'cols') {
      requiresSheetJsCellStyles = true;
    }
    if (element === 'f') hasFormulaFeatures = true;
    if (element && IMPORTED_WORKSHEET_ELEMENTS.has(element)) {
      hasImportedFeatures = true;
    }
    if (element && DIAGNOSTIC_WORKSHEET_ELEMENTS.has(element)) {
      hasDiagnosticFeatures = true;
    }
    if (
      hasDirectCellStyles &&
      hasDiagnosticFeatures &&
      hasFormulaFeatures &&
      hasImportedFeatures &&
      hasRichTextCells &&
      requiresSheetJsCellStyles
    ) {
      break;
    }
    match = WORKSHEET_SCAN_PATTERN.exec(source);
  }
  WORKSHEET_SCAN_PATTERN.lastIndex = 0;

  return {
    hasDirectCellStyles,
    hasDiagnosticFeatures,
    hasFormulaFeatures,
    hasImportedFeatures,
    hasRichTextCells,
    requiresSheetJsCellStyles,
  };
}
