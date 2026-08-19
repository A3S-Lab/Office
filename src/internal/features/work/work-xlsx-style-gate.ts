import { xmlContainsAnyElement } from './work-ooxml-package';

const XLSX_CUSTOM_ROW_LAYOUT_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?row(?=[\s/>])[^>]*\s(?:collapsed|customFormat|customHeight|hidden|ht|outlineLevel|thickBot|thickTop)\s*=/;

export function xlsxWorksheetRequiresSheetJsCellStyles(
  source: string,
): boolean {
  return (
    xmlContainsAnyElement(source, ['cols']) ||
    XLSX_CUSTOM_ROW_LAYOUT_PATTERN.test(source)
  );
}
