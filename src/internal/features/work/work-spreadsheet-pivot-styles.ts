/**
 * Built-in OOXML pivot table styles exposed for daily Traditional Office / WPS
 * authoring. Keep the catalog explicit rather than inventing custom styles.
 */
export const WORK_SPREADSHEET_PIVOT_STYLE_OPTIONS = [
  { value: 'PivotStyleLight1', label: '浅色 1' },
  { value: 'PivotStyleLight8', label: '浅色 8' },
  { value: 'PivotStyleLight16', label: '浅色 16' },
  { value: 'PivotStyleLight17', label: '浅色 17' },
  { value: 'PivotStyleLight18', label: '浅色 18' },
  { value: 'PivotStyleLight19', label: '浅色 19' },
  { value: 'PivotStyleLight20', label: '浅色 20' },
  { value: 'PivotStyleLight21', label: '浅色 21' },
  { value: 'PivotStyleLight28', label: '浅色 28' },
  { value: 'PivotStyleMedium2', label: '中等 2' },
  { value: 'PivotStyleMedium7', label: '中等 7' },
  { value: 'PivotStyleMedium9', label: '中等 9' },
  { value: 'PivotStyleMedium10', label: '中等 10' },
  { value: 'PivotStyleMedium11', label: '中等 11' },
  { value: 'PivotStyleMedium12', label: '中等 12' },
  { value: 'PivotStyleMedium28', label: '中等 28' },
  { value: 'PivotStyleDark1', label: '深色 1' },
  { value: 'PivotStyleDark2', label: '深色 2' },
  { value: 'PivotStyleDark3', label: '深色 3' },
  { value: 'PivotStyleDark4', label: '深色 4' },
  { value: 'PivotStyleDark9', label: '深色 9' },
  { value: 'PivotStyleDark10', label: '深色 10' },
  { value: 'PivotStyleDark11', label: '深色 11' },
  { value: 'PivotStyleDark28', label: '深色 28' },
] as const;

export type WorkSpreadsheetPivotStyleName =
  (typeof WORK_SPREADSHEET_PIVOT_STYLE_OPTIONS)[number]['value'];

export const WORK_SPREADSHEET_DEFAULT_PIVOT_STYLE: WorkSpreadsheetPivotStyleName =
  'PivotStyleLight16';

export function isWorkSpreadsheetPivotStyleName(
  value: string,
): value is WorkSpreadsheetPivotStyleName {
  return WORK_SPREADSHEET_PIVOT_STYLE_OPTIONS.some(
    (option) => option.value === value,
  );
}
