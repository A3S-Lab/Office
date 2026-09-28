/**
 * Built-in OOXML pivot table styles exposed for daily Traditional Office / WPS
 * authoring. Keep the catalog explicit rather than inventing custom styles.
 */
import { officeMessage, resolveOfficeMessages } from '../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../i18n/office-messages';

const PIVOT_STYLE_VALUES = [
  'PivotStyleLight1',
  'PivotStyleLight8',
  'PivotStyleLight16',
  'PivotStyleLight17',
  'PivotStyleLight18',
  'PivotStyleLight19',
  'PivotStyleLight20',
  'PivotStyleLight21',
  'PivotStyleLight28',
  'PivotStyleMedium2',
  'PivotStyleMedium7',
  'PivotStyleMedium9',
  'PivotStyleMedium10',
  'PivotStyleMedium11',
  'PivotStyleMedium12',
  'PivotStyleMedium28',
  'PivotStyleDark1',
  'PivotStyleDark2',
  'PivotStyleDark3',
  'PivotStyleDark4',
  'PivotStyleDark9',
  'PivotStyleDark10',
  'PivotStyleDark11',
  'PivotStyleDark28',
] as const;

export type WorkSpreadsheetPivotStyleName = (typeof PIVOT_STYLE_VALUES)[number];

export const WORK_SPREADSHEET_DEFAULT_PIVOT_STYLE: WorkSpreadsheetPivotStyleName =
  'PivotStyleLight16';

export function workSpreadsheetPivotStyleLabel(
  value: WorkSpreadsheetPivotStyleName,
  messages: OfficeMessageCatalog = resolveOfficeMessages(),
): string {
  return officeMessage(
    messages,
    `spreadsheet.pivot.style.${value}` as Parameters<typeof officeMessage>[1],
  );
}

export const WORK_SPREADSHEET_PIVOT_STYLE_OPTIONS = PIVOT_STYLE_VALUES.map(
  (value) => ({
    value,
    get label() {
      return workSpreadsheetPivotStyleLabel(value);
    },
  }),
);

export function isWorkSpreadsheetPivotStyleName(
  value: string,
): value is WorkSpreadsheetPivotStyleName {
  return PIVOT_STYLE_VALUES.some((option) => option === value);
}
