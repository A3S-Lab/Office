export const spreadsheetUnderlineStyles = [
  'none',
  'single',
  'double',
  'singleAccounting',
  'doubleAccounting',
] as const;

export type SpreadsheetUnderlineStyle =
  (typeof spreadsheetUnderlineStyles)[number];

export type SpreadsheetUnderlineCellValue = 0 | 1 | 2 | 3 | 4;

const underlineCellValues = {
  none: 0,
  single: 1,
  double: 2,
  singleAccounting: 3,
  doubleAccounting: 4,
} as const satisfies Record<
  SpreadsheetUnderlineStyle,
  SpreadsheetUnderlineCellValue
>;

export function isSpreadsheetUnderlineStyle(
  value: unknown,
): value is SpreadsheetUnderlineStyle {
  return (
    typeof value === 'string' &&
    spreadsheetUnderlineStyles.includes(value as SpreadsheetUnderlineStyle)
  );
}

export function spreadsheetUnderlineCellValue(
  style: SpreadsheetUnderlineStyle,
): SpreadsheetUnderlineCellValue {
  return underlineCellValues[style];
}

export function spreadsheetUnderlineStyle(
  value: unknown,
): SpreadsheetUnderlineStyle {
  switch (Number(value)) {
    case 1:
      return 'single';
    case 2:
      return 'double';
    case 3:
    case 0x21:
      return 'singleAccounting';
    case 4:
    case 0x22:
      return 'doubleAccounting';
    default:
      return 'none';
  }
}

export function spreadsheetUnderlineCellValueFromXlsx(
  value: string | null,
): SpreadsheetUnderlineCellValue {
  switch (value?.trim()) {
    case null:
    case undefined:
    case '':
    case 'single':
      return 1;
    case 'double':
      return 2;
    case 'singleAccounting':
      return 3;
    case 'doubleAccounting':
      return 4;
    default:
      return 0;
  }
}

export function spreadsheetUnderlineCellValueFromSheetJs(
  value: unknown,
): SpreadsheetUnderlineCellValue {
  if (value === true) return 1;
  if (isSpreadsheetUnderlineStyle(value)) {
    return spreadsheetUnderlineCellValue(value);
  }
  return spreadsheetUnderlineCellValue(spreadsheetUnderlineStyle(value));
}
