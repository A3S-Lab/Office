import type {
  WorkSpreadsheetDataValidationErrorStyle,
  WorkSpreadsheetDataValidationItem,
} from '../work-types';

export type SpreadsheetDataValidationInteraction = 'notice' | 'confirm';

export interface SpreadsheetDataValidationEditRequest {
  column: number;
  failureText: string;
  interaction: SpreadsheetDataValidationInteraction;
  item: WorkSpreadsheetDataValidationItem;
  row: number;
  sheetId: string;
  value: unknown;
}

type SpreadsheetDataValidationMessageSource = Pick<
  WorkSpreadsheetDataValidationItem,
  'errorMessage' | 'errorTitle'
>;

export function spreadsheetDataValidationErrorStyle(
  item: Pick<WorkSpreadsheetDataValidationItem, 'errorStyle'>,
): WorkSpreadsheetDataValidationErrorStyle {
  return item.errorStyle === 'warning' || item.errorStyle === 'information'
    ? item.errorStyle
    : 'stop';
}

export function spreadsheetDataValidationInteraction(
  style: WorkSpreadsheetDataValidationErrorStyle,
): SpreadsheetDataValidationInteraction {
  return style === 'stop' ? 'notice' : 'confirm';
}

export function spreadsheetDataValidationDialogTitle(
  item: Pick<WorkSpreadsheetDataValidationItem, 'errorStyle' | 'errorTitle'>,
): string {
  const customTitle = item.errorTitle?.trim();
  if (customTitle) return customTitle;
  switch (spreadsheetDataValidationErrorStyle(item)) {
    case 'warning':
      return '数据验证警告';
    case 'information':
      return '数据验证提示';
    default:
      return '数据验证错误';
  }
}

export function spreadsheetDataValidationDialogDescription(
  failureText: string,
  value: unknown,
  item?: SpreadsheetDataValidationMessageSource,
): string {
  const customTitle = item?.errorTitle?.trim();
  const customMessage = item?.errorMessage?.trim();
  let message = customMessage || failureText.trim();
  if (customTitle && message.startsWith(`${customTitle}\n`)) {
    message = message.slice(customTitle.length + 1).trim();
  }
  message ||= '当前输入不符合此单元格的数据验证规则。';
  const valueText = spreadsheetDataValidationValueText(value);
  return valueText ? `${message}\n\n当前输入：${valueText}` : message;
}

export function spreadsheetDataValidationConfirmLabels(
  style: WorkSpreadsheetDataValidationErrorStyle,
): { cancelLabel: string; confirmLabel: string } {
  return style === 'warning'
    ? { cancelLabel: '取消', confirmLabel: '继续输入' }
    : { cancelLabel: '返回修改', confirmLabel: '保留输入' };
}

export function spreadsheetDataValidationValueText(value: unknown): string {
  if (typeof value === 'string') return value.trim();
  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }
  return '';
}

export function cloneSpreadsheetDataValidationItem(
  item: WorkSpreadsheetDataValidationItem,
): WorkSpreadsheetDataValidationItem {
  return {
    ...item,
    errorTitle: item.errorTitle,
    errorMessage: item.errorMessage,
    hintTitle: item.hintTitle,
  };
}
