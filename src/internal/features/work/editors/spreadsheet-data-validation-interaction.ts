import type {
  WorkSpreadsheetDataValidationErrorStyle,
  WorkSpreadsheetDataValidationItem,
} from '../work-types';
import { officeMessage, resolveOfficeMessages } from '../../../i18n/office-locale';

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
  const messages = resolveOfficeMessages();
  switch (spreadsheetDataValidationErrorStyle(item)) {
    case 'warning':
      return officeMessage(messages, 'spreadsheet.dv.alert.warningTitle');
    case 'information':
      return officeMessage(messages, 'spreadsheet.dv.alert.infoTitle');
    default:
      return officeMessage(messages, 'spreadsheet.dv.alert.errorTitle');
  }
}

export function spreadsheetDataValidationDialogDescription(
  failureText: string,
  value: unknown,
  item?: SpreadsheetDataValidationMessageSource,
): string {
  const catalog = resolveOfficeMessages();
  const customTitle = item?.errorTitle?.trim();
  const customMessage = item?.errorMessage?.trim();
  let message = customMessage || failureText.trim();
  if (customTitle && message.startsWith(`${customTitle}\n`)) {
    message = message.slice(customTitle.length + 1).trim();
  }
  message ||= officeMessage(catalog, 'spreadsheet.dv.alert.defaultMessage');
  const valueText = spreadsheetDataValidationValueText(value);
  if (!valueText) return message;
  const currentInput = officeMessage(
    catalog,
    'spreadsheet.dv.alert.currentInput',
    { value: valueText },
  );
  return `${message}\n\n${currentInput}`;
}

export function spreadsheetDataValidationConfirmLabels(
  style: WorkSpreadsheetDataValidationErrorStyle,
): { cancelLabel: string; confirmLabel: string } {
  const messages = resolveOfficeMessages();
  return style === 'warning'
    ? {
        cancelLabel: officeMessage(
          messages,
          'spreadsheet.dv.alert.cancelContinue',
        ),
        confirmLabel: officeMessage(
          messages,
          'spreadsheet.dv.alert.confirmContinue',
        ),
      }
    : {
        cancelLabel: officeMessage(
          messages,
          'spreadsheet.dv.alert.cancelRevise',
        ),
        confirmLabel: officeMessage(
          messages,
          'spreadsheet.dv.alert.confirmKeep',
        ),
      };
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
