import { documentUiMessagesEnUS } from './document-ui-messages-en-US';
import { documentUiMessagesZhCN } from './document-ui-messages-zh-CN';
import { officeChromeUiMessagesEnUS } from './office-chrome-ui-messages-en-US';
import { officeChromeUiMessagesZhCN } from './office-chrome-ui-messages-zh-CN';
import {
  OFFICE_DEFAULT_LOCALE,
  type OfficeLocale,
  type OfficeMessageCatalog,
  type OfficeMessageKey,
  type OfficeMessagesOverride,
} from './office-messages';
import { officeMessagesEnUS } from './office-messages-en-US';
import { officeMessagesZhCN } from './office-messages-zh-CN';
import { markdownUiMessagesEnUS } from './markdown-ui-messages-en-US';
import { markdownUiMessagesZhCN } from './markdown-ui-messages-zh-CN';
import { pdfUiMessagesEnUS } from './pdf-ui-messages-en-US';
import { pdfUiMessagesZhCN } from './pdf-ui-messages-zh-CN';
import { presentationUiMessagesEnUS } from './presentation-ui-messages-en-US';
import { presentationUiMessagesZhCN } from './presentation-ui-messages-zh-CN';
import { spreadsheetUiMessagesEnUS } from './spreadsheet-ui-messages-en-US';
import { spreadsheetUiMessagesZhCN } from './spreadsheet-ui-messages-zh-CN';

export type {
  OfficeLocale,
  OfficeMessageCatalog,
  OfficeMessageKey,
  OfficeMessagesOverride,
};
export { OFFICE_DEFAULT_LOCALE };

const catalogs: Record<OfficeLocale, OfficeMessageCatalog> = {
  'zh-CN': {
    ...officeMessagesZhCN,
    ...officeChromeUiMessagesZhCN,
    ...documentUiMessagesZhCN,
    ...spreadsheetUiMessagesZhCN,
    ...presentationUiMessagesZhCN,
    ...markdownUiMessagesZhCN,
    ...pdfUiMessagesZhCN,
  },
  'en-US': {
    ...officeMessagesEnUS,
    ...officeChromeUiMessagesEnUS,
    ...documentUiMessagesEnUS,
    ...spreadsheetUiMessagesEnUS,
    ...presentationUiMessagesEnUS,
    ...markdownUiMessagesEnUS,
    ...pdfUiMessagesEnUS,
  },
};

export function resolveOfficeLocale(locale: string | undefined): OfficeLocale {
  if (locale === 'en-US' || locale === 'zh-CN') return locale;
  return OFFICE_DEFAULT_LOCALE;
}

export function resolveOfficeMessages(options?: {
  locale?: string;
  messages?: OfficeMessagesOverride;
}): OfficeMessageCatalog {
  const locale = resolveOfficeLocale(options?.locale);
  const base = catalogs[locale];
  if (!options?.messages) return base;
  return { ...base, ...options.messages };
}

export function officeMessage(
  catalog: OfficeMessageCatalog,
  key: OfficeMessageKey,
  vars?: Record<string, string>,
): string {
  const template = catalog[key];
  if (!vars) return template;
  return template.replace(/\{(\w+)\}/g, (_, name: string) => vars[name] ?? '');
}
