import type { OfficeChromeUiMessageKey } from './office-chrome-ui-messages-zh-CN';
import type { DocumentUiMessageKey } from './document-ui-messages-zh-CN';
import type { MarkdownUiMessageKey } from './markdown-ui-messages-zh-CN';
import type { PdfUiMessageKey } from './pdf-ui-messages-zh-CN';
import type { PresentationUiMessageKey } from './presentation-ui-messages-zh-CN';
import type { SpreadsheetUiMessageKey } from './spreadsheet-ui-messages-zh-CN';

export type OfficeLocale = 'zh-CN' | 'en-US';

export const OFFICE_DEFAULT_LOCALE: OfficeLocale = 'zh-CN';

/**
 * Host-facing UI strings. Keys are stable; values are locale-specific.
 * Editor chrome migration adds keys here before removing CJK literals from
 * `.tsx` sources.
 */
export type OfficeHostMessageKey =
  | 'editor.loading.document'
  | 'editor.loading.markdown'
  | 'editor.loading.spreadsheet'
  | 'editor.loading.presentation'
  | 'editor.loading.pdf'
  | 'editor.error.failed'
  | 'editor.error.retry'
  | 'document.lineHeight.default'
  | 'document.lineHeight.single'
  | 'document.lineHeight.1_15'
  | 'document.lineHeight.1_5'
  | 'document.lineHeight.double'
  | 'document.lineHeight.custom';

export type OfficeMessageKey =
  | OfficeHostMessageKey
  | OfficeChromeUiMessageKey
  | DocumentUiMessageKey
  | SpreadsheetUiMessageKey
  | PresentationUiMessageKey
  | MarkdownUiMessageKey
  | PdfUiMessageKey;

export type OfficeMessageCatalog = Record<OfficeMessageKey, string>;

export type OfficeMessagesOverride = Partial<OfficeMessageCatalog>;
