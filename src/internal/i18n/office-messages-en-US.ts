import type { OfficeHostMessageKey } from './office-messages';

export const officeMessagesEnUS: Record<OfficeHostMessageKey, string> = {
  'editor.loading.document': 'Opening document editor',
  'editor.loading.markdown': 'Opening Markdown editor',
  'editor.loading.spreadsheet': 'Opening spreadsheet editor',
  'editor.loading.presentation': 'Opening presentation editor',
  'editor.loading.pdf': 'Opening PDF',
  'editor.error.failed': '{title} failed to load.',
  'editor.error.retry': 'Retry',
  'document.lineHeight.default': 'Default line spacing',
  'document.lineHeight.single': 'Single',
  'document.lineHeight.1_15': '1.15',
  'document.lineHeight.1_5': '1.5',
  'document.lineHeight.double': 'Double',
  'document.lineHeight.custom': '{value}',
};
