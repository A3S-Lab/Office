import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import {
  type DocumentCommandId,
  type DocumentRibbonTabId,
  getDocumentCommandDefinition,
} from './document-command-catalog';

export function documentCommandLabel(
  commandId: DocumentCommandId,
  messages: OfficeMessageCatalog,
): string {
  return messages[`document.command.${commandId}`];
}

export function documentCommandTitleWithShortcut(
  commandId: DocumentCommandId,
  messages: OfficeMessageCatalog,
): string {
  const label = documentCommandLabel(commandId, messages);
  const shortcut = getDocumentCommandDefinition(commandId).shortcut?.label;
  if (!shortcut) return label;
  return officeMessage(messages, 'document.command.withShortcut', {
    label,
    shortcut,
  });
}

export function documentRibbonTabLabel(
  tabId: DocumentRibbonTabId,
  messages: OfficeMessageCatalog,
  options?: { compact?: boolean },
): string {
  if (options?.compact) {
    const compactKey = `document.ribbon.${tabId}.compact` as const;
    if (compactKey in messages) {
      return messages[compactKey as keyof OfficeMessageCatalog];
    }
  }
  return messages[`document.ribbon.${tabId}`];
}

export type LocalizedDocumentRibbonTab = {
  id: DocumentRibbonTabId;
  label: string;
  compactLabel?: string;
  contextual?: boolean;
};

export function localizeDocumentRibbonTab(
  tab: {
    id: DocumentRibbonTabId;
    hasCompactLabel?: boolean;
    contextual?: boolean;
  },
  messages: OfficeMessageCatalog,
): LocalizedDocumentRibbonTab {
  return {
    id: tab.id,
    label: documentRibbonTabLabel(tab.id, messages),
    ...(tab.hasCompactLabel
      ? {
          compactLabel: documentRibbonTabLabel(tab.id, messages, {
            compact: true,
          }),
        }
      : {}),
    ...(tab.contextual ? { contextual: true } : {}),
  };
}
