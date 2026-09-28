import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import {
  type SpreadsheetCommandId,
  type SpreadsheetRibbonTabId,
  getSpreadsheetCommandDefinition,
  spreadsheetRibbonTabs,
  spreadsheetTableDesignRibbonTab,
} from './spreadsheet-command-catalog';

export function spreadsheetCommandLabel(
  commandId: SpreadsheetCommandId,
  messages: OfficeMessageCatalog,
): string {
  return messages[`spreadsheet.command.${commandId}`];
}

export function spreadsheetCommandTitleWithShortcut(
  commandId: SpreadsheetCommandId,
  messages: OfficeMessageCatalog,
): string {
  const label = spreadsheetCommandLabel(commandId, messages);
  const shortcut = getSpreadsheetCommandDefinition(commandId).shortcut?.label;
  if (!shortcut) return label;
  return officeMessage(messages, 'spreadsheet.command.withShortcut', {
    label,
    shortcut,
  });
}

export function spreadsheetNumberFormatLabel(
  preset: string,
  messages: OfficeMessageCatalog,
): string {
  return messages[`spreadsheet.numberFormat.${preset}` as keyof OfficeMessageCatalog];
}

export function spreadsheetRibbonTabLabel(
  tabId: SpreadsheetRibbonTabId,
  messages: OfficeMessageCatalog,
  options?: { compact?: boolean },
): string {
  if (options?.compact) {
    const compactKey = `spreadsheet.ribbon.${tabId}.compact` as const;
    if (compactKey in messages) {
      return messages[compactKey as keyof OfficeMessageCatalog];
    }
  }
  return messages[`spreadsheet.ribbon.${tabId}`];
}

export type LocalizedSpreadsheetRibbonTab = {
  id: SpreadsheetRibbonTabId;
  label: string;
  compactLabel?: string;
  contextual?: boolean;
};

export function localizeSpreadsheetRibbonTab(
  tab: {
    id: SpreadsheetRibbonTabId;
    hasCompactLabel?: boolean;
    contextual?: boolean;
  },
  messages: OfficeMessageCatalog,
): LocalizedSpreadsheetRibbonTab {
  return {
    id: tab.id,
    label: spreadsheetRibbonTabLabel(tab.id, messages),
    ...(tab.hasCompactLabel
      ? {
          compactLabel: spreadsheetRibbonTabLabel(tab.id, messages, {
            compact: true,
          }),
        }
      : {}),
    ...(tab.contextual ? { contextual: true } : {}),
  };
}

export function localizedSpreadsheetRibbonTabs(
  messages: OfficeMessageCatalog,
  options?: { includeTableDesign?: boolean },
): LocalizedSpreadsheetRibbonTab[] {
  const tabs = spreadsheetRibbonTabs.map((tab) =>
    localizeSpreadsheetRibbonTab(
      {
        id: tab.id,
        hasCompactLabel:
          'hasCompactLabel' in tab
            ? Boolean(
                (tab as { hasCompactLabel?: boolean }).hasCompactLabel,
              )
            : 'compactLabel' in tab,
      },
      messages,
    ),
  );
  if (options?.includeTableDesign) {
    tabs.push(
      localizeSpreadsheetRibbonTab(
        {
          id: spreadsheetTableDesignRibbonTab.id,
          hasCompactLabel: true,
          contextual: true,
        },
        messages,
      ),
    );
  }
  return tabs;
}
