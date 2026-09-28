import { createContext, use, type ReactNode } from 'react';
import {
  officeMessage,
  resolveOfficeMessages,
} from '../../../i18n/office-locale';
import type {
  OfficeMessageCatalog,
  OfficeMessageKey,
} from '../../../i18n/office-messages';

const OfficeMessagesContext = createContext<OfficeMessageCatalog | null>(null);

export function OfficeMessagesProvider({
  children,
  catalog,
}: {
  children: ReactNode;
  catalog: OfficeMessageCatalog;
}) {
  return (
    <OfficeMessagesContext.Provider value={catalog}>
      {children}
    </OfficeMessagesContext.Provider>
  );
}

export function useOfficeMessages(): OfficeMessageCatalog {
  const catalog = use(OfficeMessagesContext);
  if (!catalog) {
    return resolveOfficeMessages();
  }
  return catalog;
}

export function useOfficeMessage(
  key: OfficeMessageKey,
  vars?: Record<string, string>,
): string {
  return officeMessage(useOfficeMessages(), key, vars);
}
