import type { Editor } from '@tiptap/core';
import { ALargeSmall, CaseUpper, Type } from 'lucide-react';
import { Popover } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import {
  normalizeDocumentTextCase,
  type WorkDocumentTextCase,
} from '../work-document-text-case';
import { getDocumentCommandDefinition } from './document-command-catalog';
import { documentCommandLabel } from './document-command-i18n';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import { useOfficeMessages } from './office-messages-context';

export function DocumentTextCaseRibbon({ editor }: { editor: Editor }) {
  const messages = useOfficeMessages();
  const textCaseOptions = [
    {
      value: 'none' as const,
      label: messages['document.textCase.none'],
      icon: Type,
    },
    {
      value: 'all-caps' as const,
      label: documentCommandLabel('allCaps', messages),
      icon: CaseUpper,
      command: 'allCaps' as const,
    },
    {
      value: 'small-caps' as const,
      label: documentCommandLabel('smallCaps', messages),
      icon: ALargeSmall,
      command: 'smallCaps' as const,
    },
  ];
  const value =
    normalizeDocumentTextCase(editor.getAttributes('textStyle').textCase) ??
    'none';
  const current =
    textCaseOptions.find((option) => option.value === value) ??
    textCaseOptions[0];
  const changeCaseShortcut =
    getDocumentCommandDefinition('changeCase').shortcut;
  const menuLabel = messages['document.textCase.menu'];

  return (
    <Popover
      label={menuLabel}
      panelLabel={menuLabel}
      panelRole="menu"
      portal
      placement="bottom-end"
      className="work-document-text-case-root"
      panelClassName="work-office-context-menu work-document-text-case-menu"
      focusFirstOnOpen
      onPanelKeyDown={moveOfficeMenuFocus}
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          type="button"
          className={value !== 'none' || open ? 'active' : ''}
          aria-keyshortcuts={changeCaseShortcut?.aria}
          title={officeMessage(messages, 'document.textCase.title', {
            current: current.label,
            shortcut: changeCaseShortcut
              ? officeMessage(messages, 'document.textCase.titleShortcut', {
                  shortcut: changeCaseShortcut.label,
                })
              : '',
          })}
        >
          <ALargeSmall size={16} aria-hidden="true" />
        </button>
      )}
    >
      {(close) =>
        textCaseOptions.map((option) => {
          const Icon = option.icon;
          const shortcut =
            option.value === 'all-caps' || option.value === 'small-caps'
              ? getDocumentCommandDefinition(
                  option.value === 'all-caps' ? 'allCaps' : 'smallCaps',
                ).shortcut
              : undefined;
          return (
            <button
              key={option.value}
              type="button"
              role="menuitemradio"
              tabIndex={-1}
              aria-label={option.label}
              aria-checked={value === option.value}
              aria-keyshortcuts={shortcut?.aria}
              onClick={() => {
                close();
                editor.commands.setDocumentTextCase(
                  option.value as WorkDocumentTextCase,
                );
              }}
            >
              <Icon aria-hidden="true" />
              <span>{option.label}</span>
              {shortcut && <kbd>{shortcut.label}</kbd>}
            </button>
          );
        })
      }
    </Popover>
  );
}
