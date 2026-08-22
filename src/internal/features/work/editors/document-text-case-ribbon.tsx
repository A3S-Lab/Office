import type { Editor } from '@tiptap/core';
import { ALargeSmall, CaseUpper, Type } from 'lucide-react';
import { Popover } from '../../../design-system/primitives';
import {
  normalizeDocumentTextCase,
  type WorkDocumentTextCase,
} from '../work-document-text-case';
import { getDocumentCommandDefinition } from './document-command-catalog';
import { moveOfficeMenuFocus } from './office-menu-keyboard';

const textCaseOptions = [
  { value: 'none', label: '常规', icon: Type },
  {
    value: 'all-caps',
    label: '全部大写',
    icon: CaseUpper,
    command: 'allCaps',
  },
  {
    value: 'small-caps',
    label: '小型大写',
    icon: ALargeSmall,
    command: 'smallCaps',
  },
] as const;

export function DocumentTextCaseRibbon({ editor }: { editor: Editor }) {
  const value =
    normalizeDocumentTextCase(editor.getAttributes('textStyle').textCase) ??
    'none';
  const current =
    textCaseOptions.find((option) => option.value === value) ??
    textCaseOptions[0];

  return (
    <Popover
      label="大小写效果"
      panelLabel="大小写效果"
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
          aria-pressed={value !== 'none'}
          title={`大小写效果（${current.label}）`}
        >
          <ALargeSmall size={16} aria-hidden="true" />
        </button>
      )}
    >
      {(close) =>
        textCaseOptions.map((option) => {
          const Icon = option.icon;
          const shortcut =
            'command' in option
              ? getDocumentCommandDefinition(option.command).shortcut
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
