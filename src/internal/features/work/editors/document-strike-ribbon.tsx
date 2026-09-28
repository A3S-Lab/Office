import type { Editor } from '@tiptap/core';
import { ChevronDown, Strikethrough } from 'lucide-react';
import {
  type ReactNode,
  useCallback,
  useMemo,
  useSyncExternalStore,
} from 'react';
import { Popover } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import {
  documentStrikeStyle,
  type WorkDocumentStrikeStyle,
} from '../work-document-strike';
import { getDocumentCommandDefinition } from './document-command-catalog';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import { useOfficeMessages } from './office-messages-context';

const strikeOptionDefs = [
  { value: 'none', messageKey: 'document.strike.none' },
  {
    value: 'single',
    messageKey: 'document.strike.single',
    command: 'strike' as const,
  },
  {
    value: 'double',
    messageKey: 'document.strike.double',
    command: 'doubleStrike' as const,
  },
] as const satisfies readonly {
  value: WorkDocumentStrikeStyle;
  messageKey: keyof OfficeMessageCatalog;
  command?: 'strike' | 'doubleStrike';
}[];

export function DocumentStrikeRibbon({
  editor,
  label,
  menuLabel,
  className = '',
}: {
  editor: Editor;
  label?: string;
  menuLabel?: string;
  className?: string;
}) {
  const messages = useOfficeMessages();
  const resolvedLabel = label ?? messages['document.strike.label'];
  const resolvedMenuLabel = menuLabel ?? messages['document.strike.menu'];
  const strikeOptions = useMemo(
    () =>
      strikeOptionDefs.map((option) => ({
        ...option,
        label: messages[option.messageKey],
      })),
    [messages],
  );
  const subscribe = useCallback(
    (notify: () => void) => {
      if (editor.isDestroyed) return () => undefined;
      editor.on('transaction', notify);
      return () => editor.off('transaction', notify);
    },
    [editor],
  );
  useSyncExternalStore(
    subscribe,
    () => editor.state,
    () => editor.state,
  );
  if (editor.isDestroyed) return null;
  const style = documentStrikeStyle(editor);
  const active = style !== 'none';
  const currentLabel =
    strikeOptions.find((option) => option.value === style)?.label ??
    strikeOptions[0].label;

  return (
    <span
      className={`work-document-strike-control${className ? ` ${className}` : ''}`}
    >
      <Popover
        label={officeMessage(messages, 'document.strike.more', {
          label: resolvedLabel,
        })}
        panelLabel={resolvedMenuLabel}
        panelRole="menu"
        portal
        placement="bottom-end"
        className="work-document-strike-split-root"
        panelClassName="work-office-context-menu work-document-strike-menu"
        focusFirstOnOpen
        onPanelKeyDown={moveOfficeMenuFocus}
        trigger={(triggerProps, { open }) => (
          <>
            <button
              type="button"
              className={`work-document-strike-primary${active ? ' active' : ''}`}
              aria-label={resolvedLabel}
              aria-pressed={active}
              title={officeMessage(messages, 'document.strike.title', {
                label: resolvedLabel,
                current: currentLabel,
              })}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => {
                // Keep ribbon focus (e.g. disclosure) — chain().focus() steals
                // into the editor and breaks L2 formatting loops.
                editor.commands.toggleStrike();
              }}
            >
              <Strikethrough size={16} aria-hidden="true" />
            </button>
            <button
              {...triggerProps}
              type="button"
              className={`work-document-strike-disclosure${open ? ' active' : ''}`}
              title={officeMessage(messages, 'document.strike.more', {
                label: resolvedLabel,
              })}
              onMouseDown={(event) => event.preventDefault()}
            >
              <ChevronDown size={11} aria-hidden="true" />
            </button>
          </>
        )}
      >
        {(close) =>
          strikeOptions.map((option) => {
            const optionShortcut =
              'command' in option && option.command
                ? getDocumentCommandDefinition(option.command).shortcut
                : undefined;
            return (
              <button
                key={option.value}
                type="button"
                role="menuitemradio"
                tabIndex={-1}
                aria-label={option.label}
                aria-checked={style === option.value}
                aria-keyshortcuts={optionShortcut?.aria}
                onClick={() => {
                  // Keep ribbon focus on the disclosure via Popover restore.
                  // chain().focus() steals into the editor and breaks L2 loops.
                  editor.commands.setDocumentStrike(option.value);
                  close();
                }}
              >
                <DocumentStrikeGlyph style={option.value} />
                <span>{option.label}</span>
                {optionShortcut && <kbd>{optionShortcut.label}</kbd>}
              </button>
            );
          })
        }
      </Popover>
    </span>
  );
}

function DocumentStrikeGlyph({
  style,
}: {
  style: WorkDocumentStrikeStyle;
}): ReactNode {
  return (
    <span
      className="work-document-strike-glyph"
      data-strike-style={style}
      aria-hidden="true"
    >
      <span>Aa</span>
    </span>
  );
}
