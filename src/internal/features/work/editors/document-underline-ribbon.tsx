import type { Editor } from '@tiptap/core';
import { ChevronDown, Underline } from 'lucide-react';
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
  documentUnderlineColor,
  documentUnderlineStyle,
  type WorkDocumentUnderlineStyle,
} from '../work-document-underline';
import { getDocumentCommandDefinition } from './document-command-catalog';
import { documentCommandLabel } from './document-command-i18n';
import { OfficeColorPicker } from './office-controls';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import { useOfficeMessages } from './office-messages-context';

const underlineOptionDefs = [
  { value: 'none', messageKey: 'document.underline.none' },
  {
    value: 'single',
    messageKey: 'document.underline.single',
    command: 'underline',
  },
  {
    value: 'words',
    messageKey: 'document.underline.words',
    command: 'wordsUnderline',
  },
  {
    value: 'double',
    messageKey: 'document.underline.double',
    command: 'doubleUnderline',
  },
  { value: 'thick', messageKey: 'document.underline.thick' },
  { value: 'dotted', messageKey: 'document.underline.dotted' },
  { value: 'dottedHeavy', messageKey: 'document.underline.dottedHeavy' },
  { value: 'dash', messageKey: 'document.underline.dash' },
  { value: 'dashedHeavy', messageKey: 'document.underline.dashedHeavy' },
  { value: 'dashLong', messageKey: 'document.underline.dashLong' },
  { value: 'dashLongHeavy', messageKey: 'document.underline.dashLongHeavy' },
  { value: 'dotDash', messageKey: 'document.underline.dotDash' },
  { value: 'dashDotHeavy', messageKey: 'document.underline.dashDotHeavy' },
  { value: 'dotDotDash', messageKey: 'document.underline.dotDotDash' },
  {
    value: 'dashDotDotHeavy',
    messageKey: 'document.underline.dashDotDotHeavy',
  },
  { value: 'wave', messageKey: 'document.underline.wave' },
  { value: 'wavyHeavy', messageKey: 'document.underline.wavyHeavy' },
  { value: 'wavyDouble', messageKey: 'document.underline.wavyDouble' },
] as const satisfies readonly {
  value: WorkDocumentUnderlineStyle;
  messageKey: keyof OfficeMessageCatalog;
  command?: 'doubleUnderline' | 'underline' | 'wordsUnderline';
}[];

export function DocumentUnderlineRibbon({
  editor,
  label,
  menuLabel,
  colorLabel,
  showColor = true,
  className = '',
}: {
  editor: Editor;
  label?: string;
  menuLabel?: string;
  colorLabel?: string;
  showColor?: boolean;
  className?: string;
}) {
  const messages = useOfficeMessages();
  const resolvedLabel = label ?? documentCommandLabel('underline', messages);
  const resolvedMenuLabel = menuLabel ?? messages['document.underline.menu'];
  const resolvedColorLabel = colorLabel ?? messages['document.underline.color'];
  const underlineOptions = useMemo(
    () =>
      underlineOptionDefs.map((option) => ({
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
  const style = documentUnderlineStyle(editor);
  const active = style !== 'none';
  const color = documentUnderlineColor(editor) ?? '#172033';
  const currentLabel =
    underlineOptions.find((option) => option.value === style)?.label ??
    underlineOptions[0].label;
  const shortcut = getDocumentCommandDefinition('underline').shortcut;

  return (
    <span
      className={`work-document-underline-control${className ? ` ${className}` : ''}`}
    >
      <Popover
        label={officeMessage(messages, 'document.underline.more', {
          label: resolvedLabel,
        })}
        panelLabel={resolvedMenuLabel}
        panelRole="menu"
        portal
        placement="bottom-end"
        className="work-document-underline-split-root"
        panelClassName="work-office-context-menu work-document-underline-menu"
        focusFirstOnOpen
        onPanelKeyDown={moveOfficeMenuFocus}
        trigger={(triggerProps, { open }) => (
          <>
            <button
              type="button"
              className={`work-document-underline-primary${active ? ' active' : ''}`}
              aria-label={resolvedLabel}
              aria-keyshortcuts={shortcut?.aria}
              aria-pressed={active}
              title={officeMessage(messages, 'document.underline.title', {
                label: resolvedLabel,
                current: currentLabel,
                shortcut: shortcut?.label ?? '',
              })}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => {
                // Keep ribbon focus (e.g. disclosure) — chain().focus() steals
                // into the editor and breaks L2 formatting loops.
                editor.commands.toggleUnderline();
              }}
            >
              <Underline size={16} aria-hidden="true" />
            </button>
            <button
              {...triggerProps}
              type="button"
              className={`work-document-underline-disclosure${open ? ' active' : ''}`}
              title={officeMessage(messages, 'document.underline.more', {
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
          underlineOptions.map((option) => {
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
                  editor.commands.setDocumentUnderline(option.value);
                  close();
                }}
              >
                <DocumentUnderlineGlyph style={option.value} />
                <span>{option.label}</span>
                {optionShortcut && <kbd>{optionShortcut.label}</kbd>}
              </button>
            );
          })
        }
      </Popover>
      {showColor && (
        <OfficeColorPicker
          compact
          className="work-document-underline-color"
          ariaLabel={resolvedColorLabel}
          value={color}
          resetAction={{
            kind: 'automatic',
            label: messages['document.color.automatic'],
            onSelect: () => editor.commands.setDocumentUnderlineColor(null),
          }}
          onValueChange={(value) =>
            editor.commands.setDocumentUnderlineColor(value)
          }
        />
      )}
    </span>
  );
}

function DocumentUnderlineGlyph({
  style,
}: {
  style: WorkDocumentUnderlineStyle;
}): ReactNode {
  return (
    <span
      className="work-document-underline-glyph"
      data-underline-style={style}
      aria-hidden="true"
    >
      <span>Aa</span>
    </span>
  );
}
