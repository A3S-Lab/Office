import type { Editor } from '@tiptap/core';
import { ChevronDown, Underline } from 'lucide-react';
import { type ReactNode, useCallback, useSyncExternalStore } from 'react';
import { Popover } from '../../../design-system/primitives';
import {
  documentUnderlineColor,
  documentUnderlineStyle,
  type WorkDocumentUnderlineStyle,
} from '../work-document-underline';
import { getDocumentCommandDefinition } from './document-command-catalog';
import { OfficeColorPicker } from './office-controls';
import { moveOfficeMenuFocus } from './office-menu-keyboard';

const underlineOptions = [
  { value: 'none', label: '无下划线' },
  { value: 'single', label: '单下划线', command: 'underline' },
  { value: 'words', label: '仅字下划线', command: 'wordsUnderline' },
  { value: 'double', label: '双下划线', command: 'doubleUnderline' },
  { value: 'thick', label: '粗下划线' },
  { value: 'dotted', label: '点线' },
  { value: 'dottedHeavy', label: '粗点线' },
  { value: 'dash', label: '短划线' },
  { value: 'dashedHeavy', label: '粗短划线' },
  { value: 'dashLong', label: '长划线' },
  { value: 'dashLongHeavy', label: '粗长划线' },
  { value: 'dotDash', label: '点划线' },
  { value: 'dashDotHeavy', label: '粗点划线' },
  { value: 'dotDotDash', label: '双点划线' },
  { value: 'dashDotDotHeavy', label: '粗双点划线' },
  { value: 'wave', label: '波浪线' },
  { value: 'wavyHeavy', label: '粗波浪线' },
  { value: 'wavyDouble', label: '双波浪线' },
] as const satisfies readonly {
  value: WorkDocumentUnderlineStyle;
  label: string;
  command?: 'doubleUnderline' | 'underline' | 'wordsUnderline';
}[];

export function DocumentUnderlineRibbon({
  editor,
  label = '下划线',
  menuLabel = '下划线样式',
  colorLabel = '下划线颜色',
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
        label={`更多${label}`}
        panelLabel={menuLabel}
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
              aria-label={label}
              aria-keyshortcuts={shortcut?.aria}
              aria-pressed={active}
              title={`${label}（${currentLabel}；${shortcut?.label ?? ''}）`}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => editor.chain().focus().toggleUnderline().run()}
            >
              <Underline size={16} aria-hidden="true" />
            </button>
            <button
              {...triggerProps}
              type="button"
              className={`work-document-underline-disclosure${open ? ' active' : ''}`}
              title={`更多${label}`}
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
                aria-checked={style === option.value}
                aria-keyshortcuts={optionShortcut?.aria}
                onClick={() => {
                  close();
                  editor
                    .chain()
                    .focus()
                    .setDocumentUnderline(option.value)
                    .run();
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
          ariaLabel={colorLabel}
          value={color}
          resetAction={{
            kind: 'automatic',
            label: '自动颜色',
            onSelect: () =>
              editor.chain().focus().setDocumentUnderlineColor(null).run(),
          }}
          onValueChange={(value) =>
            editor.chain().focus().setDocumentUnderlineColor(value).run()
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
