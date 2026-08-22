import type { Editor } from '@tiptap/core';
import { ChevronDown, Strikethrough } from 'lucide-react';
import { type ReactNode, useCallback, useSyncExternalStore } from 'react';
import { Popover } from '../../../design-system/primitives';
import {
  documentStrikeStyle,
  type WorkDocumentStrikeStyle,
} from '../work-document-strike';
import { moveOfficeMenuFocus } from './office-menu-keyboard';

const strikeOptions = [
  { value: 'none', label: '无删除线' },
  { value: 'single', label: '单删除线' },
  { value: 'double', label: '双删除线' },
] as const satisfies readonly {
  value: WorkDocumentStrikeStyle;
  label: string;
}[];

export function DocumentStrikeRibbon({
  editor,
  label = '删除线',
  menuLabel = '删除线样式',
  className = '',
}: {
  editor: Editor;
  label?: string;
  menuLabel?: string;
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
        label={`更多${label}`}
        panelLabel={menuLabel}
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
              aria-label={label}
              aria-pressed={active}
              title={`${label}（${currentLabel}）`}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => editor.chain().focus().toggleStrike().run()}
            >
              <Strikethrough size={16} aria-hidden="true" />
            </button>
            <button
              {...triggerProps}
              type="button"
              className={`work-document-strike-disclosure${open ? ' active' : ''}`}
              title={`更多${label}`}
              onMouseDown={(event) => event.preventDefault()}
            >
              <ChevronDown size={11} aria-hidden="true" />
            </button>
          </>
        )}
      >
        {(close) =>
          strikeOptions.map((option) => (
            <button
              key={option.value}
              type="button"
              role="menuitemradio"
              tabIndex={-1}
              aria-label={option.label}
              aria-checked={style === option.value}
              onClick={() => {
                close();
                editor.chain().focus().setDocumentStrike(option.value).run();
              }}
            >
              <DocumentStrikeGlyph style={option.value} />
              <span>{option.label}</span>
            </button>
          ))
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
