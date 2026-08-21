import type { Selection } from '@fortune-sheet/core';
import { PanelLeft, PanelsTopLeft, PanelTop, X } from 'lucide-react';
import type { ReactNode } from 'react';
import { Popover } from '../../../design-system/primitives';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import {
  type SpreadsheetFreezePanePreset,
  spreadsheetFreezePanesSelectionLabel,
} from './spreadsheet-freeze-panes';

export function SpreadsheetFreezePanesMenu({
  active,
  can,
  commands,
  selection,
}: {
  active: boolean;
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
  selection: Selection;
}) {
  const presets: readonly {
    preset: SpreadsheetFreezePanePreset;
    label: string;
    icon: ReactNode;
  }[] = [
    ...(active
      ? [
          {
            preset: 'none' as const,
            label: '取消冻结窗格',
            icon: <X size={16} />,
          },
        ]
      : []),
    {
      preset: 'selection',
      label: spreadsheetFreezePanesSelectionLabel(selection),
      icon: <PanelsTopLeft size={16} />,
    },
    {
      preset: 'topRow',
      label: '冻结首行',
      icon: <PanelTop size={16} />,
    },
    {
      preset: 'firstColumn',
      label: '冻结首列',
      icon: <PanelLeft size={16} />,
    },
  ];
  const disabled = !presets.some(({ preset }) => can.setFreezePanes(preset));

  return (
    <Popover
      label={spreadsheetCommandCatalog.freezePanes.label}
      panelLabel="冻结窗格选项"
      panelRole="menu"
      portal
      className="work-spreadsheet-ribbon-menu-root"
      panelClassName="work-office-context-menu work-spreadsheet-ribbon-menu"
      disabled={disabled}
      focusFirstOnOpen
      onPanelKeyDown={moveOfficeMenuFocus}
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          className={`with-label work-spreadsheet-ribbon-menu-trigger${active || open ? ' active' : ''}`}
          aria-pressed={active}
          title={active ? '冻结窗格（已启用）' : '冻结窗格'}
        >
          <PanelsTopLeft size={19} />
          <span>{spreadsheetCommandCatalog.freezePanes.label}</span>
        </button>
      )}
    >
      {(close) =>
        presets.map(({ preset, label, icon }) => (
          <button
            key={preset}
            type="button"
            role="menuitem"
            tabIndex={-1}
            disabled={!can.setFreezePanes(preset)}
            onClick={() => {
              close();
              commands.setFreezePanes(preset);
            }}
          >
            <span className="work-spreadsheet-ribbon-menu-item-icon">
              {icon}
            </span>
            <span>{label}</span>
          </button>
        ))
      }
    </Popover>
  );
}
