import type { Selection } from '@fortune-sheet/core';
import { PanelLeft, PanelsTopLeft, PanelTop, X } from 'lucide-react';
import type { ReactNode } from 'react';
import { Popover } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import { useOfficeMessages } from './office-messages-context';
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
  const messages = useOfficeMessages();
  const presets: readonly {
    preset: SpreadsheetFreezePanePreset;
    label: string;
    icon: ReactNode;
  }[] = [
    ...(active
      ? [
          {
            preset: 'none' as const,
            label: officeMessage(messages, 'spreadsheet.freeze.unfreeze'),
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
      label: officeMessage(messages, 'spreadsheet.freeze.topRow'),
      icon: <PanelTop size={16} />,
    },
    {
      preset: 'firstColumn',
      label: officeMessage(messages, 'spreadsheet.freeze.firstColumn'),
      icon: <PanelLeft size={16} />,
    },
  ];
  const disabled = !presets.some(({ preset }) => can.setFreezePanes(preset));

  return (
    <Popover
      label={spreadsheetCommandCatalog.freezePanes.label}
      panelLabel={officeMessage(messages, 'spreadsheet.freeze.options')}
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
          title={
            active
              ? officeMessage(messages, 'spreadsheet.freeze.titleActive')
              : officeMessage(messages, 'spreadsheet.freeze.title')
          }
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
