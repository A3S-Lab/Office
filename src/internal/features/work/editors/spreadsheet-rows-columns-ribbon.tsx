import {
  BetweenHorizontalEnd,
  BetweenHorizontalStart,
  BetweenVerticalEnd,
  BetweenVerticalStart,
  Columns3,
  Eye,
  EyeOff,
  Rows3,
} from 'lucide-react';
import { Fragment, type ReactNode } from 'react';
import { Popover } from '../../../design-system/primitives';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';

interface RowsAndColumnsItem {
  id: string;
  label: string;
  icon: ReactNode;
  danger?: boolean;
  dividerBefore?: boolean;
  disabled: boolean;
  execute: () => boolean;
  shortcut?: {
    aria: string;
    label: string;
  };
}

export function SpreadsheetRowsAndColumnsMenu({
  can,
  commands,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
}) {
  const items: readonly RowsAndColumnsItem[] = [
    {
      id: spreadsheetCommandCatalog.insertRowsAbove.id,
      label: spreadsheetCommandCatalog.insertRowsAbove.label,
      icon: <BetweenHorizontalStart size={16} />,
      disabled: !can.insertSelectedStructure('row', 'before'),
      execute: () => commands.insertSelectedStructure('row', 'before'),
    },
    {
      id: spreadsheetCommandCatalog.insertRowsBelow.id,
      label: spreadsheetCommandCatalog.insertRowsBelow.label,
      icon: <BetweenHorizontalEnd size={16} />,
      disabled: !can.insertSelectedStructure('row', 'after'),
      execute: () => commands.insertSelectedStructure('row', 'after'),
    },
    {
      id: spreadsheetCommandCatalog.insertColumnsLeft.id,
      label: spreadsheetCommandCatalog.insertColumnsLeft.label,
      icon: <BetweenVerticalStart size={16} />,
      disabled: !can.insertSelectedStructure('column', 'before'),
      execute: () => commands.insertSelectedStructure('column', 'before'),
    },
    {
      id: spreadsheetCommandCatalog.insertColumnsRight.id,
      label: spreadsheetCommandCatalog.insertColumnsRight.label,
      icon: <BetweenVerticalEnd size={16} />,
      disabled: !can.insertSelectedStructure('column', 'after'),
      execute: () => commands.insertSelectedStructure('column', 'after'),
    },
    {
      id: spreadsheetCommandCatalog.deleteRows.id,
      label: spreadsheetCommandCatalog.deleteRows.label,
      icon: <Rows3 size={16} />,
      danger: true,
      dividerBefore: true,
      disabled: !can.deleteSelectedStructure('row'),
      execute: () => commands.deleteSelectedStructure('row'),
    },
    {
      id: spreadsheetCommandCatalog.deleteColumns.id,
      label: spreadsheetCommandCatalog.deleteColumns.label,
      icon: <Columns3 size={16} />,
      danger: true,
      disabled: !can.deleteSelectedStructure('column'),
      execute: () => commands.deleteSelectedStructure('column'),
    },
    {
      id: spreadsheetCommandCatalog.hideRows.id,
      label: spreadsheetCommandCatalog.hideRows.label,
      icon: <EyeOff size={16} />,
      dividerBefore: true,
      disabled: !can.setSelectedStructureHidden('row', true),
      execute: () => commands.setSelectedStructureHidden('row', true),
      shortcut: spreadsheetCommandCatalog.hideRows.shortcut,
    },
    {
      id: spreadsheetCommandCatalog.hideColumns.id,
      label: spreadsheetCommandCatalog.hideColumns.label,
      icon: <EyeOff size={16} />,
      disabled: !can.setSelectedStructureHidden('column', true),
      execute: () => commands.setSelectedStructureHidden('column', true),
      shortcut: spreadsheetCommandCatalog.hideColumns.shortcut,
    },
    {
      id: spreadsheetCommandCatalog.unhideRows.id,
      label: spreadsheetCommandCatalog.unhideRows.label,
      icon: <Eye size={16} />,
      disabled: !can.setSelectedStructureHidden('row', false),
      execute: () => commands.setSelectedStructureHidden('row', false),
      shortcut: spreadsheetCommandCatalog.unhideRows.shortcut,
    },
    {
      id: spreadsheetCommandCatalog.unhideColumns.id,
      label: spreadsheetCommandCatalog.unhideColumns.label,
      icon: <Eye size={16} />,
      disabled: !can.setSelectedStructureHidden('column', false),
      execute: () => commands.setSelectedStructureHidden('column', false),
      shortcut: spreadsheetCommandCatalog.unhideColumns.shortcut,
    },
  ];
  const disabled = items.every((item) => item.disabled);

  return (
    <Popover
      label="行和列"
      panelLabel="行和列选项"
      panelRole="menu"
      portal
      className="work-spreadsheet-ribbon-menu-root"
      panelClassName="work-office-context-menu work-spreadsheet-ribbon-menu work-spreadsheet-rows-columns-menu"
      disabled={disabled}
      focusFirstOnOpen
      onPanelKeyDown={moveOfficeMenuFocus}
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          className={`with-label work-spreadsheet-ribbon-menu-trigger${open ? ' active' : ''}`}
          title="行和列"
        >
          <Rows3 size={19} />
          <span>行和列</span>
        </button>
      )}
    >
      {(close) =>
        items.map((item) => (
          <Fragment key={item.id}>
            {item.dividerBefore && (
              <hr className="work-spreadsheet-ribbon-menu-separator" />
            )}
            <button
              type="button"
              role="menuitem"
              tabIndex={-1}
              aria-label={item.label}
              aria-keyshortcuts={item.shortcut?.aria}
              className={item.danger ? 'danger' : undefined}
              disabled={item.disabled}
              onClick={() => {
                close();
                item.execute();
              }}
            >
              <span className="work-spreadsheet-ribbon-menu-item-icon">
                {item.icon}
              </span>
              <span>{item.label}</span>
              {item.shortcut && <kbd>{item.shortcut.label}</kbd>}
            </button>
          </Fragment>
        ))
      }
    </Popover>
  );
}
