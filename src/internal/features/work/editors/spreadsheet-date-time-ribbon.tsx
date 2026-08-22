import { CalendarClock, CalendarDays, Clock } from 'lucide-react';
import { Popover } from '../../../design-system/primitives';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import type { SpreadsheetDateTimeKind } from './spreadsheet-date-time-command';

type SpreadsheetDateTimeCanCommands = Pick<
  SpreadsheetEditorCanCommands,
  'insertCurrentDateTime'
>;
type SpreadsheetDateTimeCommands = Pick<
  SpreadsheetEditorCommands,
  'insertCurrentDateTime'
>;

export function SpreadsheetDateTimeMenu({
  can,
  commands,
}: {
  can: SpreadsheetDateTimeCanCommands;
  commands: SpreadsheetDateTimeCommands;
}) {
  const items = [
    {
      definition: spreadsheetCommandCatalog.insertCurrentDate,
      icon: <CalendarDays size={16} />,
      kind: 'date',
    },
    {
      definition: spreadsheetCommandCatalog.insertCurrentTime,
      icon: <Clock size={16} />,
      kind: 'time',
    },
  ] as const satisfies readonly {
    definition: (typeof spreadsheetCommandCatalog)[
      | 'insertCurrentDate'
      | 'insertCurrentTime'];
    icon: React.ReactNode;
    kind: SpreadsheetDateTimeKind;
  }[];
  const disabled = items.every(({ kind }) => !can.insertCurrentDateTime(kind));

  return (
    <Popover
      label="日期和时间"
      panelLabel="插入日期和时间"
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
          className={`with-label work-spreadsheet-ribbon-menu-trigger work-spreadsheet-date-time-trigger${open ? ' active' : ''}`}
          title="插入当前日期或时间（Ctrl+; / Ctrl+Shift+;）"
        >
          <CalendarClock size={19} />
          <span>日期和时间</span>
        </button>
      )}
    >
      {(close) =>
        items.map(({ definition, icon, kind }) => (
          <button
            key={definition.id}
            type="button"
            role="menuitem"
            tabIndex={-1}
            aria-label={definition.label}
            aria-keyshortcuts={definition.shortcut.aria}
            disabled={!can.insertCurrentDateTime(kind)}
            onClick={() => {
              close();
              commands.insertCurrentDateTime(kind);
            }}
          >
            <span className="work-spreadsheet-ribbon-menu-item-icon">
              {icon}
            </span>
            <span>{definition.label}</span>
            <kbd>{definition.shortcut.label}</kbd>
          </button>
        ))
      }
    </Popover>
  );
}
