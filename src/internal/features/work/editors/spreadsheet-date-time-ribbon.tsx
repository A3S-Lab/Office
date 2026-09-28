import { CalendarClock, CalendarDays, Clock } from 'lucide-react';
import { Popover } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import { useOfficeMessages } from './office-messages-context';
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
  const messages = useOfficeMessages();
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
      label={officeMessage(messages, 'spreadsheet.dateTime.label')}
      panelLabel={officeMessage(messages, 'spreadsheet.dateTime.options')}
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
          title={officeMessage(messages, 'spreadsheet.dateTime.title')}
        >
          <CalendarClock size={19} />
          <span>{officeMessage(messages, 'spreadsheet.dateTime.label')}</span>
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
