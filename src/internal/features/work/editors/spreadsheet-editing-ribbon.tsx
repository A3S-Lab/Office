import {
  ArrowDown,
  ArrowDownToLine,
  ArrowLeftToLine,
  ArrowRightToLine,
  ArrowUp,
  ArrowUpToLine,
  Calculator,
  ChevronDown,
  Eraser,
  FileX2,
  Hash,
  Link2Off,
  LocateFixed,
  MessageSquareX,
  Paintbrush,
  Search,
  Sigma,
  Trash2,
} from 'lucide-react';
import type { ReactNode } from 'react';
import { Popover } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import { useOfficeMessages } from './office-messages-context';
import type { SpreadsheetAutoSumFunction } from './spreadsheet-auto-sum';
import type { SpreadsheetCellClearMode } from './spreadsheet-cell-clear';
import type { SpreadsheetCellFillDirection } from './spreadsheet-cell-fill';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { WorkOfficeRibbonGroup } from './work-office-chrome';

export function SpreadsheetEditingRibbonGroup({
  can,
  commands,
  findOpen,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
  findOpen: boolean;
}) {
  const messages = useOfficeMessages();
  return (
    <WorkOfficeRibbonGroup label={officeMessage(messages, 'spreadsheet.ribbon.editing')} priority="low">
      <SpreadsheetAutoSumMenu can={can} commands={commands} />
      <SpreadsheetFillMenu can={can} commands={commands} />
      <SpreadsheetClearMenu can={can} commands={commands} />
      <SpreadsheetFindAndSelectMenu
        can={can}
        commands={commands}
        findOpen={findOpen}
      />
    </WorkOfficeRibbonGroup>
  );
}

function SpreadsheetFindAndSelectMenu({
  can,
  commands,
  findOpen,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
  findOpen: boolean;
}) {
  const messages = useOfficeMessages();
  const findDisabled = !can.openFind();
  const goToDisabled = !can.openGoTo();
  return (
    <Popover
      label={spreadsheetCommandCatalog.findAndSelect.label}
      panelLabel={officeMessage(messages, 'spreadsheet.ribbon.findSelectPanel')}
      panelRole="menu"
      portal
      className="work-spreadsheet-ribbon-menu-root"
      panelClassName="work-office-context-menu work-spreadsheet-ribbon-menu work-spreadsheet-find-select-menu"
      disabled={findDisabled && goToDisabled}
      focusFirstOnOpen
      onPanelKeyDown={moveOfficeMenuFocus}
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          className={`with-label work-spreadsheet-ribbon-menu-trigger${findOpen || open ? ' active' : ''}`}
          title={
            findOpen
              ? officeMessage(messages, 'spreadsheet.ribbon.findOpenTitle', { label: spreadsheetCommandCatalog.findAndSelect.label })
              : spreadsheetCommandCatalog.findAndSelect.label
          }
        >
          <Search size={19} />
          <span>{spreadsheetCommandCatalog.findAndSelect.label}</span>
        </button>
      )}
    >
      {(close) => (
        <>
          <button
            type="button"
            role="menuitem"
            tabIndex={-1}
            aria-label={spreadsheetCommandCatalog.find.label}
            aria-keyshortcuts={spreadsheetCommandCatalog.find.shortcut.aria}
            disabled={findDisabled}
            onClick={() => {
              close();
              commands.openFind();
            }}
          >
            <span className="work-spreadsheet-ribbon-menu-item-icon">
              <Search size={16} />
            </span>
            <span>{spreadsheetCommandCatalog.find.label}</span>
            <kbd>{spreadsheetCommandCatalog.find.shortcut.label}</kbd>
          </button>
          <button
            type="button"
            role="menuitem"
            tabIndex={-1}
            aria-label={spreadsheetCommandCatalog.goTo.label}
            aria-keyshortcuts={spreadsheetCommandCatalog.goTo.shortcut.aria}
            disabled={goToDisabled}
            onClick={() => {
              close();
              commands.openGoTo();
            }}
          >
            <span className="work-spreadsheet-ribbon-menu-item-icon">
              <LocateFixed size={16} />
            </span>
            <span>{spreadsheetCommandCatalog.goTo.label}</span>
            <kbd>{spreadsheetCommandCatalog.goTo.shortcut.label}</kbd>
          </button>
        </>
      )}
    </Popover>
  );
}

function SpreadsheetClearMenu({
  can,
  commands,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
}) {
  const messages = useOfficeMessages();
  const items: readonly {
    mode: SpreadsheetCellClearMode;
    id: string;
    label: string;
    icon: ReactNode;
  }[] = [
    {
      mode: 'all',
      id: spreadsheetCommandCatalog.clearAll.id,
      label: spreadsheetCommandCatalog.clearAll.label,
      icon: <Trash2 size={16} />,
    },
    {
      mode: 'formats',
      id: spreadsheetCommandCatalog.clearFormats.id,
      label: spreadsheetCommandCatalog.clearFormats.label,
      icon: <Paintbrush size={16} />,
    },
    {
      mode: 'contents',
      id: spreadsheetCommandCatalog.clearContents.id,
      label: spreadsheetCommandCatalog.clearContents.label,
      icon: <FileX2 size={16} />,
    },
    {
      mode: 'comments',
      id: spreadsheetCommandCatalog.clearComments.id,
      label: spreadsheetCommandCatalog.clearComments.label,
      icon: <MessageSquareX size={16} />,
    },
    {
      mode: 'hyperlinks',
      id: spreadsheetCommandCatalog.clearHyperlinks.id,
      label: spreadsheetCommandCatalog.clearHyperlinks.label,
      icon: <Link2Off size={16} />,
    },
  ];
  const disabled = items.every(({ mode }) => !can.clearSelectedCells(mode));

  return (
    <Popover
      label={officeMessage(messages, 'spreadsheet.ribbon.clear')}
      panelLabel={officeMessage(messages, 'spreadsheet.ribbon.clearPanel')}
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
          className={`with-label work-spreadsheet-ribbon-menu-trigger${open ? ' active' : ''}`}
          title={officeMessage(messages, 'spreadsheet.ribbon.clear')}
        >
          <Eraser size={19} />
          <span>{officeMessage(messages, 'spreadsheet.ribbon.clear')}</span>
        </button>
      )}
    >
      {(close) =>
        items.map(({ mode, id, label, icon }) => (
          <button
            key={id}
            type="button"
            role="menuitem"
            tabIndex={-1}
            aria-keyshortcuts={
              mode === 'contents'
                ? spreadsheetCommandCatalog.clearContents.shortcut.aria
                : undefined
            }
            disabled={!can.clearSelectedCells(mode)}
            onClick={() => {
              close();
              commands.clearSelectedCells(mode);
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

function SpreadsheetAutoSumMenu({
  can,
  commands,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
}) {
  const messages = useOfficeMessages();
  const items: readonly {
    functionName: SpreadsheetAutoSumFunction;
    id: string;
    label: string;
    icon: ReactNode;
    shortcut?: string;
  }[] = [
    {
      functionName: 'sum',
      id: spreadsheetCommandCatalog.autoSum.id,
      label: spreadsheetCommandCatalog.autoSum.label,
      icon: <Sigma size={16} />,
      shortcut: spreadsheetCommandCatalog.autoSum.shortcut.aria,
    },
    {
      functionName: 'average',
      id: spreadsheetCommandCatalog.autoAverage.id,
      label: spreadsheetCommandCatalog.autoAverage.label,
      icon: <Calculator size={16} />,
    },
    {
      functionName: 'count',
      id: spreadsheetCommandCatalog.autoCount.id,
      label: spreadsheetCommandCatalog.autoCount.label,
      icon: <Hash size={16} />,
    },
    {
      functionName: 'max',
      id: spreadsheetCommandCatalog.autoMaximum.id,
      label: spreadsheetCommandCatalog.autoMaximum.label,
      icon: <ArrowUp size={16} />,
    },
    {
      functionName: 'min',
      id: spreadsheetCommandCatalog.autoMinimum.id,
      label: spreadsheetCommandCatalog.autoMinimum.label,
      icon: <ArrowDown size={16} />,
    },
  ];
  const primaryDisabled = !can.applyAutoSum('sum');

  return (
    <Popover
      label={officeMessage(messages, 'spreadsheet.ribbon.moreAutoCalc')}
      panelLabel={officeMessage(messages, 'spreadsheet.ribbon.autoCalcPanel')}
      panelRole="menu"
      portal
      className="work-spreadsheet-ribbon-split-root"
      panelClassName="work-office-context-menu work-spreadsheet-ribbon-menu"
      disabled={primaryDisabled}
      focusFirstOnOpen
      onPanelKeyDown={moveOfficeMenuFocus}
      trigger={(triggerProps, { open }) => (
        <>
          <button
            type="button"
            className="with-label work-spreadsheet-ribbon-split-primary"
            aria-label={spreadsheetCommandCatalog.autoSum.label}
            aria-keyshortcuts={spreadsheetCommandCatalog.autoSum.shortcut.aria}
            title={`${spreadsheetCommandCatalog.autoSum.label}（${spreadsheetCommandCatalog.autoSum.shortcut.label}）`}
            disabled={primaryDisabled}
            onClick={() => commands.applyAutoSum('sum')}
          >
            <Sigma size={19} />
            <span>{spreadsheetCommandCatalog.autoSum.label}</span>
          </button>
          <button
            {...triggerProps}
            className={`work-spreadsheet-ribbon-split-disclosure${open ? ' active' : ''}`}
            title={officeMessage(messages, 'spreadsheet.ribbon.moreAutoCalc')}
          >
            <ChevronDown size={13} aria-hidden="true" />
          </button>
        </>
      )}
    >
      {(close) =>
        items.map(({ functionName, id, label, icon, shortcut }) => (
          <button
            key={id}
            type="button"
            role="menuitem"
            tabIndex={-1}
            aria-keyshortcuts={shortcut}
            disabled={primaryDisabled}
            onClick={() => {
              close();
              commands.applyAutoSum(functionName);
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

function SpreadsheetFillMenu({
  can,
  commands,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
}) {
  const messages = useOfficeMessages();
  const items: readonly {
    direction: SpreadsheetCellFillDirection;
    id: string;
    label: string;
    icon: ReactNode;
    shortcut?: string;
  }[] = [
    {
      direction: 'down',
      id: spreadsheetCommandCatalog.fillDown.id,
      label: spreadsheetCommandCatalog.fillDown.label,
      icon: <ArrowDownToLine size={16} />,
      shortcut: spreadsheetCommandCatalog.fillDown.shortcut.aria,
    },
    {
      direction: 'right',
      id: spreadsheetCommandCatalog.fillRight.id,
      label: spreadsheetCommandCatalog.fillRight.label,
      icon: <ArrowRightToLine size={16} />,
      shortcut: spreadsheetCommandCatalog.fillRight.shortcut.aria,
    },
    {
      direction: 'up',
      id: spreadsheetCommandCatalog.fillUp.id,
      label: spreadsheetCommandCatalog.fillUp.label,
      icon: <ArrowUpToLine size={16} />,
    },
    {
      direction: 'left',
      id: spreadsheetCommandCatalog.fillLeft.id,
      label: spreadsheetCommandCatalog.fillLeft.label,
      icon: <ArrowLeftToLine size={16} />,
    },
  ];
  const disabled = items.every(
    ({ direction }) => !can.fillSelectedCells(direction),
  );

  return (
    <Popover
      label={officeMessage(messages, 'spreadsheet.ribbon.fill')}
      panelLabel={officeMessage(messages, 'spreadsheet.ribbon.fillPanel')}
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
          className={`with-label work-spreadsheet-ribbon-menu-trigger${open ? ' active' : ''}`}
          title={officeMessage(messages, 'spreadsheet.ribbon.fill')}
        >
          <ArrowDownToLine size={19} />
          <span>{officeMessage(messages, 'spreadsheet.ribbon.fill')}</span>
        </button>
      )}
    >
      {(close) =>
        items.map(({ direction, id, label, icon, shortcut }) => (
          <button
            key={id}
            type="button"
            role="menuitem"
            tabIndex={-1}
            aria-keyshortcuts={shortcut}
            disabled={!can.fillSelectedCells(direction)}
            onClick={() => {
              close();
              commands.fillSelectedCells(direction);
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
