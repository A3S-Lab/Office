import type { Cell } from '@fortune-sheet/core';
import { SwatchBook } from 'lucide-react';
import { Popover } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import { OfficeMenuGroup } from './office-menu-group';
import { moveOfficeGridMenuFocus } from './office-menu-keyboard';
import { useOfficeMessages } from './office-messages-context';
import type { SpreadsheetResolvedCellBorders } from './spreadsheet-cell-border';
import {
  spreadsheetCellStyleDefinitions,
  spreadsheetCellStyleGroupLabel,
  spreadsheetCellStyleGroups,
  spreadsheetCellStylePreset,
} from './spreadsheet-cell-style';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';

export function SpreadsheetCellStyleRibbon({
  can,
  commands,
  toolbarCell,
  toolbarCellBorders,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
  toolbarCell: Cell | null | undefined;
  toolbarCellBorders?: SpreadsheetResolvedCellBorders;
}) {
  const messages = useOfficeMessages();
  const definitions = spreadsheetCellStyleDefinitions(messages);
  const current = spreadsheetCellStylePreset(toolbarCell, toolbarCellBorders);
  const currentDefinition = definitions.find(({ id }) => id === current);
  const enabled = definitions.map(({ id }) => can.applyCellStyle(id));
  const firstEnabledIndex = enabled.findIndex(Boolean);
  const currentIndex = definitions.findIndex(({ id }) => id === current);
  const focusIndex = enabled[currentIndex] ? currentIndex : firstEnabledIndex;
  const definition = spreadsheetCommandCatalog.cellStyles;

  return (
    <Popover
      label={definition.label}
      panelLabel={officeMessage(messages, 'spreadsheet.cellStyle.panel')}
      panelRole="menu"
      portal
      placement="bottom-end"
      className="work-spreadsheet-cell-style-root"
      panelClassName="work-spreadsheet-cell-style-panel"
      disabled={firstEnabledIndex < 0}
      focusFirstOnOpen
      onPanelKeyDown={(event) =>
        moveOfficeGridMenuFocus(
          event,
          typeof window !== 'undefined' && window.innerWidth <= 520 ? 2 : 3,
        )
      }
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          className={`with-label work-spreadsheet-ribbon-menu-trigger work-spreadsheet-cell-style-trigger${open ? ' active' : ''}`}
          title={
            currentDefinition
              ? officeMessage(messages, 'spreadsheet.cellStyle.currentTitle', {
                  label: definition.label,
                  current: currentDefinition.label,
                })
              : definition.label
          }
        >
          <SwatchBook size={19} aria-hidden="true" />
          <span>{definition.label}</span>
        </button>
      )}
    >
      {(close) =>
        spreadsheetCellStyleGroups.map((group) => {
          const groupLabel = spreadsheetCellStyleGroupLabel(group, messages);
          return (
            <OfficeMenuGroup
              key={group}
              className="work-spreadsheet-cell-style-group"
              ariaLabel={groupLabel}
              data-office-menu-grid
            >
              <span className="work-spreadsheet-cell-style-group-label">
                {groupLabel}
              </span>
              <div
                className="work-spreadsheet-cell-style-grid"
                role="presentation"
              >
                {definitions.map((style, index) =>
                  style.group === group ? (
                    <button
                      key={style.id}
                      type="button"
                      role="menuitemradio"
                      aria-label={officeMessage(
                        messages,
                        'spreadsheet.cellStyle.applyAria',
                        { label: style.label },
                      )}
                      aria-checked={style.id === current}
                      tabIndex={index === focusIndex ? 0 : -1}
                      disabled={!enabled[index]}
                      title={style.description}
                      onClick={() => {
                        close();
                        commands.applyCellStyle(style.id);
                      }}
                    >
                      <span
                        className="work-spreadsheet-cell-style-preview"
                        style={style.preview}
                        aria-hidden="true"
                      >
                        {style.label}
                      </span>
                      <span className="work-spreadsheet-cell-style-description">
                        {style.description}
                      </span>
                    </button>
                  ) : null,
                )}
              </div>
            </OfficeMenuGroup>
          );
        })
      }
    </Popover>
  );
}
