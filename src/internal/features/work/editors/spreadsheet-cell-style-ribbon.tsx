import type { Cell } from '@fortune-sheet/core';
import { SwatchBook } from 'lucide-react';
import { Popover } from '../../../design-system/primitives';
import type { SpreadsheetResolvedCellBorders } from './spreadsheet-cell-border';
import {
  spreadsheetCellStyleDefinitions,
  spreadsheetCellStylePreset,
  type SpreadsheetCellStyleDefinition,
} from './spreadsheet-cell-style';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { moveOfficeGridMenuFocus } from './office-menu-keyboard';

const spreadsheetCellStyleGroups = [
  '常用',
  '数据和模型',
  '标题和汇总',
] as const satisfies readonly SpreadsheetCellStyleDefinition['group'][];

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
  const current = spreadsheetCellStylePreset(toolbarCell, toolbarCellBorders);
  const currentDefinition = spreadsheetCellStyleDefinitions.find(
    ({ id }) => id === current,
  );
  const enabled = spreadsheetCellStyleDefinitions.map(({ id }) =>
    can.applyCellStyle(id),
  );
  const firstEnabledIndex = enabled.findIndex(Boolean);
  const currentIndex = spreadsheetCellStyleDefinitions.findIndex(
    ({ id }) => id === current,
  );
  const focusIndex = enabled[currentIndex] ? currentIndex : firstEnabledIndex;
  const definition = spreadsheetCommandCatalog.cellStyles;

  return (
    <Popover
      label={definition.label}
      panelLabel="单元格样式库"
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
              ? `${definition.label}（当前：${currentDefinition.label}）`
              : definition.label
          }
        >
          <SwatchBook size={19} aria-hidden="true" />
          <span>{definition.label}</span>
        </button>
      )}
    >
      {(close) =>
        spreadsheetCellStyleGroups.map((group) => (
          <fieldset
            key={group}
            className="work-spreadsheet-cell-style-group"
            aria-label={group}
            data-office-menu-grid
          >
            <legend className="work-spreadsheet-cell-style-group-label">
              {group}
            </legend>
            <div
              className="work-spreadsheet-cell-style-grid"
              role="presentation"
            >
              {spreadsheetCellStyleDefinitions.map((style, index) =>
                style.group === group ? (
                  <button
                    key={style.id}
                    type="button"
                    role="menuitemradio"
                    aria-label={`应用单元格样式：${style.label}`}
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
          </fieldset>
        ))
      }
    </Popover>
  );
}
