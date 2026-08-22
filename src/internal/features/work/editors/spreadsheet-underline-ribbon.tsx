import { ChevronDown, Underline } from 'lucide-react';
import { Popover } from '../../../design-system/primitives';
import {
  type SpreadsheetUnderlineStyle,
  spreadsheetUnderlineCellValue,
  spreadsheetUnderlineStyle,
} from '../work-spreadsheet-underline';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';

const underlineOptions = [
  { value: 'none', label: '无下划线' },
  { value: 'single', label: '单下划线' },
  { value: 'double', label: '双下划线' },
  { value: 'singleAccounting', label: '单会计用下划线' },
  { value: 'doubleAccounting', label: '双会计用下划线' },
] as const satisfies readonly {
  label: string;
  value: SpreadsheetUnderlineStyle;
}[];

export function SpreadsheetUnderlineRibbon({
  can,
  commands,
  value,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
  value: unknown;
}) {
  const style = spreadsheetUnderlineStyle(value);
  const active = style !== 'none';
  const definition = spreadsheetCommandCatalog.underline;
  const currentLabel =
    underlineOptions.find((option) => option.value === style)?.label ??
    underlineOptions[0].label;
  const menuDisabled = underlineOptions.every(
    (option) =>
      !can.setCellFormat('un', spreadsheetUnderlineCellValue(option.value)),
  );

  return (
    <Popover
      label="更多下划线"
      panelLabel="下划线样式"
      panelRole="menu"
      portal
      placement="bottom-end"
      className="work-spreadsheet-underline-split-root"
      panelClassName="work-office-context-menu work-spreadsheet-ribbon-menu work-spreadsheet-underline-menu"
      disabled={menuDisabled}
      focusFirstOnOpen
      onPanelKeyDown={moveOfficeMenuFocus}
      trigger={(triggerProps, { open }) => (
        <>
          <button
            type="button"
            className={`work-spreadsheet-underline-primary${active ? ' active' : ''}`}
            aria-label={definition.label}
            aria-keyshortcuts={definition.shortcut.aria}
            aria-pressed={active}
            title={`${definition.label}（${currentLabel}；${definition.shortcut.label}）`}
            disabled={!can.toggleCellFormat('un')}
            onClick={() => commands.toggleCellFormat('un')}
          >
            <Underline size={15} aria-hidden="true" />
          </button>
          <button
            {...triggerProps}
            className={`work-spreadsheet-underline-disclosure${open ? ' active' : ''}`}
            title="更多下划线"
          >
            <ChevronDown size={12} aria-hidden="true" />
          </button>
        </>
      )}
    >
      {(close) =>
        underlineOptions.map((option) => {
          const cellValue = spreadsheetUnderlineCellValue(option.value);
          return (
            <button
              key={option.value}
              type="button"
              role="menuitemradio"
              tabIndex={-1}
              aria-checked={style === option.value}
              aria-label={option.label}
              disabled={!can.setCellFormat('un', cellValue)}
              onClick={() => {
                close();
                commands.setCellFormat('un', cellValue);
              }}
            >
              <SpreadsheetUnderlineGlyph style={option.value} />
              <span>{option.label}</span>
            </button>
          );
        })
      }
    </Popover>
  );
}

function SpreadsheetUnderlineGlyph({
  style,
}: {
  style: SpreadsheetUnderlineStyle;
}) {
  return (
    <span
      className="work-spreadsheet-underline-glyph"
      data-underline-style={style}
      aria-hidden="true"
    >
      <Underline size={16} />
    </span>
  );
}
