import { ChevronDown } from 'lucide-react';
import { useState } from 'react';
import { Popover } from '../../../design-system/primitives';
import type {
  SpreadsheetCellBorderFormat,
  SpreadsheetCellBorderStyle,
  SpreadsheetCellBorderTarget,
} from './spreadsheet-cell-border';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { moveOfficeMenuFocus } from './office-menu-keyboard';

const spreadsheetBorderStyleOptions: readonly {
  value: SpreadsheetCellBorderStyle;
  label: string;
}[] = [
  { value: 'thin', label: '细实线' },
  { value: 'dotted', label: '点线' },
  { value: 'dashed', label: '虚线' },
  { value: 'dash-dot', label: '点划线' },
  { value: 'dash-dot-dot', label: '双点划线' },
  { value: 'medium', label: '中等实线' },
  { value: 'medium-dashed', label: '中等虚线' },
  { value: 'medium-dash-dot', label: '中等点划线' },
  { value: 'medium-dash-dot-dot', label: '中等双点划线' },
  { value: 'thick', label: '粗实线' },
];

const spreadsheetBorderTargetOptions: readonly {
  target: SpreadsheetCellBorderTarget;
  definition:
    | (typeof spreadsheetCommandCatalog)['borderTop']
    | (typeof spreadsheetCommandCatalog)['borderBottom']
    | (typeof spreadsheetCommandCatalog)['borderLeft']
    | (typeof spreadsheetCommandCatalog)['borderRight']
    | (typeof spreadsheetCommandCatalog)['borderNone']
    | (typeof spreadsheetCommandCatalog)['borderAll']
    | (typeof spreadsheetCommandCatalog)['borderOutside']
    | (typeof spreadsheetCommandCatalog)['borderInside']
    | (typeof spreadsheetCommandCatalog)['borderHorizontal']
    | (typeof spreadsheetCommandCatalog)['borderVertical']
    | (typeof spreadsheetCommandCatalog)['borderDiagonal'];
}[] = [
  { target: 'top', definition: spreadsheetCommandCatalog.borderTop },
  { target: 'bottom', definition: spreadsheetCommandCatalog.borderBottom },
  { target: 'left', definition: spreadsheetCommandCatalog.borderLeft },
  { target: 'right', definition: spreadsheetCommandCatalog.borderRight },
  { target: 'none', definition: spreadsheetCommandCatalog.borderNone },
  { target: 'all', definition: spreadsheetCommandCatalog.borderAll },
  { target: 'outside', definition: spreadsheetCommandCatalog.borderOutside },
  { target: 'inside', definition: spreadsheetCommandCatalog.borderInside },
  {
    target: 'horizontal',
    definition: spreadsheetCommandCatalog.borderHorizontal,
  },
  {
    target: 'vertical',
    definition: spreadsheetCommandCatalog.borderVertical,
  },
  {
    target: 'diagonal',
    definition: spreadsheetCommandCatalog.borderDiagonal,
  },
];

export function SpreadsheetBorderRibbon({
  can,
  commands,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
}) {
  const [target, setTarget] = useState<SpreadsheetCellBorderTarget>('all');
  const [style, setStyle] = useState<SpreadsheetCellBorderStyle>('thin');
  const [color, setColor] = useState('#000000');
  const format = { target, style, color } satisfies SpreadsheetCellBorderFormat;
  const currentDefinition = spreadsheetBorderTargetOptions.find(
    (option) => option.target === target,
  )?.definition;
  const currentShortcut =
    currentDefinition && 'shortcut' in currentDefinition
      ? currentDefinition.shortcut
      : undefined;
  const currentLabel = currentDefinition?.label ?? '所有框线';
  const styleLabel = spreadsheetBorderStyleOptions.find(
    (option) => option.value === style,
  )?.label;
  const menuDisabled = spreadsheetBorderTargetOptions.every(
    (option) =>
      !can.setSelectedCellBorders({ ...format, target: option.target }),
  );

  return (
    <Popover
      label="更多框线"
      panelLabel="框线设置"
      panelRole="dialog"
      portal
      placement="bottom-end"
      className="work-spreadsheet-border-split-root"
      panelClassName="work-spreadsheet-border-panel"
      disabled={menuDisabled}
      focusFirstOnOpen
      trigger={(triggerProps, { open }) => (
        <>
          <button
            type="button"
            className="work-spreadsheet-border-primary"
            aria-label={currentLabel}
            aria-keyshortcuts={currentShortcut?.aria}
            title={`${currentLabel}（${styleLabel}，${color.toUpperCase()}${
              currentShortcut ? `；${currentShortcut.label}` : ''
            }）`}
            disabled={!can.setSelectedCellBorders(format)}
            onClick={() => commands.setSelectedCellBorders(format)}
          >
            <SpreadsheetBorderGlyph target={target} />
          </button>
          <button
            {...triggerProps}
            className={`work-spreadsheet-border-disclosure${open ? ' active' : ''}`}
            title="更多框线"
          >
            <ChevronDown size={12} aria-hidden="true" />
          </button>
        </>
      )}
    >
      {(close) => (
        <>
          <div className="work-spreadsheet-border-section-label">框线位置</div>
          <div
            className="work-spreadsheet-border-targets"
            role="menu"
            aria-label="框线位置"
            onKeyDown={moveOfficeMenuFocus}
          >
            {spreadsheetBorderTargetOptions.map(
              ({ target: option, definition }, index) => {
                const next = { ...format, target: option };
                const shortcut =
                  'shortcut' in definition ? definition.shortcut : undefined;
                return (
                  <button
                    key={definition.id}
                    type="button"
                    role="menuitemradio"
                    tabIndex={index === 0 ? 0 : -1}
                    aria-checked={option === target}
                    aria-label={definition.label}
                    aria-keyshortcuts={shortcut?.aria}
                    disabled={!can.setSelectedCellBorders(next)}
                    onClick={() => {
                      setTarget(option);
                      close();
                      commands.setSelectedCellBorders(next);
                    }}
                  >
                    <SpreadsheetBorderGlyph target={option} />
                    <span className="work-spreadsheet-border-target-label">
                      {definition.label}
                    </span>
                    {shortcut && <kbd>{shortcut.label}</kbd>}
                  </button>
                );
              },
            )}
          </div>
          <div className="work-spreadsheet-border-settings">
            <label>
              <span>线型</span>
              <select
                aria-label="框线样式"
                value={style}
                onChange={(event) =>
                  setStyle(event.target.value as SpreadsheetCellBorderStyle)
                }
              >
                {spreadsheetBorderStyleOptions.map((option) => (
                  <option key={option.value} value={option.value}>
                    {option.label}
                  </option>
                ))}
              </select>
            </label>
            <label>
              <span>颜色</span>
              <span className="work-spreadsheet-border-color-control">
                <span
                  className="work-spreadsheet-border-color-swatch"
                  style={{ backgroundColor: color }}
                  aria-hidden="true"
                />
                <input
                  type="color"
                  aria-label="框线颜色"
                  value={color}
                  onChange={(event) => setColor(event.target.value)}
                />
                <output>{color.toUpperCase()}</output>
              </span>
            </label>
          </div>
        </>
      )}
    </Popover>
  );
}

function SpreadsheetBorderGlyph({
  target,
}: {
  target: SpreadsheetCellBorderTarget;
}) {
  return (
    <span
      className="work-spreadsheet-border-glyph"
      data-border-target={target}
      aria-hidden="true"
    />
  );
}
