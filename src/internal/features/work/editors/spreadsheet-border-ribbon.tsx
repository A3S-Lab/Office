import { ChevronDown } from 'lucide-react';
import { useRef, useState } from 'react';
import { Popover } from '../../../design-system/primitives';
import { OfficeColorPicker, OfficeSelect } from './office-controls';
import { moveOfficeGridMenuFocus } from './office-menu-keyboard';
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
    | (typeof spreadsheetCommandCatalog)['borderDiagonalDown']
    | (typeof spreadsheetCommandCatalog)['borderDiagonalUp'];
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
    target: 'diagonalDown',
    definition: spreadsheetCommandCatalog.borderDiagonalDown,
  },
  {
    target: 'diagonalUp',
    definition: spreadsheetCommandCatalog.borderDiagonalUp,
  },
];

export function SpreadsheetBorderRibbon({
  can,
  commands,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
}) {
  const [open, setOpen] = useState(false);
  const [target, setTarget] = useState<SpreadsheetCellBorderTarget>('all');
  const [style, setStyle] = useState<SpreadsheetCellBorderStyle>('thin');
  const [color, setColor] = useState('#000000');
  const [baseline, setBaseline] = useState({
    target: 'all' as SpreadsheetCellBorderTarget,
    style: 'thin' as SpreadsheetCellBorderStyle,
    color: '#000000',
  });
  const appliedRef = useRef(false);
  const format = { target, style, color } satisfies SpreadsheetCellBorderFormat;
  const dirty =
    target !== baseline.target ||
    style !== baseline.style ||
    color !== baseline.color;
  const focusIndex = Math.max(
    0,
    spreadsheetBorderTargetOptions.findIndex(
      (option) => option.target === target,
    ),
  );
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

  const restoreBaseline = () => {
    setTarget(baseline.target);
    setStyle(baseline.style);
    setColor(baseline.color);
  };

  return (
    <Popover
      label="更多框线"
      panelLabel="框线设置"
      panelRole="dialog"
      portal
      placement="bottom-end"
      open={open}
      onOpenChange={(nextOpen) => {
        setOpen(nextOpen);
        if (nextOpen) {
          appliedRef.current = false;
          setBaseline({ target, style, color });
          return;
        }
        if (!appliedRef.current) {
          setTarget(baseline.target);
          setStyle(baseline.style);
          setColor(baseline.color);
        }
      }}
      className="work-spreadsheet-border-split-root"
      panelClassName="work-spreadsheet-border-panel"
      disabled={menuDisabled}
      focusFirstOnOpen
      trigger={(triggerProps, { open: popoverOpen }) => (
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
            className={`work-spreadsheet-border-disclosure${popoverOpen ? ' active' : ''}`}
            title="更多框线"
          >
            <ChevronDown size={12} aria-hidden="true" />
          </button>
        </>
      )}
    >
      {(close) => (
        <fieldset
          data-office-escape-consumer={dirty || undefined}
          onKeyDown={(event) => {
            if (event.key !== 'Escape' || !dirty) return;
            event.preventDefault();
            event.stopPropagation();
            restoreBaseline();
          }}
        >
          <legend className="sr-only">更多框线</legend>
          <div className="work-spreadsheet-border-section-label">框线位置</div>
          <div
            className="work-spreadsheet-border-targets"
            role="radiogroup"
            aria-label="框线位置"
            onKeyDown={(event) => moveOfficeGridMenuFocus(event, 2)}
          >
            {spreadsheetBorderTargetOptions.map(
              ({ target: option, definition }, index) => {
                const next = { ...format, target: option };
                const shortcut =
                  'shortcut' in definition ? definition.shortcut : undefined;
                return (
                  // biome-ignore lint/a11y/useSemanticElements: styled border radios; native input cannot host glyph and shortcut kbd.
                  <button
                    key={definition.id}
                    type="button"
                    role="radio"
                    tabIndex={index === focusIndex ? 0 : -1}
                    aria-checked={option === target}
                    aria-label={definition.label}
                    aria-keyshortcuts={shortcut?.aria}
                    disabled={!can.setSelectedCellBorders(next)}
                    onClick={() => {
                      appliedRef.current = true;
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
            <div className="work-office-field">
              <span>线型</span>
              <OfficeSelect<SpreadsheetCellBorderStyle>
                ariaLabel="框线样式"
                value={style}
                options={spreadsheetBorderStyleOptions}
                onValueChange={setStyle}
              />
            </div>
            <div className="work-office-field">
              <span>颜色</span>
              <OfficeColorPicker
                ariaLabel="框线颜色"
                value={color}
                onValueChange={setColor}
              />
            </div>
          </div>
        </fieldset>
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
