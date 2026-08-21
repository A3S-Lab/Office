import type { Cell } from '@fortune-sheet/core';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  ChevronDown,
  Grid3X3,
  Merge,
  Rows3,
  WrapText,
  X,
} from 'lucide-react';
import type { ReactNode } from 'react';
import { Popover } from '../../../design-system/primitives';
import {
  spreadsheetTextOrientationChoiceFromCell,
  spreadsheetTextOrientationFromCell,
  type SpreadsheetTextOrientationId,
} from '../work-spreadsheet-text-orientation';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import type { SpreadsheetCellMergeCommand } from './spreadsheet-cell-merge';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

export function SpreadsheetAlignmentRibbonGroup({
  can,
  commands,
  toolbarCell,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
  toolbarCell: Cell | null | undefined;
}) {
  return (
    <WorkOfficeRibbonGroup label="对齐" priority="high">
      <WorkOfficeRibbonButton
        label="左对齐"
        displayLabel={false}
        active={String(toolbarCell?.ht ?? '1') === '1'}
        disabled={!can.setCellFormat('ht', '1')}
        onClick={() => commands.setCellFormat('ht', '1')}
      >
        <AlignLeft size={15} />
      </WorkOfficeRibbonButton>
      <WorkOfficeRibbonButton
        label="居中"
        displayLabel={false}
        active={String(toolbarCell?.ht) === '0'}
        disabled={!can.setCellFormat('ht', '0')}
        onClick={() => commands.setCellFormat('ht', '0')}
      >
        <AlignCenter size={15} />
      </WorkOfficeRibbonButton>
      <WorkOfficeRibbonButton
        label="右对齐"
        displayLabel={false}
        active={String(toolbarCell?.ht) === '2'}
        disabled={!can.setCellFormat('ht', '2')}
        onClick={() => commands.setCellFormat('ht', '2')}
      >
        <AlignRight size={15} />
      </WorkOfficeRibbonButton>
      <WorkOfficeRibbonButton
        label="顶端对齐"
        displayLabel={false}
        active={Number(toolbarCell?.vt ?? 1) === 1}
        disabled={!can.setCellFormat('vt', 1)}
        onClick={() => commands.setCellFormat('vt', 1)}
      >
        <AlignVerticalJustifyStart size={15} />
      </WorkOfficeRibbonButton>
      <WorkOfficeRibbonButton
        label="垂直居中"
        displayLabel={false}
        active={Number(toolbarCell?.vt) === 0}
        disabled={!can.setCellFormat('vt', 0)}
        onClick={() => commands.setCellFormat('vt', 0)}
      >
        <AlignVerticalJustifyCenter size={15} />
      </WorkOfficeRibbonButton>
      <WorkOfficeRibbonButton
        label="底端对齐"
        displayLabel={false}
        active={Number(toolbarCell?.vt) === 2}
        disabled={!can.setCellFormat('vt', 2)}
        onClick={() => commands.setCellFormat('vt', 2)}
      >
        <AlignVerticalJustifyEnd size={15} />
      </WorkOfficeRibbonButton>
      <WorkOfficeRibbonButton
        label="自动换行"
        displayLabel={false}
        active={String(toolbarCell?.tb) === '2'}
        disabled={
          !can.setCellFormat('tb', String(toolbarCell?.tb) === '2' ? '1' : '2')
        }
        onClick={() =>
          commands.setCellFormat(
            'tb',
            String(toolbarCell?.tb) === '2' ? '1' : '2',
          )
        }
      >
        <WrapText size={15} />
      </WorkOfficeRibbonButton>
      <SpreadsheetTextOrientationMenu
        can={can}
        commands={commands}
        toolbarCell={toolbarCell}
      />
      <SpreadsheetMergeMenu can={can} commands={commands} />
    </WorkOfficeRibbonGroup>
  );
}

const orientationItems: readonly {
  id: SpreadsheetTextOrientationId;
  definition:
    | typeof spreadsheetCommandCatalog.textOrientationHorizontal
    | typeof spreadsheetCommandCatalog.textOrientationAngleCounterclockwise
    | typeof spreadsheetCommandCatalog.textOrientationAngleClockwise
    | typeof spreadsheetCommandCatalog.textOrientationVertical
    | typeof spreadsheetCommandCatalog.textOrientationRotateUp
    | typeof spreadsheetCommandCatalog.textOrientationRotateDown;
}[] = [
  {
    id: 'horizontal',
    definition: spreadsheetCommandCatalog.textOrientationHorizontal,
  },
  {
    id: 'angleCounterclockwise',
    definition: spreadsheetCommandCatalog.textOrientationAngleCounterclockwise,
  },
  {
    id: 'angleClockwise',
    definition: spreadsheetCommandCatalog.textOrientationAngleClockwise,
  },
  {
    id: 'vertical',
    definition: spreadsheetCommandCatalog.textOrientationVertical,
  },
  {
    id: 'rotateUp',
    definition: spreadsheetCommandCatalog.textOrientationRotateUp,
  },
  {
    id: 'rotateDown',
    definition: spreadsheetCommandCatalog.textOrientationRotateDown,
  },
];

function SpreadsheetTextOrientationMenu({
  can,
  commands,
  toolbarCell,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
  toolbarCell: Cell | null | undefined;
}) {
  const selected = spreadsheetTextOrientationChoiceFromCell(toolbarCell);
  const orientation = spreadsheetTextOrientationFromCell(toolbarCell);
  const active = orientation.kind === 'stacked' || orientation.angle !== 0;
  const currentLabel =
    orientationItems.find(({ id }) => id === selected)?.definition.label ??
    (orientation.kind === 'rotation' ? `${orientation.angle}°` : '竖排文字');
  const disabled = orientationItems.every(
    ({ id }) => !can.setTextOrientation(id),
  );

  return (
    <Popover
      label="文字方向"
      panelLabel="文字方向选项"
      panelRole="menu"
      portal
      className="work-spreadsheet-ribbon-menu-root"
      panelClassName="work-office-context-menu work-spreadsheet-ribbon-menu work-spreadsheet-orientation-menu"
      disabled={disabled}
      focusFirstOnOpen
      onPanelKeyDown={moveOfficeMenuFocus}
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          className={`with-label work-spreadsheet-ribbon-menu-trigger work-spreadsheet-orientation-trigger${active || open ? ' active' : ''}`}
          aria-pressed={active}
          title={`文字方向（当前：${currentLabel}）`}
        >
          <TextOrientationGlyph orientation={selected ?? 'horizontal'} />
          <span>文字方向</span>
        </button>
      )}
    >
      {(close) =>
        orientationItems.map(({ id, definition }) => (
          <button
            key={definition.id}
            type="button"
            role="menuitemradio"
            tabIndex={-1}
            aria-checked={selected === id}
            disabled={!can.setTextOrientation(id)}
            onClick={() => {
              close();
              commands.setTextOrientation(id);
            }}
          >
            <span className="work-spreadsheet-ribbon-menu-item-icon">
              <TextOrientationGlyph orientation={id} size={16} />
            </span>
            <span>{definition.label}</span>
          </button>
        ))
      }
    </Popover>
  );
}

function TextOrientationGlyph({
  orientation,
  size = 19,
}: {
  orientation: SpreadsheetTextOrientationId;
  size?: number;
}) {
  if (orientation === 'vertical') {
    return (
      <svg
        className="work-spreadsheet-orientation-glyph"
        width={size}
        height={size}
        viewBox="0 0 20 20"
        fill="none"
        aria-hidden="true"
      >
        <path d="M7 3h6M8 7h4M7 11h6M8 15h4" />
        <path d="m5 4-2 2 2 2M15 12l2 2-2 2" />
      </svg>
    );
  }
  const rotation = {
    horizontal: 0,
    angleCounterclockwise: -45,
    angleClockwise: 45,
    rotateUp: -90,
    rotateDown: 90,
  }[orientation];
  return (
    <svg
      className="work-spreadsheet-orientation-glyph"
      width={size}
      height={size}
      viewBox="0 0 20 20"
      fill="none"
      aria-hidden="true"
    >
      <g transform={`rotate(${rotation} 10 10)`}>
        <path d="m5.5 14.5 3.7-9h1.6l3.7 9M7 11h6M4 16.5h12" />
      </g>
    </svg>
  );
}

function SpreadsheetMergeMenu({
  can,
  commands,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
}) {
  const items: readonly {
    command: SpreadsheetCellMergeCommand;
    id: string;
    label: string;
    icon: ReactNode;
  }[] = [
    {
      command: 'merge-and-center',
      id: spreadsheetCommandCatalog.mergeAndCenter.id,
      label: spreadsheetCommandCatalog.mergeAndCenter.label,
      icon: <Merge size={16} />,
    },
    {
      command: 'merge-cells',
      id: spreadsheetCommandCatalog.mergeCells.id,
      label: spreadsheetCommandCatalog.mergeCells.label,
      icon: <Grid3X3 size={16} />,
    },
    {
      command: 'merge-across',
      id: spreadsheetCommandCatalog.mergeAcross.id,
      label: spreadsheetCommandCatalog.mergeAcross.label,
      icon: <Rows3 size={16} />,
    },
    {
      command: 'unmerge-cells',
      id: spreadsheetCommandCatalog.unmergeCells.id,
      label: spreadsheetCommandCatalog.unmergeCells.label,
      icon: <X size={16} />,
    },
    {
      command: 'unmerge-and-fill',
      id: spreadsheetCommandCatalog.unmergeAndFill.id,
      label: spreadsheetCommandCatalog.unmergeAndFill.label,
      icon: <Grid3X3 size={16} />,
    },
  ];
  const primaryDisabled = !can.mergeSelectedCells('merge-and-center');
  const menuDisabled = items.every(
    ({ command }) => can.mergeSelectedCells(command) === false,
  );

  return (
    <Popover
      label="更多合并方式"
      panelLabel="合并选项"
      panelRole="menu"
      portal
      className="work-spreadsheet-ribbon-split-root"
      panelClassName="work-office-context-menu work-spreadsheet-ribbon-menu"
      disabled={menuDisabled}
      focusFirstOnOpen
      onPanelKeyDown={moveOfficeMenuFocus}
      trigger={(triggerProps, { open }) => (
        <>
          <button
            type="button"
            className="with-label work-spreadsheet-ribbon-split-primary"
            aria-label={spreadsheetCommandCatalog.mergeAndCenter.label}
            aria-keyshortcuts={
              spreadsheetCommandCatalog.mergeAndCenter.shortcut.aria
            }
            title={`${spreadsheetCommandCatalog.mergeAndCenter.label}（${spreadsheetCommandCatalog.mergeAndCenter.shortcut.label}）`}
            disabled={primaryDisabled}
            onClick={() => commands.mergeSelectedCells('merge-and-center')}
          >
            <Merge size={19} />
            <span>{spreadsheetCommandCatalog.mergeAndCenter.label}</span>
          </button>
          <button
            {...triggerProps}
            className={`work-spreadsheet-ribbon-split-disclosure${open ? ' active' : ''}`}
            title="更多合并方式"
          >
            <ChevronDown size={13} aria-hidden="true" />
          </button>
        </>
      )}
    >
      {(close) =>
        items.map(({ command, id, label, icon }) => (
          <button
            key={id}
            type="button"
            role="menuitem"
            tabIndex={-1}
            aria-keyshortcuts={
              command === 'merge-and-center'
                ? spreadsheetCommandCatalog.mergeAndCenter.shortcut.aria
                : undefined
            }
            disabled={!can.mergeSelectedCells(command)}
            onClick={() => {
              close();
              commands.mergeSelectedCells(command);
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
