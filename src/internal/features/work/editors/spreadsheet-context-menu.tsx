import {
  ArrowDown,
  ArrowUp,
  ClipboardPaste,
  Copy,
  Eraser,
  Eye,
  EyeOff,
  MoveHorizontal,
  MoveVertical,
  Plus,
  Scissors,
  Trash2,
} from 'lucide-react';
import { showToast } from '../../../state/app-state';
import type { WorkspaceContextMenuItem } from '../../workspace/components/workspace-context-menu';
import type { WorkSpreadsheetAgentSelection } from '../work-spreadsheet-agent-context';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
  SpreadsheetStructureAxis,
} from './spreadsheet-command-controller';

type SpreadsheetContextCan = Pick<
  SpreadsheetEditorCanCommands,
  | 'clearSelectedCells'
  | 'copySelection'
  | 'cutSelection'
  | 'openPasteSpecial'
  | 'pasteSelection'
>;

type SpreadsheetContextCommands = Pick<
  SpreadsheetEditorCommands,
  | 'clearSelectedCells'
  | 'copySelection'
  | 'cutSelection'
  | 'openPasteSpecial'
  | 'pasteSelection'
>;

type SpreadsheetStructureCan = Pick<
  SpreadsheetEditorCanCommands,
  | 'deleteSelectedStructure'
  | 'insertSelectedStructure'
  | 'setSelectedStructureHidden'
  | 'setSelectedStructureSize'
  | 'sortSelectedCells'
>;

type SpreadsheetStructureCommands = Pick<
  SpreadsheetEditorCommands,
  | 'deleteSelectedStructure'
  | 'insertSelectedStructure'
  | 'setSelectedStructureHidden'
  | 'setSelectedStructureSize'
  | 'sortSelectedCells'
>;

type SpreadsheetSortCan = Pick<
  SpreadsheetEditorCanCommands,
  'sortSelectedCells'
>;

type SpreadsheetSortCommands = Pick<
  SpreadsheetEditorCommands,
  'sortSelectedCells'
>;

export function spreadsheetCoreContextMenuItems({
  can,
  commands,
  selection,
}: {
  can: SpreadsheetContextCan;
  commands: SpreadsheetContextCommands;
  selection: Pick<WorkSpreadsheetAgentSelection, 'clipboard' | 'reference'>;
}): WorkspaceContextMenuItem[] {
  return [
    {
      id: 'cut-cells',
      label: '剪切',
      icon: <Scissors size={14} />,
      shortcut: '⌘X',
      ariaKeyShortcut: 'Control+X Meta+X',
      disabled: !can.cutSelection(),
      onSelect: commands.cutSelection,
    },
    {
      id: 'copy-cells',
      label: '复制',
      icon: <Copy size={14} />,
      shortcut: '⌘C',
      ariaKeyShortcut: 'Control+C Meta+C',
      disabled: !can.copySelection(),
      onSelect: commands.copySelection,
    },
    {
      id: 'paste-cells',
      label: '粘贴',
      icon: <ClipboardPaste size={14} />,
      shortcut: '⌘V',
      ariaKeyShortcut: 'Control+V Meta+V',
      disabled: !can.pasteSelection(),
      onSelect: commands.pasteSelection,
    },
    {
      id: 'paste-special-cells',
      label: '选择性粘贴…',
      icon: <ClipboardPaste size={14} />,
      shortcut: '⌘⌥V',
      ariaKeyShortcut: 'Control+Alt+V Meta+Alt+V',
      disabled: !can.openPasteSpecial(),
      onSelect: commands.openPasteSpecial,
    },
    {
      id: 'clear-cells',
      label: '清除内容',
      icon: <Eraser size={14} />,
      shortcut: 'Delete',
      ariaKeyShortcut: 'Delete',
      separatorBefore: true,
      disabled: !can.clearSelectedCells(),
      onSelect: () => {
        if (!commands.clearSelectedCells()) {
          showToast(`无法清除选区 ${selection.reference}。`, 'error');
        }
      },
    },
  ];
}

export function spreadsheetStructureContextMenuItems({
  axis,
  can,
  commands,
  onResize,
}: {
  axis: SpreadsheetStructureAxis;
  can: SpreadsheetStructureCan;
  commands: SpreadsheetStructureCommands;
  onResize(axis: SpreadsheetStructureAxis): void;
}): WorkspaceContextMenuItem[] {
  const row = axis === 'row';
  const subject = row ? '行' : '列';
  const before = row ? '上方' : '左侧';
  const after = row ? '下方' : '右侧';
  const defaultSize = row ? 24 : 96;
  return [
    ...spreadsheetSortContextMenuItems({ can, commands, idSuffix: axis }),
    {
      id: `insert-${axis}-before`,
      label: `在${before}插入${subject}`,
      icon: <Plus size={14} />,
      separatorBefore: true,
      disabled: !can.insertSelectedStructure(axis, 'before'),
      onSelect: () => commands.insertSelectedStructure(axis, 'before'),
    },
    {
      id: `insert-${axis}-after`,
      label: `在${after}插入${subject}`,
      icon: <Plus size={14} />,
      disabled: !can.insertSelectedStructure(axis, 'after'),
      onSelect: () => commands.insertSelectedStructure(axis, 'after'),
    },
    {
      id: `delete-${axis}`,
      label: `删除所选${subject}`,
      icon: <Trash2 size={14} />,
      danger: true,
      disabled: !can.deleteSelectedStructure(axis),
      onSelect: () => commands.deleteSelectedStructure(axis),
    },
    {
      id: `resize-${axis}`,
      label: row ? '行高…' : '列宽…',
      icon: row ? <MoveVertical size={14} /> : <MoveHorizontal size={14} />,
      disabled: !can.setSelectedStructureSize(axis, defaultSize),
      onSelect: () => onResize(axis),
    },
    {
      id: `hide-${axis}`,
      label: `隐藏所选${subject}`,
      icon: <EyeOff size={14} />,
      separatorBefore: true,
      disabled: !can.setSelectedStructureHidden(axis, true),
      onSelect: () => commands.setSelectedStructureHidden(axis, true),
    },
    {
      id: `show-${axis}`,
      label: `取消隐藏${subject}`,
      icon: <Eye size={14} />,
      disabled: !can.setSelectedStructureHidden(axis, false),
      onSelect: () => commands.setSelectedStructureHidden(axis, false),
    },
  ];
}

export function spreadsheetSortContextMenuItems({
  can,
  commands,
  idSuffix = 'selection',
  separatorBefore = false,
}: {
  can: SpreadsheetSortCan;
  commands: SpreadsheetSortCommands;
  idSuffix?: string;
  separatorBefore?: boolean;
}): WorkspaceContextMenuItem[] {
  return [
    {
      id: `sort-${idSuffix}-ascending`,
      label: '升序排列',
      icon: <ArrowUp size={14} />,
      separatorBefore,
      disabled: !can.sortSelectedCells('ascending'),
      onSelect: () => commands.sortSelectedCells('ascending'),
    },
    {
      id: `sort-${idSuffix}-descending`,
      label: '降序排列',
      icon: <ArrowDown size={14} />,
      disabled: !can.sortSelectedCells('descending'),
      onSelect: () => commands.sortSelectedCells('descending'),
    },
  ];
}

export {
  browserSpreadsheetClipboard,
  copySpreadsheetSelection,
  parseSpreadsheetClipboardText,
  readSpreadsheetClipboardText,
  writeSpreadsheetClipboardText,
} from './spreadsheet-clipboard';
export type { SpreadsheetClipboardPort } from './spreadsheet-clipboard';
