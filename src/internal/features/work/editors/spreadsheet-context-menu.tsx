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
import {
  officeMessage,
  resolveOfficeMessages,
} from '../../../i18n/office-locale';
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
  const catalog = resolveOfficeMessages();
  return [
    {
      id: 'cut-cells',
      label: officeMessage(catalog, 'spreadsheet.context.cut'),
      icon: <Scissors size={14} />,
      shortcut: '⌘X',
      ariaKeyShortcut: 'Control+X Meta+X',
      disabled: !can.cutSelection(),
      onSelect: commands.cutSelection,
    },
    {
      id: 'copy-cells',
      label: officeMessage(catalog, 'spreadsheet.context.copy'),
      icon: <Copy size={14} />,
      shortcut: '⌘C',
      ariaKeyShortcut: 'Control+C Meta+C',
      disabled: !can.copySelection(),
      onSelect: commands.copySelection,
    },
    {
      id: 'paste-cells',
      label: officeMessage(catalog, 'spreadsheet.context.paste'),
      icon: <ClipboardPaste size={14} />,
      shortcut: '⌘V',
      ariaKeyShortcut: 'Control+V Meta+V',
      disabled: !can.pasteSelection(),
      onSelect: commands.pasteSelection,
    },
    {
      id: 'paste-special-cells',
      label: officeMessage(catalog, 'spreadsheet.context.pasteSpecial'),
      icon: <ClipboardPaste size={14} />,
      shortcut: '⌘⌥V',
      ariaKeyShortcut: 'Control+Alt+V Meta+Alt+V',
      disabled: !can.openPasteSpecial(),
      onSelect: commands.openPasteSpecial,
    },
    {
      id: 'clear-cells',
      label: officeMessage(catalog, 'spreadsheet.context.clearContents'),
      icon: <Eraser size={14} />,
      shortcut: 'Delete',
      ariaKeyShortcut: 'Delete',
      separatorBefore: true,
      disabled: !can.clearSelectedCells(),
      onSelect: () => {
        if (!commands.clearSelectedCells()) {
          showToast(
            officeMessage(catalog, 'spreadsheet.context.clearFailed', {
              reference: selection.reference,
            }),
            'error',
          );
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
  const catalog = resolveOfficeMessages();
  const row = axis === 'row';
  const subject = officeMessage(
    catalog,
    row ? 'spreadsheet.context.row' : 'spreadsheet.context.column',
  );
  const before = officeMessage(
    catalog,
    row ? 'spreadsheet.context.beforeRow' : 'spreadsheet.context.beforeColumn',
  );
  const after = officeMessage(
    catalog,
    row ? 'spreadsheet.context.afterRow' : 'spreadsheet.context.afterColumn',
  );
  const defaultSize = row ? 24 : 96;
  return [
    ...spreadsheetSortContextMenuItems({ can, commands, idSuffix: axis }),
    {
      id: `insert-${axis}-before`,
      label: officeMessage(catalog, 'spreadsheet.context.insertBefore', {
        before,
        subject,
      }),
      icon: <Plus size={14} />,
      separatorBefore: true,
      disabled: !can.insertSelectedStructure(axis, 'before'),
      onSelect: () => commands.insertSelectedStructure(axis, 'before'),
    },
    {
      id: `insert-${axis}-after`,
      label: officeMessage(catalog, 'spreadsheet.context.insertAfter', {
        after,
        subject,
      }),
      icon: <Plus size={14} />,
      disabled: !can.insertSelectedStructure(axis, 'after'),
      onSelect: () => commands.insertSelectedStructure(axis, 'after'),
    },
    {
      id: `delete-${axis}`,
      label: officeMessage(catalog, 'spreadsheet.context.deleteSelected', {
        subject,
      }),
      icon: <Trash2 size={14} />,
      danger: true,
      disabled: !can.deleteSelectedStructure(axis),
      onSelect: () => commands.deleteSelectedStructure(axis),
    },
    {
      id: `resize-${axis}`,
      label: officeMessage(
        catalog,
        row ? 'spreadsheet.context.rowHeight' : 'spreadsheet.context.columnWidth',
      ),
      icon: row ? <MoveVertical size={14} /> : <MoveHorizontal size={14} />,
      disabled: !can.setSelectedStructureSize(axis, defaultSize),
      onSelect: () => onResize(axis),
    },
    {
      id: `hide-${axis}`,
      label: officeMessage(catalog, 'spreadsheet.context.hideSelected', {
        subject,
      }),
      icon: <EyeOff size={14} />,
      separatorBefore: true,
      disabled: !can.setSelectedStructureHidden(axis, true),
      onSelect: () => commands.setSelectedStructureHidden(axis, true),
    },
    {
      id: `show-${axis}`,
      label: officeMessage(catalog, 'spreadsheet.context.unhide', { subject }),
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
  const catalog = resolveOfficeMessages();
  return [
    {
      id: `sort-${idSuffix}-ascending`,
      label: officeMessage(catalog, 'spreadsheet.context.sortAsc'),
      icon: <ArrowUp size={14} />,
      separatorBefore,
      disabled: !can.sortSelectedCells('ascending'),
      onSelect: () => commands.sortSelectedCells('ascending'),
    },
    {
      id: `sort-${idSuffix}-descending`,
      label: officeMessage(catalog, 'spreadsheet.context.sortDesc'),
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
