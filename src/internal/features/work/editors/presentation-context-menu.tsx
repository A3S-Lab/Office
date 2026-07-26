import {
  ClipboardPaste,
  Copy,
  CopyPlus,
  Plus,
  Scissors,
  Trash2,
} from 'lucide-react';
import type { WorkspaceContextMenuItem } from '../../workspace/components/workspace-context-menu';
import type {
  PresentationEditorCanCommands,
  PresentationEditorCommands,
} from './presentation-command-types';

type PresentationContextCan = Pick<
  PresentationEditorCanCommands,
  | 'addSlide'
  | 'copySelection'
  | 'cutSelection'
  | 'deleteSelection'
  | 'deleteSlide'
  | 'duplicateSelection'
  | 'duplicateSlide'
  | 'pasteSelection'
>;

type PresentationContextCommands = Pick<
  PresentationEditorCommands,
  | 'addSlide'
  | 'copySelection'
  | 'cutSelection'
  | 'deleteSelection'
  | 'deleteSlideById'
  | 'duplicateSelection'
  | 'duplicateSlide'
  | 'pasteSelection'
>;

export function presentationCoreContextMenuItems({
  can,
  commands,
  slideId,
  target,
}: {
  can: PresentationContextCan;
  commands: PresentationContextCommands;
  slideId: string;
  target: 'slide' | 'element';
}): WorkspaceContextMenuItem[] {
  if (target === 'slide') {
    return [
      {
        id: 'add-slide',
        label: '新建幻灯片',
        icon: <Plus size={14} />,
        shortcut: 'Ctrl+M',
        ariaKeyShortcut: 'Control+M Meta+Shift+N',
        disabled: !can.addSlide(),
        onSelect: () => void commands.addSlide(),
      },
      {
        id: 'duplicate-slide',
        label: '复制幻灯片',
        icon: <CopyPlus size={14} />,
        disabled: !can.duplicateSlide(),
        onSelect: () => void commands.duplicateSlide(),
      },
      {
        id: 'paste-slide',
        label: '粘贴',
        icon: <ClipboardPaste size={14} />,
        shortcut: '⌘V',
        ariaKeyShortcut: 'Control+V Meta+V',
        disabled: !can.pasteSelection(),
        onSelect: () => void commands.pasteSelection(),
      },
      {
        id: 'delete-slide',
        label: '删除幻灯片',
        icon: <Trash2 size={14} />,
        danger: true,
        disabled: !can.deleteSlide(),
        onSelect: () => void commands.deleteSlideById(slideId),
      },
    ];
  }

  return [
    {
      id: 'copy-object',
      label: '复制对象',
      icon: <Copy size={14} />,
      shortcut: '⌘C',
      ariaKeyShortcut: 'Control+C Meta+C',
      disabled: !can.copySelection(),
      onSelect: () => void commands.copySelection(),
    },
    {
      id: 'cut-object',
      label: '剪切对象',
      icon: <Scissors size={14} />,
      shortcut: '⌘X',
      ariaKeyShortcut: 'Control+X Meta+X',
      disabled: !can.cutSelection(),
      onSelect: () => void commands.cutSelection(),
    },
    {
      id: 'paste-object',
      label: '粘贴',
      icon: <ClipboardPaste size={14} />,
      shortcut: '⌘V',
      ariaKeyShortcut: 'Control+V Meta+V',
      disabled: !can.pasteSelection(),
      onSelect: () => void commands.pasteSelection(),
    },
    {
      id: 'duplicate-object',
      label: '创建副本',
      icon: <CopyPlus size={14} />,
      shortcut: '⌘D',
      ariaKeyShortcut: 'Control+D Meta+D',
      disabled: !can.duplicateSelection(),
      onSelect: () => void commands.duplicateSelection(),
    },
    {
      id: 'delete-object',
      label: '删除对象',
      icon: <Trash2 size={14} />,
      shortcut: 'Delete',
      ariaKeyShortcut: 'Delete',
      danger: true,
      disabled: !can.deleteSelection(),
      onSelect: () => void commands.deleteSelection(),
    },
  ];
}
