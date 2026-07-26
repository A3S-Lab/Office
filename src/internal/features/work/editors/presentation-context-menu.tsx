import { Copy, CopyPlus, Scissors, Trash2 } from 'lucide-react';
import type { WorkspaceContextMenuItem } from '../../workspace/components/workspace-context-menu';
import type {
  PresentationEditorCanCommands,
  PresentationEditorCommands,
} from './presentation-command-types';

type PresentationContextCan = Pick<
  PresentationEditorCanCommands,
  | 'copySelection'
  | 'cutSelection'
  | 'deleteSelection'
  | 'deleteSlide'
  | 'duplicateSelection'
  | 'duplicateSlide'
>;

type PresentationContextCommands = Pick<
  PresentationEditorCommands,
  | 'copySelection'
  | 'cutSelection'
  | 'deleteSelection'
  | 'deleteSlideById'
  | 'duplicateSelection'
  | 'duplicateSlide'
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
        id: 'duplicate-slide',
        label: '复制幻灯片',
        icon: <CopyPlus size={14} />,
        disabled: !can.duplicateSlide(),
        onSelect: () => void commands.duplicateSlide(),
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
