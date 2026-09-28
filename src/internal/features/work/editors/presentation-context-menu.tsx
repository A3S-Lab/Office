import {
  ClipboardPaste,
  Copy,
  CopyPlus,
  Plus,
  Scissors,
  Trash2,
} from 'lucide-react';
import {
  officeMessage,
  resolveOfficeMessages,
} from '../../../i18n/office-locale';
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
  const catalog = resolveOfficeMessages();
  if (target === 'slide') {
    return [
      {
        id: 'add-slide',
        label: officeMessage(catalog, 'presentation.context.newSlide'),
        icon: <Plus size={14} />,
        shortcut: 'Ctrl+M / ⌘⇧N',
        ariaKeyShortcut: 'Control+M Meta+Shift+N',
        disabled: !can.addSlide(),
        onSelect: () => void commands.addSlide(),
      },
      {
        id: 'duplicate-slide',
        label: officeMessage(catalog, 'presentation.context.duplicateSlide'),
        icon: <CopyPlus size={14} />,
        shortcut: 'Ctrl+D / ⌘D',
        ariaKeyShortcut: 'Control+D Meta+D',
        disabled: !can.duplicateSlide(),
        onSelect: () => void commands.duplicateSlide(),
      },
      {
        id: 'paste-slide',
        label: officeMessage(catalog, 'presentation.context.paste'),
        icon: <ClipboardPaste size={14} />,
        shortcut: '⌘V',
        ariaKeyShortcut: 'Control+V Meta+V',
        disabled: !can.pasteSelection(),
        onSelect: () => void commands.pasteSelection(),
      },
      {
        id: 'delete-slide',
        label: officeMessage(catalog, 'presentation.context.deleteSlide'),
        icon: <Trash2 size={14} />,
        shortcut: 'Delete / Backspace',
        ariaKeyShortcut: 'Delete Backspace',
        danger: true,
        disabled: !can.deleteSlide(),
        onSelect: () => void commands.deleteSlideById(slideId),
      },
    ];
  }

  return [
    {
      id: 'copy-object',
      label: officeMessage(catalog, 'presentation.context.copyObject'),
      icon: <Copy size={14} />,
      shortcut: '⌘C',
      ariaKeyShortcut: 'Control+C Meta+C',
      disabled: !can.copySelection(),
      onSelect: () => void commands.copySelection(),
    },
    {
      id: 'cut-object',
      label: officeMessage(catalog, 'presentation.context.cutObject'),
      icon: <Scissors size={14} />,
      shortcut: '⌘X',
      ariaKeyShortcut: 'Control+X Meta+X',
      disabled: !can.cutSelection(),
      onSelect: () => void commands.cutSelection(),
    },
    {
      id: 'paste-object',
      label: officeMessage(catalog, 'presentation.context.paste'),
      icon: <ClipboardPaste size={14} />,
      shortcut: '⌘V',
      ariaKeyShortcut: 'Control+V Meta+V',
      disabled: !can.pasteSelection(),
      onSelect: () => void commands.pasteSelection(),
    },
    {
      id: 'duplicate-object',
      label: officeMessage(catalog, 'presentation.context.duplicateObject'),
      icon: <CopyPlus size={14} />,
      shortcut: '⌘D',
      ariaKeyShortcut: 'Control+D Meta+D',
      disabled: !can.duplicateSelection(),
      onSelect: () => void commands.duplicateSelection(),
    },
    {
      id: 'delete-object',
      label: officeMessage(catalog, 'presentation.context.deleteObject'),
      icon: <Trash2 size={14} />,
      shortcut: 'Delete',
      ariaKeyShortcut: 'Delete',
      danger: true,
      disabled: !can.deleteSelection(),
      onSelect: () => void commands.deleteSelection(),
    },
  ];
}
