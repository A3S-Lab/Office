import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';
import { isOfficeShortcutBlocked } from './office-shortcuts';
import type {
  PresentationCommandContext,
  PresentationEditorCanCommands,
  PresentationEditorCommands,
} from './presentation-command-types';
import { presentationElementCanEditContent } from './presentation-selection';

export function createPresentationKeyboardExtension(): OfficeEditorExtension<
  PresentationCommandContext,
  PresentationEditorCommands
> {
  return createOfficeEditorExtension<
    PresentationCommandContext,
    PresentationEditorCommands
  >({
    name: 'presentationKeyboardShortcuts',
    addKeyboardShortcuts: () => ({
      'Shift-F5': ({ can, commands }, event) => {
        if (event.repeat || !can.startSlideshow('current')) return false;
        if (isOfficeShortcutBlocked(event.target)) return true;
        commands.startSlideshow('current');
        return true;
      },
      F5: ({ can, commands }, event) => {
        if (event.repeat || !can.startSlideshow('beginning')) return false;
        if (isOfficeShortcutBlocked(event.target)) return true;
        commands.startSlideshow('beginning');
        return true;
      },
      'Control-m': ({ commands }, event) => {
        if (presentationShortcutBlocked(event)) return false;
        commands.addSlide();
        return true;
      },
      'Meta-Shift-n': ({ commands }, event) => {
        if (presentationShortcutBlocked(event)) return false;
        commands.addSlide();
        return true;
      },
      Enter: ({ can, commands, context }, event) => {
        const selected = context.keyboard.selectedElement;
        if (
          presentationShortcutBlocked(event) ||
          context.keyboard.selectedElementCount !== 1 ||
          !selected ||
          !presentationElementCanEditContent(selected) ||
          !isPresentationObjectKeyboardTarget(event.target) ||
          !can.editElement(selected.id)
        ) {
          return false;
        }
        commands.editElement(selected.id);
        return true;
      },
      'Mod-g': ({ can, commands, context }, event) =>
        runPresentationObjectShortcut(
          event,
          context.keyboard.selectedElementCount,
          can.groupElements,
          commands.groupElements,
        ),
      'Mod-Shift-g': ({ can, commands, context }, event) =>
        runPresentationObjectShortcut(
          event,
          context.keyboard.selectedElementCount,
          can.ungroupElements,
          commands.ungroupElements,
        ),
      'Mod-b': ({ can, commands, context }, event) =>
        runPresentationObjectShortcut(
          event,
          context.keyboard.selectedElementCount,
          can.toggleBold,
          commands.toggleBold,
        ),
      'Mod-z': ({ can, commands }, event) =>
        runPresentationHistoryShortcut(event, can.undo, commands.undo),
      'Mod-Shift-z': ({ can, commands }, event) =>
        runPresentationHistoryShortcut(event, can.redo, commands.redo),
      'Mod-y': ({ can, commands }, event) =>
        runPresentationHistoryShortcut(event, can.redo, commands.redo),
      'Mod-c': ({ can, commands }, event) =>
        runPresentationSelectionShortcut(
          event,
          can.copySelection,
          commands.copySelection,
        ),
      'Mod-x': ({ can, commands }, event) =>
        runPresentationSelectionShortcut(
          event,
          can.cutSelection,
          commands.cutSelection,
        ),
      'Mod-v': ({ can, commands }, event) =>
        runPresentationSelectionShortcut(
          event,
          can.pasteSelection,
          commands.pasteSelection,
        ),
      'Mod-d': ({ can, commands }, event) =>
        runPresentationSelectionShortcut(
          event,
          can.duplicateSelection,
          commands.duplicateSelection,
        ),
      ArrowLeft: ({ can, commands, context }, event) =>
        runPresentationNudgeShortcut(event, context, can, commands, 1),
      ArrowRight: ({ can, commands, context }, event) =>
        runPresentationNudgeShortcut(event, context, can, commands, 1),
      ArrowUp: ({ can, commands, context }, event) =>
        runPresentationNudgeShortcut(event, context, can, commands, 1),
      ArrowDown: ({ can, commands, context }, event) =>
        runPresentationNudgeShortcut(event, context, can, commands, 1),
      'Shift-ArrowLeft': ({ can, commands, context }, event) =>
        runPresentationNudgeShortcut(event, context, can, commands, 5),
      'Shift-ArrowRight': ({ can, commands, context }, event) =>
        runPresentationNudgeShortcut(event, context, can, commands, 5),
      'Shift-ArrowUp': ({ can, commands, context }, event) =>
        runPresentationNudgeShortcut(event, context, can, commands, 5),
      'Shift-ArrowDown': ({ can, commands, context }, event) =>
        runPresentationNudgeShortcut(event, context, can, commands, 5),
      Delete: ({ can, commands, context }, event) =>
        runPresentationObjectShortcut(
          event,
          context.keyboard.selectedElementCount,
          can.deleteSelection,
          commands.deleteSelection,
        ),
      Backspace: ({ can, commands, context }, event) =>
        runPresentationObjectShortcut(
          event,
          context.keyboard.selectedElementCount,
          can.deleteSelection,
          commands.deleteSelection,
        ),
      Escape: ({ can, commands, context }, event) =>
        runPresentationEscapeShortcut(event, context, can, commands),
    }),
  });
}

function runPresentationHistoryShortcut(
  event: KeyboardEvent,
  canExecute: () => boolean,
  execute: () => boolean,
): boolean {
  if (presentationShortcutBlocked(event)) return false;
  if (
    isPresentationTextEditingTarget(event.target) &&
    !isPresentationHistoryEditingTarget(event.target)
  ) {
    return false;
  }
  return canExecute() && execute();
}

function runPresentationSelectionShortcut(
  event: KeyboardEvent,
  canExecute: () => boolean,
  execute: () => boolean,
): boolean {
  if (
    presentationShortcutBlocked(event) ||
    isPresentationTextEditingTarget(event.target) ||
    !canExecute()
  ) {
    return false;
  }
  return execute();
}

function runPresentationObjectShortcut(
  event: KeyboardEvent,
  selectedElementCount: number,
  canExecute: () => boolean,
  execute: () => boolean,
): boolean {
  if (
    presentationShortcutBlocked(event) ||
    selectedElementCount === 0 ||
    !isPresentationObjectKeyboardTarget(event.target) ||
    !canExecute()
  ) {
    return false;
  }
  return execute();
}

function runPresentationNudgeShortcut(
  event: KeyboardEvent,
  context: PresentationCommandContext,
  can: PresentationEditorCanCommands,
  commands: PresentationEditorCommands,
  distance: number,
): boolean {
  return runPresentationObjectShortcut(
    event,
    context.keyboard.selectedElementCount,
    () => can.nudgeSelection(event.key, distance),
    () => commands.nudgeSelection(event.key, distance),
  );
}

function runPresentationEscapeShortcut(
  event: KeyboardEvent,
  context: PresentationCommandContext,
  can: PresentationEditorCanCommands,
  commands: PresentationEditorCommands,
): boolean {
  if (
    context.keyboard.editingElementId &&
    isPresentationTextEditingTarget(event.target) &&
    can.exitEditing()
  ) {
    const object = presentationObjectKeyboardTarget(event.target);
    commands.exitEditing();
    window.requestAnimationFrame(() => object?.focus());
    return true;
  }
  if (
    presentationShortcutBlocked(event) ||
    context.keyboard.selectedElementCount === 0 ||
    !isPresentationObjectKeyboardTarget(event.target)
  ) {
    return false;
  }
  commands.selectElements([]);
  return true;
}

function presentationShortcutBlocked(event: KeyboardEvent): boolean {
  return event.repeat || isOfficeShortcutBlocked(event.target);
}

function isPresentationTextEditingTarget(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) return false;
  return (
    target instanceof HTMLInputElement ||
    target instanceof HTMLTextAreaElement ||
    target instanceof HTMLSelectElement ||
    target.isContentEditable ||
    Boolean(target.closest('[data-slide-editor]'))
  );
}

function isPresentationHistoryEditingTarget(
  target: EventTarget | null,
): boolean {
  if (
    !(
      target instanceof HTMLInputElement ||
      target instanceof HTMLTextAreaElement ||
      target instanceof HTMLSelectElement
    )
  ) {
    return false;
  }
  return Boolean(target.closest('.work-presentation-editor'));
}

function isPresentationObjectKeyboardTarget(
  target: EventTarget | null,
): boolean {
  return Boolean(presentationObjectKeyboardTarget(target));
}

function presentationObjectKeyboardTarget(
  target: EventTarget | null,
): HTMLElement | null {
  return target instanceof HTMLElement
    ? target.closest<HTMLElement>(
        '[data-slide-element-origin], [data-presentation-selection-control]',
      )
    : null;
}
