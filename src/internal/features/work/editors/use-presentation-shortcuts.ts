import { useEffect } from 'react';
import type { WorkSlideElement } from '../work-types';
import { isOfficeShortcutBlocked } from './office-shortcuts';
import { presentationElementCanEditContent } from './presentation-selection';

export function usePresentationShortcuts({
  copySelection,
  cutSelection,
  deleteSelection,
  duplicateSelection,
  editingElementId,
  groupSelection,
  nudgeSelection,
  onAddSlide,
  onEditElement,
  onExitEditing,
  onRedo,
  onSelectElements,
  onStartSlideshow,
  onUndo,
  pasteSelection,
  preview,
  selectedElement,
  selectedElementCount,
  toggleBold,
  ungroupSelection,
}: {
  copySelection: () => boolean;
  cutSelection: () => boolean;
  deleteSelection: () => boolean;
  duplicateSelection: () => boolean;
  editingElementId: string | null;
  groupSelection: () => boolean;
  nudgeSelection: (key: string, distance: number) => boolean;
  onAddSlide: () => void;
  onEditElement: (id: string) => void;
  onExitEditing: () => void;
  onRedo: () => boolean;
  onSelectElements: (ids: readonly string[]) => void;
  onStartSlideshow?: () => void;
  onUndo: () => boolean;
  pasteSelection: () => boolean;
  preview: boolean;
  selectedElement: WorkSlideElement | null;
  selectedElementCount: number;
  toggleBold: () => boolean;
  ungroupSelection: () => boolean;
}) {
  useEffect(() => {
    if (preview) return;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.defaultPrevented) return;
      const commandKey = event.metaKey || event.ctrlKey;
      const key = event.key.toLocaleLowerCase();
      if (
        key === 'escape' &&
        editingElementId &&
        isPresentationTextEditingTarget(event.target)
      ) {
        const object = presentationObjectKeyboardTarget(event.target);
        event.preventDefault();
        onExitEditing();
        window.requestAnimationFrame(() => object?.focus());
        return;
      }
      const addSlideShortcut =
        !event.repeat &&
        !event.altKey &&
        ((event.ctrlKey && !event.metaKey && !event.shiftKey && key === 'm') ||
          (event.metaKey && !event.ctrlKey && event.shiftKey && key === 'n'));
      if (isOfficeShortcutBlocked(event.target)) {
        if (
          !event.repeat &&
          !commandKey &&
          !event.altKey &&
          !event.shiftKey &&
          key === 'f5' &&
          onStartSlideshow
        ) {
          event.preventDefault();
        }
        return;
      }
      const historyEditingTarget = isPresentationHistoryEditingTarget(
        event.target,
      );
      let handled = false;
      if (
        !event.repeat &&
        !commandKey &&
        !event.altKey &&
        !event.shiftKey &&
        key === 'f5' &&
        onStartSlideshow
      ) {
        onStartSlideshow();
        handled = true;
      } else if (addSlideShortcut) {
        onAddSlide();
        handled = true;
      } else if (
        !event.repeat &&
        !commandKey &&
        !event.altKey &&
        !event.shiftKey &&
        key === 'enter' &&
        selectedElement &&
        selectedElementCount === 1 &&
        presentationElementCanEditContent(selectedElement) &&
        isPresentationObjectKeyboardTarget(event.target)
      ) {
        onEditElement(selectedElement.id);
        handled = true;
      } else if (
        !event.repeat &&
        commandKey &&
        !event.altKey &&
        key === 'g' &&
        selectedElementCount > 0 &&
        isPresentationObjectKeyboardTarget(event.target)
      ) {
        handled = event.shiftKey ? ungroupSelection() : groupSelection();
      } else if (
        !event.repeat &&
        commandKey &&
        !event.altKey &&
        !event.shiftKey &&
        key === 'b' &&
        selectedElement &&
        isPresentationObjectKeyboardTarget(event.target)
      ) {
        handled = toggleBold();
      } else if (
        historyEditingTarget &&
        commandKey &&
        !event.altKey &&
        key === 'z'
      ) {
        handled = event.shiftKey ? onRedo() : onUndo();
      } else if (
        historyEditingTarget &&
        commandKey &&
        !event.altKey &&
        !event.shiftKey &&
        key === 'y'
      ) {
        handled = onRedo();
      } else if (isPresentationTextEditingTarget(event.target)) {
        return;
      } else if (commandKey && !event.altKey && key === 'z') {
        handled = event.shiftKey ? onRedo() : onUndo();
      } else if (
        commandKey &&
        !event.altKey &&
        !event.shiftKey &&
        key === 'y'
      ) {
        handled = onRedo();
      } else if (
        commandKey &&
        !event.altKey &&
        !event.shiftKey &&
        key === 'c'
      ) {
        handled = copySelection();
      } else if (
        commandKey &&
        !event.altKey &&
        !event.shiftKey &&
        key === 'x'
      ) {
        handled = cutSelection();
      } else if (
        commandKey &&
        !event.altKey &&
        !event.shiftKey &&
        key === 'v'
      ) {
        handled = pasteSelection();
      } else if (
        commandKey &&
        !event.altKey &&
        !event.shiftKey &&
        key === 'd'
      ) {
        handled = duplicateSelection();
      } else if (
        !commandKey &&
        !event.altKey &&
        selectedElementCount > 0 &&
        event.key.startsWith('Arrow') &&
        isPresentationObjectKeyboardTarget(event.target)
      ) {
        handled = nudgeSelection(event.key, event.shiftKey ? 5 : 1);
      } else if (
        (event.key === 'Delete' || event.key === 'Backspace') &&
        selectedElementCount > 0 &&
        isPresentationObjectKeyboardTarget(event.target)
      ) {
        handled = deleteSelection();
      } else if (
        key === 'escape' &&
        selectedElementCount > 0 &&
        isPresentationObjectKeyboardTarget(event.target)
      ) {
        onSelectElements([]);
        handled = true;
      }
      if (handled) event.preventDefault();
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [
    copySelection,
    cutSelection,
    deleteSelection,
    duplicateSelection,
    editingElementId,
    groupSelection,
    nudgeSelection,
    onAddSlide,
    onEditElement,
    onExitEditing,
    onRedo,
    onSelectElements,
    onStartSlideshow,
    onUndo,
    pasteSelection,
    preview,
    selectedElement,
    selectedElementCount,
    toggleBold,
    ungroupSelection,
  ]);
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
