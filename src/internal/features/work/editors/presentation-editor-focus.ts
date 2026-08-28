import type {
  PresentationCommandResult,
  PresentationEditorCommands,
} from './presentation-command-types';
import type { WorkSlideElement } from '../work-types';
import { presentationElementCanEditContent } from './presentation-selection';

export interface PresentationObjectFocusState {
  editingElementId: string | null;
  selectedElementIds: readonly string[];
}

export interface PresentationWorkspaceFocusState
  extends PresentationObjectFocusState {
  selectedSlideId: string;
  viewMode: 'normal' | 'sorter';
}

const PRESENTATION_OBJECT_FOCUS_RETRY_FRAMES = 4;
const PRESENTATION_WORKSPACE_FOCUS_RETRY_FRAMES = 6;

export function presentationInitialEditingElementId(
  elements: readonly WorkSlideElement[],
): string | null {
  return (
    elements.find(
      (element) =>
        element.placeholder?.type === 'title' &&
        presentationElementCanEditContent(element) &&
        !element.text.trim() &&
        !element.textRuns?.some((run) => run.text.trim()),
    )?.id ?? null
  );
}

export function restorePresentationObjectFocus(
  container: HTMLElement | null,
  getState: () => PresentationObjectFocusState,
  focusOrigin: Element | null = document.activeElement,
): void {
  if (!container) return;
  const commandTrigger = focusOrigin;
  let lastRestoredTarget: HTMLElement | null = null;
  let remainingFrames = PRESENTATION_OBJECT_FOCUS_RETRY_FRAMES;

  const restore = () => {
    if (!container.isConnected || remainingFrames <= 0) return;
    remainingFrames -= 1;
    const { editingElementId, selectedElementIds } = getState();
    if (editingElementId) return;
    const selectedElementId = selectedElementIds.at(-1);
    if (!selectedElementId) {
      requestAnimationFrame(restore);
      return;
    }
    const target = Array.from(
      container.querySelectorAll<HTMLElement>('[data-slide-element-id]'),
    ).find(
      (candidate) =>
        candidate.dataset.slideElementId === selectedElementId &&
        candidate.dataset.slideElementSelected === 'true',
    );
    if (!target) {
      requestAnimationFrame(restore);
      return;
    }
    const activeElement = document.activeElement;
    const canRestore =
      activeElement === commandTrigger ||
      activeElement === target ||
      activeElement === lastRestoredTarget ||
      activeElement === document.body ||
      activeElement === document.documentElement ||
      !activeElement?.isConnected;
    if (!canRestore) return;
    if (activeElement !== target) target.focus({ preventScroll: true });
    lastRestoredTarget = target;
    requestAnimationFrame(restore);
  };

  requestAnimationFrame(restore);
}

export function restorePresentationWorkspaceFocus(
  root: HTMLElement | null,
  getState: () => PresentationWorkspaceFocusState,
  focusOrigin: Element | null = document.activeElement,
): void {
  if (!root) return;
  const commandTrigger = focusOrigin;
  let lastRestoredTarget: HTMLElement | null = null;
  let remainingFrames = PRESENTATION_WORKSPACE_FOCUS_RETRY_FRAMES;

  const restore = () => {
    if (!root.isConnected || remainingFrames <= 0) return;
    remainingFrames -= 1;
    const state = getState();
    if (state.viewMode === 'normal' && state.editingElementId) return;
    const target = presentationWorkspaceFocusTarget(root, state);
    if (!target) {
      requestAnimationFrame(restore);
      return;
    }
    const activeElement = document.activeElement;
    const canRestore =
      activeElement === commandTrigger ||
      activeElement === target ||
      activeElement === lastRestoredTarget ||
      activeElement === document.body ||
      activeElement === document.documentElement ||
      !activeElement?.isConnected;
    if (!canRestore) return;
    if (activeElement !== target) target.focus({ preventScroll: true });
    lastRestoredTarget = target;
    requestAnimationFrame(restore);
  };

  requestAnimationFrame(restore);
}

export function presentationWorkspaceFocusTarget(
  root: HTMLElement,
  state: PresentationWorkspaceFocusState,
): HTMLElement | undefined {
  if (state.viewMode === 'normal' && state.editingElementId) return undefined;
  const selectedElementId = state.selectedElementIds.at(-1);
  return state.viewMode === 'normal' && selectedElementId
    ? Array.from(
        root.querySelectorAll<HTMLElement>('[data-slide-element-id]'),
      ).find(
        (candidate) =>
          candidate.dataset.slideElementId === selectedElementId &&
          candidate.dataset.slideElementSelected === 'true',
      )
    : Array.from(
        root.querySelectorAll<HTMLElement>('[data-slide-thumbnail]'),
      ).find(
        (candidate) =>
          candidate.dataset.slideId === state.selectedSlideId &&
          candidate.classList.contains('active'),
      );
}

export function presentationCommandsWithObjectFocus(
  commands: PresentationEditorCommands,
  restoreFocus: () => void,
): PresentationEditorCommands {
  const afterSuccessfulCommand =
    <Arguments extends unknown[], Result extends PresentationCommandResult>(
      command: (...args: Arguments) => Result,
    ) =>
    (...args: Arguments): Result => {
      const result = command(...args);
      if (result instanceof Promise) {
        return result.then(() => restoreFocus()) as Result;
      }
      if (result !== false) restoreFocus();
      return result;
    };

  return {
    ...commands,
    addSlide: afterSuccessfulCommand(commands.addSlide),
    addChart: afterSuccessfulCommand(commands.addChart),
    addElement: afterSuccessfulCommand(commands.addElement),
    addImage: afterSuccessfulCommand(commands.addImage),
    addTable: afterSuccessfulCommand(commands.addTable),
    alignElement: afterSuccessfulCommand(commands.alignElement),
    applyTransitionToAll: afterSuccessfulCommand(commands.applyTransitionToAll),
    copySelection: afterSuccessfulCommand(commands.copySelection),
    cutSelection: afterSuccessfulCommand(commands.cutSelection),
    deleteSlide: afterSuccessfulCommand(commands.deleteSlide),
    distributeElements: afterSuccessfulCommand(commands.distributeElements),
    duplicateSlide: afterSuccessfulCommand(commands.duplicateSlide),
    groupElements: afterSuccessfulCommand(commands.groupElements),
    moveEntranceAnimation: afterSuccessfulCommand(
      commands.moveEntranceAnimation,
    ),
    pasteSelection: afterSuccessfulCommand(commands.pasteSelection),
    redo: afterSuccessfulCommand(commands.redo),
    reorderElement: afterSuccessfulCommand(commands.reorderElement),
    setTransition: afterSuccessfulCommand(commands.setTransition),
    setEntranceAnimation: afterSuccessfulCommand(commands.setEntranceAnimation),
    setViewMode: afterSuccessfulCommand(commands.setViewMode),
    toggleBold: afterSuccessfulCommand(commands.toggleBold),
    toggleItalic: afterSuccessfulCommand(commands.toggleItalic),
    toggleUnderline: afterSuccessfulCommand(commands.toggleUnderline),
    undo: afterSuccessfulCommand(commands.undo),
    ungroupElements: afterSuccessfulCommand(commands.ungroupElements),
    updateElement: afterSuccessfulCommand(commands.updateElement),
    updateEntranceAnimation: afterSuccessfulCommand(
      commands.updateEntranceAnimation,
    ),
  };
}
