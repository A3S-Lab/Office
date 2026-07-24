import type { OfficeKernelPresentationAlignment } from '../../../kernel/office-kernel-protocol';
import type { WorkSlideElement, WorkSlideTransition } from '../work-types';

export type PresentationViewMode = 'normal' | 'sorter';

export type PresentationEditorCommand =
  | { type: 'chart.add' }
  | { type: 'clipboard.copy' }
  | { type: 'clipboard.cut' }
  | { type: 'clipboard.paste' }
  | { type: 'comment.add' }
  | { type: 'comments.toggle' }
  | { type: 'design.toggle' }
  | {
      type: 'element.add';
      elementType: 'shape' | 'text';
    }
  | {
      type: 'element.align';
      alignment: OfficeKernelPresentationAlignment;
    }
  | {
      type: 'element.reorder';
      direction: -1 | 1;
    }
  | {
      type: 'element.update';
      patch: Partial<WorkSlideElement>;
      restoreTextFocus?: boolean;
    }
  | { type: 'history.redo' }
  | { type: 'history.undo' }
  | { type: 'image.request' }
  | { type: 'slide.add' }
  | { type: 'slide.delete' }
  | { type: 'slide.duplicate' }
  | { type: 'slide.background.set'; color: string }
  | { type: 'slideshow.start' }
  | { type: 'table.add' }
  | {
      type: 'transition.set';
      transition: WorkSlideTransition | undefined;
    }
  | { type: 'transition.applyToAll' }
  | { type: 'view.set'; mode: PresentationViewMode };

export type PresentationCommandResult = boolean | void | Promise<void>;

export interface PresentationCommandHandlers {
  addChart: () => PresentationCommandResult;
  addComment: () => PresentationCommandResult;
  addElement: (type: 'shape' | 'text') => PresentationCommandResult;
  addSlide: () => PresentationCommandResult;
  addTable: () => PresentationCommandResult;
  alignElement: (
    alignment: OfficeKernelPresentationAlignment,
  ) => PresentationCommandResult;
  applyTransitionToAll: () => PresentationCommandResult;
  copySelection: () => PresentationCommandResult;
  cutSelection: () => PresentationCommandResult;
  deleteSlide: () => PresentationCommandResult;
  duplicateSlide: () => PresentationCommandResult;
  pasteSelection: () => PresentationCommandResult;
  redo: () => PresentationCommandResult;
  reorderElement: (direction: -1 | 1) => PresentationCommandResult;
  requestImage: () => PresentationCommandResult;
  setBackground: (color: string) => PresentationCommandResult;
  setTransition: (
    transition: WorkSlideTransition | undefined,
  ) => PresentationCommandResult;
  setViewMode: (mode: PresentationViewMode) => PresentationCommandResult;
  startSlideshow: () => PresentationCommandResult;
  toggleComments: () => PresentationCommandResult;
  toggleDesign: () => PresentationCommandResult;
  undo: () => PresentationCommandResult;
  updateElement: (
    patch: Partial<WorkSlideElement>,
    options: { restoreTextFocus?: boolean },
  ) => PresentationCommandResult;
}

export type PresentationCommandDispatcher = (
  command: PresentationEditorCommand,
) => PresentationCommandResult;

export function createPresentationCommandDispatcher(
  handlers: PresentationCommandHandlers,
): PresentationCommandDispatcher {
  return (command) => {
    switch (command.type) {
      case 'chart.add':
        return handlers.addChart();
      case 'clipboard.copy':
        return handlers.copySelection();
      case 'clipboard.cut':
        return handlers.cutSelection();
      case 'clipboard.paste':
        return handlers.pasteSelection();
      case 'comment.add':
        return handlers.addComment();
      case 'comments.toggle':
        return handlers.toggleComments();
      case 'design.toggle':
        return handlers.toggleDesign();
      case 'element.add':
        return handlers.addElement(command.elementType);
      case 'element.align':
        return handlers.alignElement(command.alignment);
      case 'element.reorder':
        return handlers.reorderElement(command.direction);
      case 'element.update':
        return handlers.updateElement(command.patch, {
          restoreTextFocus: command.restoreTextFocus,
        });
      case 'history.redo':
        return handlers.redo();
      case 'history.undo':
        return handlers.undo();
      case 'image.request':
        return handlers.requestImage();
      case 'slide.add':
        return handlers.addSlide();
      case 'slide.delete':
        return handlers.deleteSlide();
      case 'slide.duplicate':
        return handlers.duplicateSlide();
      case 'slide.background.set':
        return handlers.setBackground(command.color);
      case 'slideshow.start':
        return handlers.startSlideshow();
      case 'table.add':
        return handlers.addTable();
      case 'transition.set':
        return handlers.setTransition(command.transition);
      case 'transition.applyToAll':
        return handlers.applyTransitionToAll();
      case 'view.set':
        return handlers.setViewMode(command.mode);
    }
  };
}
