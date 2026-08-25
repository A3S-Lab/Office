import type { OfficeKernelPresentationAlignment } from '../../../kernel/office-kernel-protocol';
import type {
  WorkPresentationContent,
  WorkSlideAnimation,
  WorkSlideAnimationEffect,
  WorkSlideElement,
  WorkSlideTransition,
  WorkSlideTextRun,
} from '../work-types';
import type { OfficeEditorCanCommands } from './office-editor-extension';
import type { PresentationDesignMode } from './presentation-editor-types';
import type { PresentationDistribution } from './presentation-selection';
import type { OfficeTableDimensions } from './office-table-dimensions';

export type PresentationViewMode = 'normal' | 'sorter';
export type PresentationSlideshowStart = 'beginning' | 'current';
export type PresentationCommandResult = boolean | void | Promise<void>;

export interface PresentationEditorCommands {
  addDesignPlaceholder: (type: 'body' | 'title') => PresentationCommandResult;
  addChart: () => PresentationCommandResult;
  addComment: () => PresentationCommandResult;
  addElement: (type: 'shape' | 'text') => PresentationCommandResult;
  addImage: (file: File) => PresentationCommandResult;
  addSlide: () => PresentationCommandResult;
  addTable: (dimensions: OfficeTableDimensions) => PresentationCommandResult;
  alignElement: (
    alignment: OfficeKernelPresentationAlignment,
  ) => PresentationCommandResult;
  applyTransitionToAll: (
    transition: WorkSlideTransition | undefined,
  ) => PresentationCommandResult;
  applyPresentationLayout: (layoutId: string) => PresentationCommandResult;
  closeComments: () => PresentationCommandResult;
  closeDesign: () => PresentationCommandResult;
  copySelection: () => boolean;
  createPresentationLayout: (copyCurrent: boolean) => PresentationCommandResult;
  cutSelection: () => boolean;
  deletePresentationComment: (
    slideId: string,
    commentId: string,
  ) => PresentationCommandResult;
  deletePresentationLayout: () => PresentationCommandResult;
  deleteSelection: () => boolean;
  deleteSlide: () => PresentationCommandResult;
  deleteSlideById: (slideId: string) => boolean;
  distributeElements: (
    direction: PresentationDistribution,
  ) => PresentationCommandResult;
  duplicateSelection: () => boolean;
  duplicateSlide: () => PresentationCommandResult;
  editDesign: (mode: PresentationDesignMode) => PresentationCommandResult;
  editElement: (id: string) => void;
  exitEditing: () => void;
  groupElements: () => boolean;
  instantiatePlaceholder: (
    definition: WorkSlideElement,
  ) => PresentationCommandResult;
  locatePresentationComment: (
    slideId: string,
    commentId: string,
  ) => PresentationCommandResult;
  moveEntranceAnimation: (direction: -1 | 1) => PresentationCommandResult;
  nudgeSelection: (key: string, distance: number) => boolean;
  openComment: (commentId: string) => PresentationCommandResult;
  pasteSelection: () => boolean;
  previewAnimations: () => PresentationCommandResult;
  redo: () => boolean;
  renamePresentationLayout: (name: string) => PresentationCommandResult;
  renamePresentationMaster: (name: string) => PresentationCommandResult;
  reorderElement: (direction: -1 | 1) => PresentationCommandResult;
  requestImage: () => PresentationCommandResult;
  selectElement: (
    elementId: string | null,
    additive?: boolean,
  ) => PresentationCommandResult;
  selectElements: (ids: readonly string[]) => void;
  selectSlide: (
    slideId: string,
    returnToSlideMode?: boolean,
  ) => PresentationCommandResult;
  setPresentationContent: (
    content: WorkPresentationContent,
  ) => PresentationCommandResult;
  setBackground: (color: string) => PresentationCommandResult;
  setEntranceAnimation: (
    effect: WorkSlideAnimationEffect | undefined,
  ) => PresentationCommandResult;
  setPresentationLayoutBackground: (
    color: string | undefined,
  ) => PresentationCommandResult;
  setPresentationMasterBackground: (color: string) => PresentationCommandResult;
  setTransition: (
    transition: WorkSlideTransition | undefined,
  ) => PresentationCommandResult;
  setViewMode: (mode: PresentationViewMode) => PresentationCommandResult;
  startSlideshow: (
    source: PresentationSlideshowStart,
  ) => PresentationCommandResult;
  toggleBold: () => boolean;
  toggleItalic: () => boolean;
  toggleUnderline: () => boolean;
  toggleComments: () => PresentationCommandResult;
  toggleDesign: () => PresentationCommandResult;
  togglePresentationLayoutBackground: (
    enabled: boolean,
  ) => PresentationCommandResult;
  undo: () => boolean;
  ungroupElements: () => boolean;
  updateElement: (
    patch: Partial<WorkSlideElement>,
    options?: { restoreTextFocus?: boolean },
  ) => PresentationCommandResult;
  updateEntranceAnimation: (
    patch: Partial<WorkSlideAnimation>,
  ) => PresentationCommandResult;
  updateNotes: (notes: string) => PresentationCommandResult;
  updatePresentationComment: (
    slideId: string,
    commentId: string,
    text: string,
  ) => PresentationCommandResult;
  updateTextElement: (
    elementId: string,
    value: { text: string; textRuns?: WorkSlideTextRun[] },
  ) => PresentationCommandResult;
}

export type PresentationEditorCanCommands =
  OfficeEditorCanCommands<PresentationEditorCommands>;

export interface PresentationHistoryCommandPort {
  canRedo: boolean;
  canUndo: boolean;
  redo: () => boolean;
  undo: () => boolean;
}

export interface PresentationDocumentCommandPort {
  setContent: (content: WorkPresentationContent) => PresentationCommandResult;
}

export interface PresentationClipboardCommandPort {
  canCopySelection: boolean;
  canCutSelection: boolean;
  canPasteSelection: boolean;
  copySelection: () => boolean;
  cutSelection: () => boolean;
  pasteSelection: () => boolean;
}

export interface PresentationSlideCommandPort {
  canAddSlide: boolean;
  canApplyTransitionToAll: (
    transition: WorkSlideTransition | undefined,
  ) => boolean;
  canDeleteSlide: boolean;
  canDuplicateSlide: boolean;
  canSetTransition: boolean;
  addSlide: () => PresentationCommandResult;
  applyTransitionToAll: (
    transition: WorkSlideTransition | undefined,
  ) => PresentationCommandResult;
  deleteSlide: () => PresentationCommandResult;
  deleteSlideById: (slideId: string) => boolean;
  duplicateSlide: () => PresentationCommandResult;
  selectSlide: (
    slideId: string,
    returnToSlideMode: boolean,
  ) => PresentationCommandResult;
  setBackground: (color: string) => PresentationCommandResult;
  setTransition: (
    transition: WorkSlideTransition | undefined,
  ) => PresentationCommandResult;
  updateNotes: (notes: string) => PresentationCommandResult;
}

export interface PresentationAnimationCommandPort {
  canMoveEntranceAnimation: (direction: -1 | 1) => boolean;
  canPreviewAnimations: boolean;
  canSetEntranceAnimation: boolean;
  canUpdateEntranceAnimation: boolean;
  moveEntranceAnimation: (direction: -1 | 1) => PresentationCommandResult;
  previewAnimations: () => PresentationCommandResult;
  setEntranceAnimation: (
    effect: WorkSlideAnimationEffect | undefined,
  ) => PresentationCommandResult;
  updateEntranceAnimation: (
    patch: Partial<WorkSlideAnimation>,
  ) => PresentationCommandResult;
}

export interface PresentationInsertCommandPort {
  enabled: boolean;
  addChart: () => PresentationCommandResult;
  addElement: (type: 'shape' | 'text') => PresentationCommandResult;
  addImage: (file: File) => PresentationCommandResult;
  addTable: (dimensions: OfficeTableDimensions) => PresentationCommandResult;
  instantiatePlaceholder: (
    definition: WorkSlideElement,
  ) => PresentationCommandResult;
  requestImage: () => PresentationCommandResult;
}

export interface PresentationDesignCommandPort {
  addPlaceholder: (type: 'body' | 'title') => PresentationCommandResult;
  applyLayout: (layoutId: string) => PresentationCommandResult;
  canDeleteLayout: boolean;
  close: () => PresentationCommandResult;
  createLayout: (copyCurrent: boolean) => PresentationCommandResult;
  deleteLayout: () => PresentationCommandResult;
  edit: (mode: PresentationDesignMode) => PresentationCommandResult;
  renameLayout: (name: string) => PresentationCommandResult;
  renameMaster: (name: string) => PresentationCommandResult;
  setLayoutBackground: (color: string | undefined) => PresentationCommandResult;
  setMasterBackground: (color: string) => PresentationCommandResult;
  toggleLayoutBackground: (enabled: boolean) => PresentationCommandResult;
}

export interface PresentationElementCommandPort {
  canAlignElement: boolean;
  canDistributeElements: boolean;
  canGroupElements: boolean;
  canReorderElement: boolean;
  canUngroupElements: boolean;
  canUpdateElement: boolean;
  alignElement: (
    alignment: OfficeKernelPresentationAlignment,
  ) => PresentationCommandResult;
  distributeElements: (
    direction: PresentationDistribution,
  ) => PresentationCommandResult;
  groupElements: () => boolean;
  reorderElement: (direction: -1 | 1) => PresentationCommandResult;
  ungroupElements: () => boolean;
  updateElement: (
    patch: Partial<WorkSlideElement>,
    options: { restoreTextFocus?: boolean },
  ) => PresentationCommandResult;
  updateTextElement: (
    elementId: string,
    value: { text: string; textRuns?: WorkSlideTextRun[] },
  ) => PresentationCommandResult;
}

export interface PresentationSelectionCommandPort {
  canDeleteSelection: boolean;
  canDuplicateSelection: boolean;
  canEditElement: (id: string) => boolean;
  canExitEditing: boolean;
  canNudgeSelection: boolean;
  canToggleBold: boolean;
  canToggleItalic: boolean;
  canToggleUnderline: boolean;
  deleteSelection: () => boolean;
  duplicateSelection: () => boolean;
  editElement: (id: string) => void;
  exitEditing: () => void;
  nudgeSelection: (key: string, distance: number) => boolean;
  selectElement: (elementId: string | null, additive: boolean) => void;
  selectElements: (ids: readonly string[]) => void;
  toggleBold: () => boolean;
  toggleItalic: () => boolean;
  toggleUnderline: () => boolean;
}

export interface PresentationReviewCommandPort {
  canAddComment: boolean;
  addComment: () => PresentationCommandResult;
  closeComments: () => PresentationCommandResult;
  deleteComment: (
    slideId: string,
    commentId: string,
  ) => PresentationCommandResult;
  locateComment: (
    slideId: string,
    commentId: string,
  ) => PresentationCommandResult;
  openComment: (commentId: string) => PresentationCommandResult;
  toggleComments: () => PresentationCommandResult;
  updateComment: (
    slideId: string,
    commentId: string,
    text: string,
  ) => PresentationCommandResult;
}

export interface PresentationViewCommandPort {
  canStartSlideshow: boolean;
  setViewMode: (mode: PresentationViewMode) => PresentationCommandResult;
  startSlideshow: (
    source: PresentationSlideshowStart,
  ) => PresentationCommandResult;
  toggleDesign: () => PresentationCommandResult;
}

export interface PresentationKeyboardCommandPort {
  editingElementId: string | null;
  selectedElement: WorkSlideElement | null;
  selectedElementCount: number;
}

export interface PresentationCommandContext {
  animations: PresentationAnimationCommandPort;
  clipboard: PresentationClipboardCommandPort;
  design: PresentationDesignCommandPort;
  document: PresentationDocumentCommandPort;
  elements: PresentationElementCommandPort;
  history: PresentationHistoryCommandPort;
  insert: PresentationInsertCommandPort;
  keyboard: PresentationKeyboardCommandPort;
  review: PresentationReviewCommandPort;
  selection: PresentationSelectionCommandPort;
  slides: PresentationSlideCommandPort;
  view: PresentationViewCommandPort;
}
