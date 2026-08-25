import { hasPresentationClipboard } from '../work-presentation-clipboard';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';
import type {
  PresentationCommandContext,
  PresentationEditorCommands,
} from './presentation-command-types';
import { createPresentationKeyboardExtension } from './presentation-keyboard-extension';

export function createPresentationEditorExtensions(): readonly OfficeEditorExtension<
  PresentationCommandContext,
  PresentationEditorCommands
>[] {
  return [
    createOfficeEditorExtension<
      PresentationCommandContext,
      PresentationEditorCommands
    >({
      name: 'presentationDocument',
      addCommands: () => ({
        setPresentationContent: {
          execute: ({ document }, content) => document.setContent(content),
        },
      }),
    }),
    createOfficeEditorExtension<
      PresentationCommandContext,
      PresentationEditorCommands
    >({
      name: 'presentationHistory',
      addCommands: () => ({
        redo: {
          canExecute: ({ history }) => history.canRedo,
          execute: ({ history }) => history.redo(),
        },
        undo: {
          canExecute: ({ history }) => history.canUndo,
          execute: ({ history }) => history.undo(),
        },
      }),
    }),
    createOfficeEditorExtension<
      PresentationCommandContext,
      PresentationEditorCommands,
      { hasContent: boolean }
    >({
      name: 'presentationClipboard',
      addStorage: () => ({
        hasContent: hasPresentationClipboard(),
      }),
      addCommands: ({ storage }) => ({
        copySelection: {
          canExecute: ({ clipboard }) => clipboard.canCopySelection,
          execute: ({ clipboard }) => {
            const copied = clipboard.copySelection();
            if (copied) storage.hasContent = true;
            return copied;
          },
        },
        cutSelection: {
          canExecute: ({ clipboard }) => clipboard.canCutSelection,
          execute: ({ clipboard }) => {
            const cut = clipboard.cutSelection();
            if (cut) storage.hasContent = true;
            return cut;
          },
        },
        pasteSelection: {
          canExecute: ({ clipboard }) =>
            clipboard.canPasteSelection &&
            (storage.hasContent || hasPresentationClipboard()),
          execute: ({ clipboard }) => clipboard.pasteSelection(),
        },
      }),
    }),
    createOfficeEditorExtension<
      PresentationCommandContext,
      PresentationEditorCommands
    >({
      name: 'presentationSlides',
      addCommands: () => ({
        addSlide: {
          canExecute: ({ slides }) => slides.canAddSlide,
          execute: ({ slides }) => slides.addSlide(),
        },
        applyTransitionToAll: {
          canExecute: ({ slides }, transition) =>
            slides.canApplyTransitionToAll(transition),
          execute: ({ slides }, transition) =>
            slides.applyTransitionToAll(transition),
        },
        deleteSlide: {
          canExecute: ({ slides }) => slides.canDeleteSlide,
          execute: ({ slides }) => slides.deleteSlide(),
        },
        deleteSlideById: {
          canExecute: ({ slides }) => slides.canDeleteSlide,
          execute: ({ slides }, slideId) => slides.deleteSlideById(slideId),
        },
        duplicateSlide: {
          canExecute: ({ slides }) => slides.canDuplicateSlide,
          execute: ({ slides }) => slides.duplicateSlide(),
        },
        selectSlide: {
          execute: ({ slides }, slideId, returnToSlideMode = false) =>
            slides.selectSlide(slideId, returnToSlideMode),
        },
        setBackground: {
          execute: ({ slides }, color) => slides.setBackground(color),
        },
        setTransition: {
          canExecute: ({ slides }) => slides.canSetTransition,
          execute: ({ slides }, transition) => slides.setTransition(transition),
        },
        updateNotes: {
          execute: ({ slides }, notes) => slides.updateNotes(notes),
        },
      }),
    }),
    createOfficeEditorExtension<
      PresentationCommandContext,
      PresentationEditorCommands
    >({
      name: 'presentationAnimations',
      addCommands: () => ({
        moveEntranceAnimation: {
          canExecute: ({ animations }, direction) =>
            animations.canMoveEntranceAnimation(direction),
          execute: ({ animations }, direction) =>
            animations.moveEntranceAnimation(direction),
        },
        previewAnimations: {
          canExecute: ({ animations }) => animations.canPreviewAnimations,
          execute: ({ animations }) => animations.previewAnimations(),
        },
        setEntranceAnimation: {
          canExecute: ({ animations }) => animations.canSetEntranceAnimation,
          execute: ({ animations }, effect) =>
            animations.setEntranceAnimation(effect),
        },
        updateEntranceAnimation: {
          canExecute: ({ animations }) => animations.canUpdateEntranceAnimation,
          execute: ({ animations }, patch) =>
            animations.updateEntranceAnimation(patch),
        },
      }),
    }),
    createOfficeEditorExtension<
      PresentationCommandContext,
      PresentationEditorCommands
    >({
      name: 'presentationElements',
      addCommands: () => ({
        alignElement: {
          canExecute: ({ elements }) => elements.canAlignElement,
          execute: ({ elements }, alignment) =>
            elements.alignElement(alignment),
        },
        distributeElements: {
          canExecute: ({ elements }) => elements.canDistributeElements,
          execute: ({ elements }, direction) =>
            elements.distributeElements(direction),
        },
        groupElements: {
          canExecute: ({ elements }) => elements.canGroupElements,
          execute: ({ elements }) => elements.groupElements(),
        },
        reorderElement: {
          canExecute: ({ elements }) => elements.canReorderElement,
          execute: ({ elements }, direction) =>
            elements.reorderElement(direction),
        },
        ungroupElements: {
          canExecute: ({ elements }) => elements.canUngroupElements,
          execute: ({ elements }) => elements.ungroupElements(),
        },
        updateElement: {
          canExecute: ({ elements }) => elements.canUpdateElement,
          execute: ({ elements }, patch, options = {}) =>
            elements.updateElement(patch, options),
        },
        updateTextElement: {
          execute: ({ elements }, elementId, value) =>
            elements.updateTextElement(elementId, value),
        },
      }),
    }),
    createOfficeEditorExtension<
      PresentationCommandContext,
      PresentationEditorCommands
    >({
      name: 'presentationDesign',
      addCommands: () => ({
        addDesignPlaceholder: {
          execute: ({ design }, type) => design.addPlaceholder(type),
        },
        applyPresentationLayout: {
          execute: ({ design }, layoutId) => design.applyLayout(layoutId),
        },
        closeDesign: {
          execute: ({ design }) => design.close(),
        },
        createPresentationLayout: {
          execute: ({ design }, copyCurrent) =>
            design.createLayout(copyCurrent),
        },
        deletePresentationLayout: {
          canExecute: ({ design }) => design.canDeleteLayout,
          execute: ({ design }) => design.deleteLayout(),
        },
        editDesign: {
          execute: ({ design }, mode) => design.edit(mode),
        },
        renamePresentationLayout: {
          execute: ({ design }, name) => design.renameLayout(name),
        },
        renamePresentationMaster: {
          execute: ({ design }, name) => design.renameMaster(name),
        },
        setPresentationLayoutBackground: {
          execute: ({ design }, color) => design.setLayoutBackground(color),
        },
        setPresentationMasterBackground: {
          execute: ({ design }, color) => design.setMasterBackground(color),
        },
        togglePresentationLayoutBackground: {
          execute: ({ design }, enabled) =>
            design.toggleLayoutBackground(enabled),
        },
      }),
    }),
    createOfficeEditorExtension<
      PresentationCommandContext,
      PresentationEditorCommands
    >({
      name: 'presentationInsert',
      addCommands: () => ({
        addChart: {
          canExecute: ({ insert }) => insert.enabled,
          execute: ({ insert }) => insert.addChart(),
        },
        addElement: {
          canExecute: ({ insert }) => insert.enabled,
          execute: ({ insert }, type) => insert.addElement(type),
        },
        addImage: {
          canExecute: ({ insert }) => insert.enabled,
          execute: ({ insert }, file) => insert.addImage(file),
        },
        addTable: {
          canExecute: ({ insert }) => insert.enabled,
          execute: ({ insert }, dimensions) => insert.addTable(dimensions),
        },
        instantiatePlaceholder: {
          canExecute: ({ insert }) => insert.enabled,
          execute: ({ insert }, definition) =>
            insert.instantiatePlaceholder(definition),
        },
        requestImage: {
          canExecute: ({ insert }) => insert.enabled,
          execute: ({ insert }) => insert.requestImage(),
        },
      }),
    }),
    createOfficeEditorExtension<
      PresentationCommandContext,
      PresentationEditorCommands
    >({
      name: 'presentationSelection',
      addCommands: () => ({
        deleteSelection: {
          canExecute: ({ selection }) => selection.canDeleteSelection,
          execute: ({ selection }) => selection.deleteSelection(),
        },
        duplicateSelection: {
          canExecute: ({ selection }) => selection.canDuplicateSelection,
          execute: ({ selection }) => selection.duplicateSelection(),
        },
        editElement: {
          canExecute: ({ selection }, id) => selection.canEditElement(id),
          execute: ({ selection }, id) => selection.editElement(id),
        },
        exitEditing: {
          canExecute: ({ selection }) => selection.canExitEditing,
          execute: ({ selection }) => selection.exitEditing(),
        },
        nudgeSelection: {
          canExecute: ({ selection }) => selection.canNudgeSelection,
          execute: ({ selection }, key, distance) =>
            selection.nudgeSelection(key, distance),
        },
        selectElement: {
          execute: ({ selection }, elementId, additive = false) =>
            selection.selectElement(elementId, additive),
        },
        selectElements: {
          execute: ({ selection }, ids) => selection.selectElements(ids),
        },
        toggleBold: {
          canExecute: ({ selection }) => selection.canToggleBold,
          execute: ({ selection }) => selection.toggleBold(),
        },
        toggleItalic: {
          canExecute: ({ selection }) => selection.canToggleItalic,
          execute: ({ selection }) => selection.toggleItalic(),
        },
        toggleUnderline: {
          canExecute: ({ selection }) => selection.canToggleUnderline,
          execute: ({ selection }) => selection.toggleUnderline(),
        },
      }),
    }),
    createOfficeEditorExtension<
      PresentationCommandContext,
      PresentationEditorCommands
    >({
      name: 'presentationReview',
      addCommands: () => ({
        addComment: {
          canExecute: ({ review }) => review.canAddComment,
          execute: ({ review }) => review.addComment(),
        },
        closeComments: {
          execute: ({ review }) => review.closeComments(),
        },
        deletePresentationComment: {
          execute: ({ review }, slideId, commentId) =>
            review.deleteComment(slideId, commentId),
        },
        locatePresentationComment: {
          execute: ({ review }, slideId, commentId) =>
            review.locateComment(slideId, commentId),
        },
        openComment: {
          execute: ({ review }, commentId) => review.openComment(commentId),
        },
        toggleComments: {
          execute: ({ review }) => review.toggleComments(),
        },
        updatePresentationComment: {
          execute: ({ review }, slideId, commentId, text) =>
            review.updateComment(slideId, commentId, text),
        },
      }),
    }),
    createOfficeEditorExtension<
      PresentationCommandContext,
      PresentationEditorCommands
    >({
      name: 'presentationView',
      addCommands: () => ({
        setViewMode: {
          execute: ({ view }, mode) => view.setViewMode(mode),
        },
        startSlideshow: {
          canExecute: ({ view }) => view.canStartSlideshow,
          execute: ({ view }, source) => view.startSlideshow(source),
        },
        toggleDesign: {
          execute: ({ view }) => view.toggleDesign(),
        },
      }),
    }),
    createPresentationKeyboardExtension(),
  ];
}
