import { describe, expect, test } from '@rstest/core';
import { createOfficeEditorRuntime } from '../src/internal/features/work/editors/office-editor-extension';
import { createPresentationEditorExtensions } from '../src/internal/features/work/editors/presentation-command-controller';
import type { PresentationCommandContext } from '../src/internal/features/work/editors/presentation-command-types';

describe('presentation editor extensions', () => {
  test('exposes typed element and view commands without UI label inference', () => {
    const calls: string[] = [];
    const editor = presentationEditor(calls);

    editor.commands.updateElement({ bold: true }, { restoreTextFocus: false });
    editor.commands.alignElement('center');
    editor.commands.distributeElements('horizontal');
    editor.commands.groupElements();
    editor.commands.ungroupElements();
    editor.commands.reorderElement(1);
    editor.commands.setPresentationContent({
      type: 'presentation',
      slides: [],
    });
    editor.commands.setViewMode('sorter');
    editor.commands.editDesign('layout');
    editor.commands.renamePresentationLayout('Executive');

    expect(editor.extensionNames).toEqual([
      'presentationDocument',
      'presentationHistory',
      'presentationClipboard',
      'presentationSlides',
      'presentationAnimations',
      'presentationElements',
      'presentationDesign',
      'presentationInsert',
      'presentationSelection',
      'presentationReview',
      'presentationView',
      'presentationKeyboardShortcuts',
    ]);
    expect(calls).toEqual([
      'element.update:bold:true:false',
      'element.align:center',
      'element.distribute:horizontal',
      'element.group',
      'element.ungroup',
      'element.reorder:1',
      'document.set:0',
      'view.set:sorter',
      'design.edit:layout',
      'design.layout.rename:Executive',
    ]);
  });

  test('routes slide, clipboard, and transition commands explicitly', () => {
    const calls: string[] = [];
    const { commands } = presentationEditor(calls);

    commands.addSlide();
    commands.addTable({ rows: 4, columns: 5 });
    commands.copySelection();
    commands.setTransition({
      type: 'fade',
      speed: 'medium',
      advanceOnClick: true,
    });
    commands.applyTransitionToAll({
      type: 'fade',
      speed: 'medium',
      advanceOnClick: true,
    });
    commands.setEntranceAnimation('fade');
    commands.updateEntranceAnimation({
      trigger: 'with-previous',
      durationMs: 750,
    });
    commands.moveEntranceAnimation(-1);
    commands.previewAnimations();
    commands.selectSlide('slide-2', true);
    commands.selectElement('element-2', true);
    commands.updateTextElement('element-2', { text: 'Updated' });
    commands.updateNotes('Speaker note');
    commands.locatePresentationComment('slide-2', 'comment-1');
    commands.updatePresentationComment('slide-2', 'comment-1', 'Reviewed');
    commands.deletePresentationComment('slide-2', 'comment-1');

    expect(calls).toEqual([
      'slide.add',
      'table.add:4x5',
      'clipboard.copy',
      'transition.set:fade',
      'transition.applyToAll:fade',
      'animation.set:fade',
      'animation.update:with-previous:750',
      'animation.move:-1',
      'animation.preview',
      'slide.select:slide-2:true',
      'selection.select:element-2:true',
      'element.text:element-2:Updated',
      'notes.update:Speaker note',
      'comment.locate:slide-2:comment-1',
      'comment.update:slide-2:comment-1:Reviewed',
      'comment.delete:slide-2:comment-1',
    ]);
  });

  test('uses extension capability checks for history commands', () => {
    const calls: string[] = [];
    const editor = presentationEditor(calls);

    expect(editor.can().undo()).toBe(true);
    expect(editor.can().redo()).toBe(false);
    editor.commands.undo();

    expect(calls).toEqual(['history.undo']);
  });

  test('evaluates transition capabilities against the proposed transition', () => {
    const calls: string[] = [];
    const context = presentationContext(calls);
    context.slides.canApplyTransitionToAll = (transition) =>
      transition?.type === 'push';
    context.slides.canSetTransition = false;
    const editor = createOfficeEditorRuntime(
      context,
      createPresentationEditorExtensions(),
    );
    const fade = {
      type: 'fade',
      speed: 'medium',
      advanceOnClick: true,
    } as const;
    const push = {
      ...fade,
      type: 'push',
      direction: 'left',
    } as const;

    expect(editor.can().applyTransitionToAll(fade)).toBe(false);
    expect(editor.can().applyTransitionToAll(push)).toBe(true);
    expect(editor.can().setTransition(fade)).toBe(false);
    expect(calls).toEqual([]);
  });

  test('routes presentation shortcuts through typed commands and can()', () => {
    const calls: string[] = [];
    const editor = presentationEditor(calls);
    const undo = new KeyboardEvent('keydown', {
      cancelable: true,
      key: 'z',
      metaKey: true,
    });

    expect(editor.handleKeyDown(undo)).toBe(true);
    expect(undo.defaultPrevented).toBe(true);
    expect(calls).toEqual(['history.undo']);

    const slideshow = new KeyboardEvent('keydown', {
      cancelable: true,
      key: 'F5',
    });
    expect(editor.handleKeyDown(slideshow)).toBe(true);
    expect(slideshow.defaultPrevented).toBe(true);
    expect(calls).toEqual(['history.undo', 'slideshow.start:beginning']);

    const currentSlideshow = new KeyboardEvent('keydown', {
      cancelable: true,
      key: 'F5',
      shiftKey: true,
    });
    expect(editor.handleKeyDown(currentSlideshow)).toBe(true);
    expect(currentSlideshow.defaultPrevented).toBe(true);
    expect(calls).toEqual([
      'history.undo',
      'slideshow.start:beginning',
      'slideshow.start:current',
    ]);
  });

  test('does not create slides from toolbar shortcuts when slide editing is unavailable', () => {
    const calls: string[] = [];
    const context = presentationContext(calls);
    context.slides.canAddSlide = false;
    const editor = createOfficeEditorRuntime(
      context,
      createPresentationEditorExtensions(),
    );

    expect(editor.can().addSlide()).toBe(false);
    for (const event of [
      new KeyboardEvent('keydown', {
        cancelable: true,
        ctrlKey: true,
        key: 'm',
      }),
      new KeyboardEvent('keydown', {
        cancelable: true,
        key: 'n',
        metaKey: true,
        shiftKey: true,
      }),
    ]) {
      expect(editor.handleKeyDown(event)).toBe(false);
      expect(event.defaultPrevented).toBe(false);
    }
    expect(calls).not.toContain('slide.add');
  });

  test('honors the advertised text-formatting shortcuts on slide objects', () => {
    const calls: string[] = [];
    const commandContext = presentationContext(calls);
    commandContext.keyboard = {
      editingElementId: null,
      selectedElement: {
        id: 'title',
        type: 'text',
        x: 10,
        y: 10,
        width: 80,
        height: 20,
        text: 'Quarterly plan',
        fontSize: 24,
        color: '#172033',
        fill: 'transparent',
        bold: false,
        align: 'left',
      },
      selectedElementCount: 1,
    };
    const editor = createOfficeEditorRuntime(
      commandContext,
      createPresentationEditorExtensions(),
    );
    const object = document.createElement('button');
    object.dataset.slideElementOrigin = 'slide';
    document.body.append(object);
    try {
      const handled: boolean[] = [];
      object.addEventListener('keydown', (event) => {
        handled.push(editor.handleKeyDown(event));
      });
      for (const key of ['b', 'i', 'u']) {
        object.dispatchEvent(
          new KeyboardEvent('keydown', {
            bubbles: true,
            cancelable: true,
            key,
            metaKey: true,
          }),
        );
      }

      expect(handled).toEqual([true, true, true]);
      expect(calls).toEqual([
        'selection.bold',
        'selection.italic',
        'selection.underline',
      ]);
    } finally {
      object.remove();
    }
  });
});

function presentationEditor(calls: string[]) {
  return createOfficeEditorRuntime(
    presentationContext(calls),
    createPresentationEditorExtensions(),
  );
}

function presentationContext(calls: string[]): PresentationCommandContext {
  return {
    animations: {
      canMoveEntranceAnimation: () => true,
      canPreviewAnimations: true,
      canSetEntranceAnimation: true,
      canUpdateEntranceAnimation: true,
      moveEntranceAnimation: (direction) =>
        record(calls, `animation.move:${direction}`),
      previewAnimations: () => calls.push('animation.preview'),
      setEntranceAnimation: (effect) =>
        record(calls, `animation.set:${effect ?? 'none'}`),
      updateEntranceAnimation: (patch) =>
        record(
          calls,
          `animation.update:${patch.trigger ?? 'same'}:${patch.durationMs ?? 'same'}`,
        ),
    },
    clipboard: {
      canCopySelection: true,
      canCutSelection: true,
      canPasteSelection: true,
      copySelection: () => record(calls, 'clipboard.copy'),
      cutSelection: () => record(calls, 'clipboard.cut'),
      pasteSelection: () => record(calls, 'clipboard.paste'),
    },
    design: {
      addPlaceholder: (type) => calls.push(`design.placeholder:${type}`),
      applyLayout: (layoutId) => calls.push(`design.layout.apply:${layoutId}`),
      canDeleteLayout: true,
      close: () => calls.push('design.close'),
      createLayout: (copyCurrent) =>
        calls.push(`design.layout.create:${copyCurrent}`),
      deleteLayout: () => calls.push('design.layout.delete'),
      edit: (mode) => calls.push(`design.edit:${mode}`),
      renameLayout: (name) => calls.push(`design.layout.rename:${name}`),
      renameMaster: (name) => calls.push(`design.master.rename:${name}`),
      setLayoutBackground: (color) =>
        calls.push(`design.layout.background:${color ?? 'inherited'}`),
      setMasterBackground: (color) =>
        calls.push(`design.master.background:${color}`),
      toggleLayoutBackground: (enabled) =>
        calls.push(`design.layout.inherit:${enabled}`),
    },
    document: {
      setContent: (content) =>
        calls.push(`document.set:${content.slides.length}`),
    },
    elements: {
      canAlignElement: true,
      canDistributeElements: true,
      canGroupElements: true,
      canReorderElement: true,
      canUngroupElements: true,
      canUpdateElement: true,
      alignElement: (alignment) => calls.push(`element.align:${alignment}`),
      distributeElements: (direction) =>
        calls.push(`element.distribute:${direction}`),
      groupElements: () => record(calls, 'element.group'),
      reorderElement: (direction) => calls.push(`element.reorder:${direction}`),
      ungroupElements: () => record(calls, 'element.ungroup'),
      updateElement: (patch, options) =>
        calls.push(
          `element.update:${Object.keys(patch).join(',')}:${String(
            patch.bold,
          )}:${String(options.restoreTextFocus)}`,
        ),
      updateTextElement: (elementId, value) =>
        calls.push(`element.text:${elementId}:${value.text}`),
    },
    history: {
      canRedo: false,
      canUndo: true,
      redo: () => record(calls, 'history.redo'),
      undo: () => record(calls, 'history.undo'),
    },
    insert: {
      enabled: true,
      addChart: () => calls.push('chart.add'),
      addElement: (type) => calls.push(`element.add:${type}`),
      addImage: (file) => calls.push(`image.add:${file.name}`),
      addTable: ({ rows, columns }) =>
        calls.push(`table.add:${rows}x${columns}`),
      instantiatePlaceholder: (definition) =>
        calls.push(`placeholder.add:${definition.id}`),
      requestImage: () => calls.push('image.request'),
    },
    keyboard: {
      editingElementId: null,
      selectedElement: null,
      selectedElementCount: 0,
    },
    review: {
      canAddComment: true,
      addComment: () => calls.push('comment.add'),
      closeComments: () => calls.push('comments.close'),
      deleteComment: (slideId, commentId) =>
        calls.push(`comment.delete:${slideId}:${commentId}`),
      locateComment: (slideId, commentId) =>
        calls.push(`comment.locate:${slideId}:${commentId}`),
      openComment: (commentId) => calls.push(`comment.open:${commentId}`),
      toggleComments: () => calls.push('comments.toggle'),
      updateComment: (slideId, commentId, text) =>
        calls.push(`comment.update:${slideId}:${commentId}:${text}`),
    },
    selection: {
      canDeleteSelection: true,
      canDuplicateSelection: true,
      canEditElement: () => true,
      canExitEditing: true,
      canNudgeSelection: true,
      canToggleBold: true,
      canToggleItalic: true,
      canToggleUnderline: true,
      deleteSelection: () => record(calls, 'selection.delete'),
      duplicateSelection: () => record(calls, 'selection.duplicate'),
      editElement: (id) => calls.push(`selection.edit:${id}`),
      exitEditing: () => calls.push('selection.exit'),
      nudgeSelection: (key, distance) =>
        record(calls, `selection.nudge:${key}:${distance}`),
      selectElement: (elementId, additive) =>
        calls.push(`selection.select:${elementId ?? 'none'}:${additive}`),
      selectElements: (ids) => calls.push(`selection.set:${ids.join(',')}`),
      toggleBold: () => record(calls, 'selection.bold'),
      toggleItalic: () => record(calls, 'selection.italic'),
      toggleUnderline: () => record(calls, 'selection.underline'),
    },
    slides: {
      canAddSlide: true,
      canApplyTransitionToAll: () => true,
      canDeleteSlide: true,
      canDuplicateSlide: true,
      canSetTransition: true,
      addSlide: () => calls.push('slide.add'),
      applyTransitionToAll: (transition) => {
        calls.push(`transition.applyToAll:${transition?.type ?? 'none'}`);
        return true;
      },
      deleteSlide: () => calls.push('slide.delete'),
      deleteSlideById: (slideId) => record(calls, `slide.delete:${slideId}`),
      duplicateSlide: () => calls.push('slide.duplicate'),
      selectSlide: (slideId, returnToSlideMode) =>
        calls.push(`slide.select:${slideId}:${returnToSlideMode}`),
      setBackground: (color) => calls.push(`background.set:${color}`),
      setTransition: (transition) => {
        calls.push(`transition.set:${transition?.type ?? 'none'}`);
        return true;
      },
      updateNotes: (notes) => calls.push(`notes.update:${notes}`),
    },
    view: {
      canStartSlideshow: true,
      setViewMode: (mode) => calls.push(`view.set:${mode}`),
      startSlideshow: (source) => calls.push(`slideshow.start:${source}`),
      toggleDesign: () => calls.push('design.toggle'),
    },
  };
}

function record(calls: string[], value: string): boolean {
  calls.push(value);
  return true;
}
