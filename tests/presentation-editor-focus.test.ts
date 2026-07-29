import { expect, test } from '@rstest/core';
import {
  presentationCommandsWithObjectFocus,
  restorePresentationObjectFocus,
  restorePresentationWorkspaceFocus,
} from '../src/internal/features/work/editors/presentation-editor-focus';
import type { PresentationEditorCommands } from '../src/internal/features/work/editors/presentation-command-types';

test('returns successful object-level ribbon commands to the slide selection', () => {
  const calls: string[] = [];
  const commands = {
    addSlide: () => calls.push('slide.add'),
    addChart: () => calls.push('add-chart'),
    addElement: () => calls.push('add-element'),
    addImage: async () => calls.push('add-image'),
    addTable: () => calls.push('add-table'),
    alignElement: () => calls.push('align'),
    applyTransitionToAll: (transition: { type?: string } | undefined) => {
      calls.push(`transition.apply:${transition?.type ?? 'none'}`);
      return Boolean(transition);
    },
    copySelection: () => true,
    cutSelection: () => true,
    deleteSlide: () => calls.push('slide.delete'),
    distributeElements: () => calls.push('distribute'),
    duplicateSlide: () => calls.push('slide.duplicate'),
    groupElements: () => true,
    pasteSelection: () => true,
    redo: () => false,
    reorderElement: () => calls.push('reorder'),
    setViewMode: () => calls.push('view'),
    setTransition: (transition: { type?: string } | undefined) => {
      calls.push(`transition.set:${transition?.type ?? 'none'}`);
      return Boolean(transition);
    },
    toggleBold: () => true,
    toggleItalic: () => false,
    toggleUnderline: () => true,
    undo: () => true,
    ungroupElements: () => true,
    updateElement: () => calls.push('update'),
  } as unknown as PresentationEditorCommands;
  const focusCalls: string[] = [];
  const ribbon = presentationCommandsWithObjectFocus(commands, () =>
    focusCalls.push('selection'),
  );

  ribbon.addSlide();
  ribbon.addChart();
  ribbon.addElement('shape');
  ribbon.addTable({ rows: 2, columns: 2 });
  ribbon.alignElement('left');
  expect(
    ribbon.applyTransitionToAll({
      type: 'fade',
      speed: 'medium',
      advanceOnClick: true,
    }),
  ).toBe(true);
  expect(
    ribbon.setTransition({
      type: 'fade',
      speed: 'medium',
      advanceOnClick: true,
    }),
  ).toBe(true);
  expect(ribbon.setTransition(undefined)).toBe(false);
  expect(ribbon.cutSelection()).toBe(true);
  ribbon.deleteSlide();
  ribbon.distributeElements('horizontal');
  ribbon.duplicateSlide();
  ribbon.reorderElement(1);
  ribbon.setViewMode('sorter');
  ribbon.updateElement({ bold: true });
  expect(ribbon.toggleBold()).toBe(true);
  expect(ribbon.toggleItalic()).toBe(false);
  expect(ribbon.redo()).toBe(false);

  expect(calls).toEqual([
    'slide.add',
    'add-chart',
    'add-element',
    'add-table',
    'align',
    'transition.apply:fade',
    'transition.set:fade',
    'transition.set:none',
    'slide.delete',
    'distribute',
    'slide.duplicate',
    'reorder',
    'view',
    'update',
  ]);
  expect(focusCalls).toEqual([
    'selection',
    'selection',
    'selection',
    'selection',
    'selection',
    'selection',
    'selection',
    'selection',
    'selection',
    'selection',
    'selection',
    'selection',
    'selection',
    'selection',
    'selection',
  ]);
});

test('focuses the active slide when presentation views remount', async () => {
  const trigger = document.createElement('button');
  const root = document.createElement('section');
  const thumbnail = document.createElement('button');
  thumbnail.dataset.slideThumbnail = 'true';
  thumbnail.dataset.slideId = 'slide-1';
  thumbnail.className = 'active';
  root.append(thumbnail);
  document.body.append(trigger, root);
  const state = {
    editingElementId: null as string | null,
    selectedElementIds: ['element-1'] as readonly string[],
    selectedSlideId: 'slide-1',
    viewMode: 'sorter' as 'normal' | 'sorter',
  };

  trigger.focus();
  restorePresentationWorkspaceFocus(root, () => state);
  await waitForAnimationFrames(2);
  expect(thumbnail).toHaveFocus();

  const object = presentationObject('element-1', true);
  root.replaceChildren(object, thumbnail);
  state.viewMode = 'normal';
  trigger.focus();
  restorePresentationWorkspaceFocus(root, () => state);
  await waitForAnimationFrames(2);
  expect(object).toHaveFocus();

  trigger.remove();
  root.remove();
});

test('focuses the latest selected object without stealing active editing focus', async () => {
  const trigger = document.createElement('button');
  const unrelated = document.createElement('input');
  const canvas = document.createElement('section');
  const first = presentationObject('element-1', true);
  const second = presentationObject('element-2', true);
  canvas.append(first, second);
  document.body.append(trigger, unrelated, canvas);
  const state = {
    editingElementId: null as string | null,
    selectedElementIds: ['element-1', 'element-2'] as readonly string[],
  };

  trigger.focus();
  restorePresentationObjectFocus(canvas, () => state);
  await waitForAnimationFrames(2);
  expect(second).toHaveFocus();

  trigger.focus();
  state.editingElementId = 'element-2';
  restorePresentationObjectFocus(canvas, () => state);
  await waitForAnimationFrames(2);
  expect(trigger).toHaveFocus();

  state.editingElementId = null;
  restorePresentationObjectFocus(canvas, () => state);
  unrelated.focus();
  await waitForAnimationFrames(2);
  expect(unrelated).toHaveFocus();

  trigger.remove();
  unrelated.remove();
  canvas.remove();
});

function presentationObject(id: string, selected: boolean): HTMLElement {
  const element = document.createElement('fieldset');
  element.tabIndex = 0;
  element.dataset.slideElementId = id;
  element.dataset.slideElementSelected = selected ? 'true' : 'false';
  return element;
}

function waitForAnimationFrames(count: number): Promise<void> {
  return new Promise((resolve) => {
    const next = (remaining: number) => {
      if (remaining <= 0) {
        resolve();
        return;
      }
      requestAnimationFrame(() => next(remaining - 1));
    };
    next(count);
  });
}
