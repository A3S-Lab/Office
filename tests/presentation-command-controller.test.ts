import { describe, expect, test } from '@rstest/core';
import {
  createPresentationCommandDispatcher,
  type PresentationCommandHandlers,
} from '../src/internal/features/work/editors/presentation-command-controller';

describe('presentation command controller', () => {
  test('routes typed element and view commands without UI label inference', () => {
    const calls: string[] = [];
    const dispatch = createPresentationCommandDispatcher(handlers(calls));

    dispatch({
      type: 'element.update',
      patch: { bold: true },
      restoreTextFocus: false,
    });
    dispatch({ type: 'element.align', alignment: 'center' });
    dispatch({ type: 'element.reorder', direction: 1 });
    dispatch({ type: 'view.set', mode: 'sorter' });

    expect(calls).toEqual([
      'element.update:bold:true:false',
      'element.align:center',
      'element.reorder:1',
      'view.set:sorter',
    ]);
  });

  test('routes slide, clipboard, and transition commands explicitly', () => {
    const calls: string[] = [];
    const dispatch = createPresentationCommandDispatcher(handlers(calls));

    dispatch({ type: 'slide.add' });
    dispatch({ type: 'clipboard.copy' });
    dispatch({
      type: 'transition.set',
      transition: {
        type: 'fade',
        speed: 'medium',
        advanceOnClick: true,
      },
    });
    dispatch({ type: 'transition.applyToAll' });

    expect(calls).toEqual([
      'slide.add',
      'clipboard.copy',
      'transition.set:fade',
      'transition.applyToAll',
    ]);
  });
});

function handlers(calls: string[]): PresentationCommandHandlers {
  return {
    addChart: () => calls.push('chart.add'),
    addComment: () => calls.push('comment.add'),
    addElement: (type) => calls.push(`element.add:${type}`),
    addSlide: () => calls.push('slide.add'),
    addTable: () => calls.push('table.add'),
    alignElement: (alignment) => calls.push(`element.align:${alignment}`),
    applyTransitionToAll: () => calls.push('transition.applyToAll'),
    copySelection: () => calls.push('clipboard.copy'),
    cutSelection: () => calls.push('clipboard.cut'),
    deleteSlide: () => calls.push('slide.delete'),
    duplicateSlide: () => calls.push('slide.duplicate'),
    pasteSelection: () => calls.push('clipboard.paste'),
    redo: () => calls.push('history.redo'),
    reorderElement: (direction) => calls.push(`element.reorder:${direction}`),
    requestImage: () => calls.push('image.request'),
    setBackground: (color) => calls.push(`background.set:${color}`),
    setTransition: (transition) =>
      calls.push(`transition.set:${transition?.type ?? 'none'}`),
    setViewMode: (mode) => calls.push(`view.set:${mode}`),
    startSlideshow: () => calls.push('slideshow.start'),
    toggleComments: () => calls.push('comments.toggle'),
    toggleDesign: () => calls.push('design.toggle'),
    undo: () => calls.push('history.undo'),
    updateElement: (patch, options) =>
      calls.push(
        `element.update:${Object.keys(patch).join(',')}:${String(
          patch.bold,
        )}:${String(options.restoreTextFocus)}`,
      ),
  };
}
