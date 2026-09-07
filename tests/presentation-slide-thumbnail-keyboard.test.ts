import { expect, test } from '@rstest/core';
import { handlePresentationThumbnailKey } from '../src/internal/features/work/editors/presentation-slide-thumbnail-keyboard';

test('Enter activates the selected sorter slide like WPS slide sorter', () => {
  const calls: string[] = [];
  const event = keyboardEvent('Enter');
  expect(
    handlePresentationThumbnailKey(event, {
      index: 1,
      slideCount: 3,
      onDelete: () => {
        calls.push('delete');
        return true;
      },
      onNavigate: (index) => calls.push(`nav:${index}`),
      onActivate: () => calls.push('activate'),
    }),
  ).toBe(true);
  expect(event.prevented).toBe(true);
  expect(calls).toEqual(['activate']);
});

test('Enter is ignored in the strip when no activate handler is provided', () => {
  const calls: string[] = [];
  expect(
    handlePresentationThumbnailKey(keyboardEvent('Enter'), {
      index: 0,
      slideCount: 2,
      onDelete: () => {
        calls.push('delete');
        return true;
      },
      onNavigate: (index) => calls.push(`nav:${index}`),
    }),
  ).toBe(false);
  expect(calls).toEqual([]);
});

test('arrow and delete keys keep the existing thumbnail navigation contract', () => {
  const calls: string[] = [];
  expect(
    handlePresentationThumbnailKey(keyboardEvent('ArrowRight'), {
      index: 0,
      slideCount: 3,
      onDelete: () => false,
      onNavigate: (index) => calls.push(`nav:${index}`),
    }),
  ).toBe(true);
  expect(
    handlePresentationThumbnailKey(keyboardEvent('End'), {
      index: 0,
      slideCount: 3,
      onDelete: () => false,
      onNavigate: (index) => calls.push(`nav:${index}`),
    }),
  ).toBe(true);
  expect(
    handlePresentationThumbnailKey(keyboardEvent('Delete'), {
      index: 1,
      slideCount: 3,
      onDelete: () => {
        calls.push('delete');
        return true;
      },
      onNavigate: (index) => calls.push(`nav:${index}`),
    }),
  ).toBe(true);
  expect(calls).toEqual(['nav:1', 'nav:2', 'delete']);
});

function keyboardEvent(key: string) {
  const state = { prevented: false, stopped: false };
  return {
    key,
    prevented: false as boolean,
    preventDefault() {
      state.prevented = true;
      this.prevented = true;
    },
    stopPropagation() {
      state.stopped = true;
    },
  };
}
