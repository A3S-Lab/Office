import { expect, test } from '@rstest/core';
import { presentationCoreContextMenuItems } from '../src/internal/features/work/editors/presentation-context-menu';

test('provides native slide operations before optional AI actions', () => {
  const calls: string[] = [];
  const items = presentationCoreContextMenuItems({
    can: {
      copySelection: () => false,
      cutSelection: () => false,
      deleteSelection: () => false,
      deleteSlide: () => true,
      duplicateSelection: () => false,
      duplicateSlide: () => true,
    },
    commands: {
      copySelection: () => false,
      cutSelection: () => false,
      deleteSelection: () => false,
      deleteSlideById: (slideId) => {
        calls.push(`delete:${slideId}`);
        return true;
      },
      duplicateSelection: () => false,
      duplicateSlide: () => calls.push('duplicate'),
    },
    slideId: 'slide-2',
    target: 'slide',
  });

  expect(items.map(({ label }) => label)).toEqual(['复制幻灯片', '删除幻灯片']);
  items[0].onSelect();
  items[1].onSelect();
  expect(calls).toEqual(['duplicate', 'delete:slide-2']);
});

test('provides native object clipboard and delete operations', () => {
  const calls: string[] = [];
  const items = presentationCoreContextMenuItems({
    can: {
      copySelection: () => true,
      cutSelection: () => true,
      deleteSelection: () => true,
      deleteSlide: () => false,
      duplicateSelection: () => true,
      duplicateSlide: () => false,
    },
    commands: {
      copySelection: () => {
        calls.push('copy');
        return true;
      },
      cutSelection: () => {
        calls.push('cut');
        return true;
      },
      deleteSelection: () => {
        calls.push('delete');
        return true;
      },
      deleteSlideById: () => false,
      duplicateSelection: () => {
        calls.push('duplicate');
        return true;
      },
      duplicateSlide: () => undefined,
    },
    slideId: 'slide-1',
    target: 'element',
  });

  expect(items.map(({ label }) => label)).toEqual([
    '复制对象',
    '剪切对象',
    '创建副本',
    '删除对象',
  ]);
  for (const item of items) item.onSelect();
  expect(calls).toEqual(['copy', 'cut', 'duplicate', 'delete']);
});
