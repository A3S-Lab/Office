import { expect, test } from '@rstest/core';
import { presentationCoreContextMenuItems } from '../src/internal/features/work/editors/presentation-context-menu';

test('provides native slide operations before optional AI actions', () => {
  const calls: string[] = [];
  const items = presentationCoreContextMenuItems({
    can: {
      addSlide: () => true,
      copySelection: () => false,
      cutSelection: () => false,
      deleteSelection: () => false,
      deleteSlide: () => true,
      duplicateSelection: () => false,
      duplicateSlide: () => true,
      pasteSelection: () => true,
    },
    commands: {
      addSlide: () => calls.push('add'),
      copySelection: () => false,
      cutSelection: () => false,
      deleteSelection: () => false,
      deleteSlideById: (slideId) => {
        calls.push(`delete:${slideId}`);
        return true;
      },
      duplicateSelection: () => false,
      duplicateSlide: () => calls.push('duplicate'),
      pasteSelection: () => {
        calls.push('paste');
        return true;
      },
    },
    slideId: 'slide-2',
    target: 'slide',
  });

  expect(items.map(({ label }) => label)).toEqual([
    '新建幻灯片',
    '复制幻灯片',
    '粘贴',
    '删除幻灯片',
  ]);
  expect(items[0]?.shortcut).toBe('Ctrl+M / ⌘⇧N');
  for (const item of items) item.onSelect();
  expect(calls).toEqual(['add', 'duplicate', 'paste', 'delete:slide-2']);
});

test('provides native object clipboard and delete operations', () => {
  const calls: string[] = [];
  const items = presentationCoreContextMenuItems({
    can: {
      addSlide: () => true,
      copySelection: () => true,
      cutSelection: () => true,
      deleteSelection: () => true,
      deleteSlide: () => false,
      duplicateSelection: () => true,
      duplicateSlide: () => false,
      pasteSelection: () => true,
    },
    commands: {
      addSlide: () => undefined,
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
      pasteSelection: () => {
        calls.push('paste');
        return true;
      },
    },
    slideId: 'slide-1',
    target: 'element',
  });

  expect(items.map(({ label }) => label)).toEqual([
    '复制对象',
    '剪切对象',
    '粘贴',
    '创建副本',
    '删除对象',
  ]);
  for (const item of items) item.onSelect();
  expect(calls).toEqual(['copy', 'cut', 'paste', 'duplicate', 'delete']);
});
