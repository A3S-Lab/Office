import { expect, test } from '@rstest/core';
import { restoreMarkdownEditingSurfaceFocus } from '../src/internal/features/work/editors/markdown-editor-focus';

test('focuses a newly mounted Markdown editing surface', async () => {
  const trigger = document.createElement('button');
  document.body.append(trigger);
  trigger.focus();
  let target: HTMLTextAreaElement | null = null;

  restoreMarkdownEditingSurfaceFocus(() => target);
  await nextAnimationFrame();
  target = document.createElement('textarea');
  document.body.append(target);
  await nextAnimationFrame();

  expect(target).toHaveFocus();
  trigger.remove();
  target.remove();
});

test('does not steal Markdown focus from an unrelated control', async () => {
  const trigger = document.createElement('button');
  const target = document.createElement('textarea');
  const unrelated = document.createElement('input');
  document.body.append(trigger, target, unrelated);
  trigger.focus();

  restoreMarkdownEditingSurfaceFocus(() => target);
  unrelated.focus();
  await nextAnimationFrame();

  expect(unrelated).toHaveFocus();
  trigger.remove();
  target.remove();
  unrelated.remove();
});

function nextAnimationFrame(): Promise<void> {
  return new Promise((resolve) => requestAnimationFrame(() => resolve()));
}
