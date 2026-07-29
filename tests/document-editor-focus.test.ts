import { expect, test } from '@rstest/core';
import { restoreDocumentEditorFocus } from '../src/internal/features/work/editors/document-editor-focus';

test('returns focus to a newly mounted document body', async () => {
  const trigger = document.createElement('button');
  document.body.append(trigger);
  trigger.focus();
  let target: HTMLElement | null = null;

  restoreDocumentEditorFocus(() => target);
  await nextAnimationFrame();
  target = document.createElement('div');
  target.tabIndex = 0;
  document.body.append(target);
  await nextAnimationFrame();

  expect(target).toHaveFocus();
  trigger.remove();
  target.remove();
});

test('does not steal document focus from an unrelated control', async () => {
  const trigger = document.createElement('button');
  const target = document.createElement('div');
  const unrelated = document.createElement('input');
  target.tabIndex = 0;
  document.body.append(trigger, target, unrelated);
  trigger.focus();

  restoreDocumentEditorFocus(() => target);
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
