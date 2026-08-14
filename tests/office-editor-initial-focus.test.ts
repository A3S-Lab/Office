import { expect, test } from '@rstest/core';
import { stabilizeOfficeEditorInitialFocus } from '../src/internal/features/work/editors/office-editor-focus-handoff';

test('waits for an editable target and stabilizes focus across a remount', async () => {
  const opener = document.createElement('button');
  const firstTarget = document.createElement('div');
  firstTarget.tabIndex = 0;
  firstTarget.setAttribute('contenteditable', 'false');
  document.body.append(opener, firstTarget);
  opener.focus();
  let target = firstTarget;

  const stop = stabilizeOfficeEditorInitialFocus(
    () => target,
    opener,
    (candidate) => candidate.getAttribute('contenteditable') === 'true',
  );
  await waitForAnimationFrames(2);
  expect(opener).toHaveFocus();

  firstTarget.setAttribute('contenteditable', 'true');
  await waitForAnimationFrames(2);
  expect(firstTarget).toHaveFocus();

  firstTarget.blur();
  await waitForAnimationFrames(2);
  expect(firstTarget).toHaveFocus();

  const replacement = document.createElement('div');
  replacement.tabIndex = 0;
  replacement.setAttribute('contenteditable', 'true');
  firstTarget.replaceWith(replacement);
  target = replacement;
  await waitForAnimationFrames(2);

  expect(replacement).toHaveFocus();
  stop();
  opener.remove();
  replacement.remove();
});

test('does not steal focus after deliberate navigation', async () => {
  const opener = document.createElement('button');
  const target = document.createElement('textarea');
  const unrelated = document.createElement('input');
  target.readOnly = true;
  document.body.append(opener, target, unrelated);
  opener.focus();

  const stop = stabilizeOfficeEditorInitialFocus(
    () => target,
    opener,
    (candidate) =>
      candidate instanceof HTMLTextAreaElement && !candidate.readOnly,
  );
  unrelated.focus();
  target.readOnly = false;
  await waitForAnimationFrames(2);

  expect(unrelated).toHaveFocus();
  stop();
  opener.remove();
  target.remove();
  unrelated.remove();
});

test('cancels pending focus when the user presses Tab', async () => {
  const opener = document.createElement('button');
  const target = document.createElement('textarea');
  target.readOnly = true;
  document.body.append(opener, target);
  opener.focus();

  const stop = stabilizeOfficeEditorInitialFocus(
    () => target,
    opener,
    (candidate) =>
      candidate instanceof HTMLTextAreaElement && !candidate.readOnly,
  );
  document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab' }));
  target.readOnly = false;
  await waitForAnimationFrames(2);

  expect(opener).toHaveFocus();
  stop();
  opener.remove();
  target.remove();
});

test('cancels pending focus after an outside pointer action', async () => {
  const opener = document.createElement('button');
  const target = document.createElement('textarea');
  const outside = document.createElement('div');
  target.readOnly = true;
  document.body.append(opener, target, outside);
  opener.focus();

  const stop = stabilizeOfficeEditorInitialFocus(
    () => target,
    opener,
    (candidate) =>
      candidate instanceof HTMLTextAreaElement && !candidate.readOnly,
  );
  outside.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true }));
  target.readOnly = false;
  await waitForAnimationFrames(2);

  expect(opener).toHaveFocus();
  stop();
  opener.remove();
  target.remove();
  outside.remove();
});

test('cancels stabilization after an outside synthetic click', async () => {
  const opener = document.createElement('button');
  const target = document.createElement('div');
  const outside = document.createElement('button');
  target.tabIndex = 0;
  document.body.append(opener, target, outside);
  opener.focus();

  const stop = stabilizeOfficeEditorInitialFocus(() => target, opener);
  await waitForAnimationFrames(2);
  expect(target).toHaveFocus();

  outside.dispatchEvent(new MouseEvent('click', { bubbles: true }));
  target.blur();
  await waitForAnimationFrames(2);

  expect(target).not.toHaveFocus();
  stop();
  opener.remove();
  target.remove();
  outside.remove();
});

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
