import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, within } from '@testing-library/react';
import { PresentationTransitionPanel } from '../src/internal/features/work/editors/presentation-transition-panel';
import type { WorkSlideTransition } from '../src/internal/features/work/work-types';

test('uses standard ribbon groups for transition effect, timing, and apply actions', () => {
  const changes: Array<WorkSlideTransition | undefined> = [];
  let applyCount = 0;
  render(
    <PresentationTransitionPanel
      transition={undefined}
      onChange={(transition) => changes.push(transition)}
      onApplyToAll={() => {
        applyCount += 1;
      }}
    />,
  );

  const effect = screen.getByRole('region', { name: '切换效果' });
  const timing = screen.getByRole('region', { name: '换片方式' });
  const apply = screen.getByRole('region', { name: '应用' });
  expect(
    within(effect).getByRole('combobox', { name: '幻灯片切换效果' }),
  ).toBeEnabled();
  expect(
    within(effect).getByRole('combobox', { name: '切换速度' }),
  ).toBeDisabled();
  expect(
    within(timing).getByRole('checkbox', { name: '自动换片' }),
  ).toBeDisabled();

  fireEvent.click(
    within(effect).getByRole('combobox', { name: '幻灯片切换效果' }),
  );
  fireEvent.click(screen.getByRole('option', { name: '淡化' }));
  fireEvent.click(
    within(apply).getByRole('button', {
      name: '应用切换效果到全部幻灯片',
    }),
  );

  expect(changes[0]).toMatchObject({ type: 'fade' });
  expect(applyCount).toBe(1);
});
