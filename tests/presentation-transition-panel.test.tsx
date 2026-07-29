import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, within } from '@testing-library/react';
import { useState } from 'react';
import { PresentationTransitionPanel } from '../src/internal/features/work/editors/presentation-transition-panel';
import { workSlideTransitionsEqual } from '../src/internal/features/work/work-presentation-transition';
import type { WorkSlideTransition } from '../src/internal/features/work/work-types';

const fadeTransition: WorkSlideTransition = {
  type: 'fade',
  speed: 'medium',
  advanceOnClick: true,
  advanceAfterMs: 5000,
};

test('uses truthful ribbon states for transition effect, timing, and apply actions', () => {
  const changes: Array<WorkSlideTransition | undefined> = [];
  const applied: Array<WorkSlideTransition | undefined> = [];
  render(
    <TransitionHarness
      initialTransition={undefined}
      canApplyToAll={(transition) => transition?.type === 'fade'}
      onChange={(transition) => changes.push(transition)}
      onApplyToAll={(transition) => applied.push(transition)}
    />,
  );

  const effect = screen.getByRole('region', { name: '切换效果' });
  const timing = screen.getByRole('region', { name: '换片方式' });
  const apply = screen.getByRole('region', { name: '应用' });
  const applyButton = within(apply).getByRole('button', {
    name: '应用切换效果到全部幻灯片',
  });
  expect(
    within(effect).getByRole('combobox', { name: '幻灯片切换效果' }),
  ).toBeEnabled();
  expect(
    within(effect).getByRole('combobox', { name: '切换速度' }),
  ).toBeDisabled();
  expect(
    within(timing).getByRole('checkbox', { name: '自动换片' }),
  ).toBeDisabled();
  expect(applyButton).toBeDisabled();

  fireEvent.click(
    within(effect).getByRole('combobox', { name: '幻灯片切换效果' }),
  );
  fireEvent.click(screen.getByRole('option', { name: '淡化' }));

  expect(changes[0]).toMatchObject({ type: 'fade' });
  expect(
    within(effect).getByRole('combobox', { name: '切换速度' }),
  ).toBeEnabled();
  expect(applyButton).toBeEnabled();
  fireEvent.click(applyButton);
  expect(applied).toEqual([
    {
      type: 'fade',
      speed: 'medium',
      advanceOnClick: true,
      advanceAfterMs: undefined,
    },
  ]);
});

test('commits automatic advance timing only after a complete valid draft', () => {
  const changes: Array<WorkSlideTransition | undefined> = [];
  render(
    <TransitionHarness
      initialTransition={fadeTransition}
      canApplyToAll={() => true}
      onChange={(transition) => changes.push(transition)}
    />,
  );

  const seconds = screen.getByRole('textbox', { name: '自动换片秒数' });
  fireEvent.change(seconds, { target: { value: '' } });
  expect(seconds).toHaveValue('');
  expect(changes).toEqual([]);

  fireEvent.change(seconds, { target: { value: '7.75' } });
  expect(changes).toEqual([]);
  fireEvent.blur(seconds);

  expect(seconds).toHaveValue('7.75');
  expect(changes).toEqual([{ ...fadeTransition, advanceAfterMs: 7750 }]);

  fireEvent.change(seconds, { target: { value: '12' } });
  fireEvent.keyDown(seconds, { key: 'Escape' });
  expect(seconds).toHaveValue('7.75');
  expect(changes).toHaveLength(1);
});

test('enables apply-to-all for a dirty timing draft and applies that exact draft', () => {
  const applied: Array<WorkSlideTransition | undefined> = [];
  const changes: Array<WorkSlideTransition | undefined> = [];
  render(
    <TransitionHarness
      initialTransition={fadeTransition}
      canApplyToAll={(transition) =>
        !workSlideTransitionsEqual(transition, fadeTransition)
      }
      onChange={(transition) => changes.push(transition)}
      onApplyToAll={(transition) => applied.push(transition)}
    />,
  );

  const apply = screen.getByRole('button', {
    name: '应用切换效果到全部幻灯片',
  });
  const seconds = screen.getByRole('textbox', { name: '自动换片秒数' });
  expect(apply).toBeDisabled();

  fireEvent.change(seconds, { target: { value: '7.75' } });
  expect(apply).toBeEnabled();
  fireEvent.click(apply);

  expect(changes).toEqual([]);
  expect(applied).toEqual([{ ...fadeTransition, advanceAfterMs: 7750 }]);
  expect(seconds).toHaveValue('7.75');
});

test('preserves a dirty timing draft across unrelated changes and resets it for another slide', () => {
  const props = {
    editable: true,
    canApplyToAll: () => true,
    onChange: () => undefined,
    onApplyToAll: () => undefined,
  };
  const { rerender } = render(
    <PresentationTransitionPanel
      {...props}
      slideId="slide-1"
      transition={fadeTransition}
    />,
  );
  const seconds = screen.getByRole('textbox', { name: '自动换片秒数' });
  fireEvent.change(seconds, { target: { value: '12' } });

  rerender(
    <PresentationTransitionPanel
      {...props}
      slideId="slide-1"
      transition={{ ...fadeTransition, speed: 'slow' }}
    />,
  );
  expect(seconds).toHaveValue('12');

  rerender(
    <PresentationTransitionPanel
      {...props}
      slideId="slide-2"
      transition={{ ...fadeTransition, speed: 'slow' }}
    />,
  );
  expect(seconds).toHaveValue('5');
});

test('disables every transition mutation outside slide editing', () => {
  render(
    <PresentationTransitionPanel
      slideId="slide-1"
      transition={fadeTransition}
      editable={false}
      canApplyToAll={() => true}
      onChange={() => undefined}
      onApplyToAll={() => undefined}
    />,
  );

  expect(
    screen.getByRole('combobox', { name: '幻灯片切换效果' }),
  ).toBeDisabled();
  expect(screen.getByRole('combobox', { name: '切换速度' })).toBeDisabled();
  expect(screen.getByRole('checkbox', { name: '自动换片' })).toBeDisabled();
  expect(screen.getByRole('textbox', { name: '自动换片秒数' })).toBeDisabled();
  expect(
    screen.getByRole('button', {
      name: '应用切换效果到全部幻灯片',
    }),
  ).toBeDisabled();
});

function TransitionHarness({
  initialTransition,
  canApplyToAll,
  onChange = () => undefined,
  onApplyToAll = () => undefined,
}: {
  initialTransition: WorkSlideTransition | undefined;
  canApplyToAll: (transition: WorkSlideTransition | undefined) => boolean;
  onChange?: (transition: WorkSlideTransition | undefined) => void;
  onApplyToAll?: (transition: WorkSlideTransition | undefined) => void;
}) {
  const [transition, setTransition] = useState(initialTransition);
  return (
    <PresentationTransitionPanel
      slideId="slide-1"
      transition={transition}
      editable
      canApplyToAll={canApplyToAll}
      onChange={(next) => {
        setTransition(next);
        onChange(next);
      }}
      onApplyToAll={onApplyToAll}
    />
  );
}
