import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { useState } from 'react';
import { PresentationAnimationPanel } from '../src/internal/features/work/editors/presentation-animation-panel';
import { createWorkSlideAnimation } from '../src/internal/features/work/work-presentation-animation';
import type {
  WorkSlideAnimation,
  WorkSlideAnimationEffect,
} from '../src/internal/features/work/work-types';

test('authors a bounded entrance effect, trigger, direction, and timing', () => {
  const updates: Array<Partial<WorkSlideAnimation>> = [];
  render(<AnimationHarness onUpdate={(patch) => updates.push(patch)} />);

  expect(
    screen.getByRole('combobox', { name: '对象入场动画触发方式' }),
  ).toBeDisabled();
  fireEvent.click(screen.getByRole('combobox', { name: '对象入场动画效果' }));
  fireEvent.click(screen.getByRole('option', { name: '飞入' }));

  fireEvent.click(
    screen.getByRole('combobox', { name: '对象入场动画触发方式' }),
  );
  fireEvent.click(screen.getByRole('option', { name: '与上一动画同时' }));
  fireEvent.click(screen.getByRole('combobox', { name: '对象飞入方向' }));
  fireEvent.click(screen.getByRole('option', { name: '从右侧' }));
  const duration = screen.getByRole('textbox', {
    name: '对象入场动画持续秒数',
  });
  fireEvent.change(duration, { target: { value: '1.25' } });
  expect(updates).toEqual([
    { trigger: 'with-previous' },
    { direction: 'right' },
  ]);
  fireEvent.blur(duration);

  expect(duration).toHaveValue('1.3');
  expect(updates.at(-1)).toEqual({ durationMs: 1300 });
});

test('exposes truthful ordering and preview capabilities', () => {
  const moves: number[] = [];
  let previews = 0;
  const animation = createWorkSlideAnimation('element', 'fade');
  render(
    <PresentationAnimationPanel
      animation={animation}
      canMove={(direction) => direction === 1}
      canPreview
      editable
      onMove={(direction) => moves.push(direction)}
      onPreview={() => {
        previews += 1;
      }}
      onSetEffect={() => undefined}
      onUpdate={() => undefined}
    />,
  );

  expect(screen.getByRole('button', { name: '提前对象动画' })).toBeDisabled();
  fireEvent.click(screen.getByRole('button', { name: '推后对象动画' }));
  fireEvent.click(screen.getByRole('button', { name: '预览当前幻灯片动画' }));
  expect(moves).toEqual([1]);
  expect(previews).toBe(1);
});

function AnimationHarness({
  onUpdate,
}: {
  onUpdate: (patch: Partial<WorkSlideAnimation>) => void;
}) {
  const [animation, setAnimation] = useState<WorkSlideAnimation>();
  const setEffect = (effect: WorkSlideAnimationEffect | undefined) => {
    setAnimation((current) =>
      effect ? createWorkSlideAnimation('element', effect, current) : undefined,
    );
  };
  const update = (patch: Partial<WorkSlideAnimation>) => {
    setAnimation((current) => (current ? { ...current, ...patch } : current));
    onUpdate(patch);
  };
  return (
    <PresentationAnimationPanel
      animation={animation}
      canMove={() => false}
      canPreview={Boolean(animation)}
      editable
      onMove={() => undefined}
      onPreview={() => undefined}
      onSetEffect={setEffect}
      onUpdate={update}
    />
  );
}
