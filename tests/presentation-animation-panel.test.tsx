import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { useState } from 'react';
import { PresentationAnimationPanel } from '../src/internal/features/work/editors/presentation-animation-panel';
import { createWorkSlideAnimation } from '../src/internal/features/work/work-presentation-animation';
import type {
  WorkSlideAnimation,
  WorkSlideAnimationClass,
  WorkSlideAnimationEffect,
} from '../src/internal/features/work/work-types';

test('authors bounded entrance and exit effects through one WPS-style panel', () => {
  const updates: Array<[WorkSlideAnimationClass, Partial<WorkSlideAnimation>]> =
    [];
  render(
    <AnimationHarness
      onUpdate={(animationClass, patch) =>
        updates.push([animationClass, patch])
      }
    />,
  );

  expect(
    screen.getByRole('combobox', { name: '对象动画触发方式' }),
  ).toBeDisabled();
  expect(
    screen.getByRole('combobox', { name: '对象动画类型' }),
  ).toHaveTextContent('进入');
  fireEvent.click(screen.getByRole('combobox', { name: '对象动画效果' }));
  fireEvent.click(screen.getByRole('option', { name: '飞入' }));

  fireEvent.click(screen.getByRole('combobox', { name: '对象动画触发方式' }));
  fireEvent.click(screen.getByRole('option', { name: '与上一动画同时' }));
  fireEvent.click(screen.getByRole('combobox', { name: '对象动画方向' }));
  fireEvent.click(screen.getByRole('option', { name: '从右侧' }));
  const duration = screen.getByRole('textbox', {
    name: '对象动画持续秒数',
  });
  fireEvent.change(duration, { target: { value: '1.25' } });
  expect(updates).toEqual([
    ['entrance', { trigger: 'with-previous' }],
    ['entrance', { direction: 'right' }],
  ]);
  fireEvent.blur(duration);

  expect(duration).toHaveValue('1.3');
  expect(updates.at(-1)).toEqual(['entrance', { durationMs: 1300 }]);

  fireEvent.click(screen.getByRole('combobox', { name: '对象动画类型' }));
  fireEvent.click(screen.getByRole('option', { name: '退出' }));
  fireEvent.click(screen.getByRole('combobox', { name: '对象动画效果' }));
  fireEvent.click(screen.getByRole('option', { name: '飞出' }));
  fireEvent.click(screen.getByRole('combobox', { name: '对象动画方向' }));
  fireEvent.click(screen.getByRole('option', { name: '向下方' }));
  expect(updates.at(-1)).toEqual(['exit', { direction: 'down' }]);
});

test('exposes truthful ordering and preview capabilities', () => {
  const moves: number[] = [];
  let previews = 0;
  const animation = createWorkSlideAnimation('element', 'fade');
  render(
    <PresentationAnimationPanel
      animations={{ entrance: animation, exit: undefined }}
      canMove={(_animationClass, direction) => direction === 1}
      canPreview
      editable
      elementId="element"
      onMove={(_animationClass, direction) => moves.push(direction)}
      onPreview={() => {
        previews += 1;
      }}
      onSetEffect={() => undefined}
      onUpdate={() => undefined}
      canUpdate={() => true}
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
  onUpdate: (
    animationClass: WorkSlideAnimationClass,
    patch: Partial<WorkSlideAnimation>,
  ) => void;
}) {
  const [animations, setAnimations] = useState<{
    entrance?: WorkSlideAnimation;
    exit?: WorkSlideAnimation;
  }>({});
  const setEffect = (
    animationClass: WorkSlideAnimationClass,
    effect: WorkSlideAnimationEffect | undefined,
  ) => {
    setAnimations((current) => ({
      ...current,
      [animationClass]: effect
        ? createWorkSlideAnimation('element', effect, current[animationClass])
        : undefined,
    }));
  };
  const update = (
    animationClass: WorkSlideAnimationClass,
    patch: Partial<WorkSlideAnimation>,
  ) => {
    setAnimations((current) => ({
      ...current,
      [animationClass]: current[animationClass]
        ? { ...current[animationClass], ...patch }
        : undefined,
    }));
    onUpdate(animationClass, patch);
  };
  return (
    <PresentationAnimationPanel
      animations={animations}
      canMove={() => false}
      canPreview={Boolean(animations.entrance || animations.exit)}
      canUpdate={() => true}
      editable
      elementId="element"
      onMove={() => undefined}
      onPreview={() => undefined}
      onSetEffect={setEffect}
      onUpdate={update}
    />
  );
}
