import { expect, rstest, test } from '@rstest/core';
import {
  act,
  fireEvent,
  render,
  screen,
  waitFor,
} from '@testing-library/react';
import { PresentationEditor } from '../src/internal/features/work/editors/presentation-editor';
import { PresentationPlayer } from '../src/internal/features/work/editors/presentation-player';
import type { WorkPresentationContent } from '../src/internal/features/work/work-types';

test('starts the player at the requested slide and exposes an explicit exit', () => {
  let exits = 0;
  render(
    <PresentationPlayer
      content={presentation()}
      initialIndex={1}
      onExit={() => {
        exits += 1;
      }}
    />,
  );

  expect(screen.getByText('2 / 3')).toBeVisible();
  fireEvent.click(screen.getByRole('button', { name: '退出放映' }));
  expect(exits).toBe(1);
});

test('requests fullscreen automatically when a slideshow starts', () => {
  const calls: Element[] = [];
  const previous = HTMLElement.prototype.requestFullscreen;
  HTMLElement.prototype.requestFullscreen = function requestFullscreen() {
    calls.push(this);
    return Promise.resolve();
  };
  try {
    render(<PresentationPlayer autoFullscreen content={presentation()} />);
    expect(calls).toHaveLength(1);
    expect(calls[0]).toHaveClass('work-presentation-player');
  } finally {
    HTMLElement.prototype.requestFullscreen = previous;
  }
});

test('exits when fullscreen closes before the change listener observes entry', async () => {
  let fullscreenElement: Element | null = null;
  let exits = 0;
  const previousRequestFullscreen = HTMLElement.prototype.requestFullscreen;
  const previousFullscreenElement = Object.getOwnPropertyDescriptor(
    document,
    'fullscreenElement',
  );
  Object.defineProperty(document, 'fullscreenElement', {
    configurable: true,
    get: () => fullscreenElement,
  });
  HTMLElement.prototype.requestFullscreen = function requestFullscreen() {
    fullscreenElement = this;
    fullscreenElement = null;
    return Promise.resolve();
  };

  try {
    render(
      <PresentationPlayer
        autoFullscreen
        content={presentation()}
        onExit={() => {
          exits += 1;
        }}
      />,
    );
    await waitFor(() => expect(exits).toBe(1));
  } finally {
    HTMLElement.prototype.requestFullscreen = previousRequestFullscreen;
    if (previousFullscreenElement) {
      Object.defineProperty(
        document,
        'fullscreenElement',
        previousFullscreenElement,
      );
    } else {
      delete (document as Document & { fullscreenElement?: Element | null })
        .fullscreenElement;
    }
  }
});

test('exits when the browser omits the fullscreen change event', async () => {
  let fullscreenElement: Element | null = null;
  let exits = 0;
  const previousRequestFullscreen = HTMLElement.prototype.requestFullscreen;
  const previousFullscreenElement = Object.getOwnPropertyDescriptor(
    document,
    'fullscreenElement',
  );
  Object.defineProperty(document, 'fullscreenElement', {
    configurable: true,
    get: () => fullscreenElement,
  });
  HTMLElement.prototype.requestFullscreen = function requestFullscreen() {
    fullscreenElement = this;
    return Promise.resolve();
  };

  try {
    render(
      <PresentationPlayer
        autoFullscreen
        content={presentation()}
        onExit={() => {
          exits += 1;
        }}
      />,
    );
    await waitFor(() => expect(fullscreenElement).not.toBeNull());
    fullscreenElement = null;
    await waitFor(() => expect(exits).toBe(1));
  } finally {
    HTMLElement.prototype.requestFullscreen = previousRequestFullscreen;
    if (previousFullscreenElement) {
      Object.defineProperty(
        document,
        'fullscreenElement',
        previousFullscreenElement,
      );
    } else {
      delete (document as Document & { fullscreenElement?: Element | null })
        .fullscreenElement;
    }
  }
});

test('opens the built-in slideshow instead of switching to ordinary preview', () => {
  let starts = 0;
  render(
    <PresentationEditor
      content={presentation()}
      preview={false}
      onChange={() => undefined}
      onStartSlideshow={() => {
        starts += 1;
      }}
    />,
  );

  fireEvent.click(screen.getByRole('tab', { name: '幻灯片放映' }));
  fireEvent.click(screen.getByRole('button', { name: '从头开始放映' }));

  expect(starts).toBe(1);
  expect(screen.getByRole('dialog', { name: '幻灯片放映' })).toBeVisible();
  expect(screen.getByText('1 / 3')).toBeVisible();
  expect(screen.queryByText('只读预览')).toBeNull();
  fireEvent.click(screen.getByRole('button', { name: '退出放映' }));
  expect(screen.queryByRole('dialog', { name: '幻灯片放映' })).toBeNull();
});

test('returns focus to the slideshow launcher after playback exits', async () => {
  render(
    <PresentationEditor
      content={presentation()}
      preview={false}
      onChange={() => undefined}
    />,
  );

  fireEvent.click(screen.getByRole('tab', { name: '幻灯片放映' }));
  const launcher = screen.getByRole('button', {
    name: '从当前幻灯片放映',
  });
  document.body.focus();
  fireEvent.click(launcher);
  fireEvent.click(screen.getByRole('button', { name: '退出放映' }));

  await waitFor(() => expect(launcher).toHaveFocus());
});

test('keeps slideshow keyboard commands active after presenter controls receive focus', () => {
  let exits = 0;
  render(
    <PresentationPlayer
      content={presentation()}
      onExit={() => {
        exits += 1;
      }}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '演讲者视图' }));
  const presenterToggle = screen.getByRole('button', {
    name: '退出演讲者视图',
  });
  presenterToggle.focus();

  expect(screen.queryByRole('button', { name: '演讲者上一张' })).toBeNull();
  expect(screen.queryByRole('button', { name: '演讲者下一张' })).toBeNull();
  expect(screen.getAllByRole('button', { name: '上一张' })).toHaveLength(1);
  expect(screen.getAllByRole('button', { name: '下一张' })).toHaveLength(1);

  fireEvent.keyDown(presenterToggle, { key: 'ArrowRight' });
  expect(screen.getByText('2 / 3')).toBeVisible();
  expect(
    screen.getByRole('region', { name: '当前幻灯片' }).querySelector('h2'),
  ).toHaveTextContent('Two');

  fireEvent.keyDown(presenterToggle, { key: 'End' });
  expect(screen.getByText('3 / 3')).toBeVisible();
  fireEvent.keyDown(presenterToggle, { key: 'Home' });
  expect(screen.getByText('1 / 3')).toBeVisible();
  fireEvent.keyDown(presenterToggle, { key: 'Escape' });
  expect(exits).toBe(1);
});

test('keeps the presenter timer continuous for the slideshow session', () => {
  rstest.useFakeTimers();
  rstest.setSystemTime(new Date('2026-08-01T00:00:00.000Z'));
  const view = render(<PresentationPlayer content={presentation()} />);

  try {
    act(() => rstest.advanceTimersByTime(5_000));
    fireEvent.click(screen.getByRole('button', { name: '演讲者视图' }));
    expect(screen.getByText('00:05')).toBeVisible();

    fireEvent.click(screen.getByRole('button', { name: '暂停计时' }));
    act(() => rstest.advanceTimersByTime(3_000));
    expect(screen.getByText('00:05')).toBeVisible();

    fireEvent.click(screen.getByRole('button', { name: '退出演讲者视图' }));
    act(() => rstest.advanceTimersByTime(2_000));
    fireEvent.click(screen.getByRole('button', { name: '演讲者视图' }));
    expect(screen.getByText('00:05')).toBeVisible();

    fireEvent.click(screen.getByRole('button', { name: '继续计时' }));
    act(() => rstest.advanceTimersByTime(2_000));
    expect(screen.getByText('00:07')).toBeVisible();

    fireEvent.click(screen.getByRole('button', { name: '重置计时' }));
    expect(screen.getByText('00:00')).toBeVisible();
  } finally {
    view.unmount();
    rstest.useRealTimers();
  }
});

test('consumes ordered click animation cues before advancing the slide', () => {
  render(<PresentationPlayer content={animatedPresentation()} />);

  const first = screen
    .getByText('Animated one')
    .closest<HTMLElement>('[data-slide-preview-element-id="element-one"]');
  const second = screen
    .getByText('Animated two')
    .closest<HTMLElement>('[data-slide-preview-element-id="element-two"]');
  const advance = screen.getByRole('button', {
    name: '单击换到下一张幻灯片',
  });
  expect(first).toHaveAttribute('data-slide-animation-state', 'hidden');
  expect(second).toHaveAttribute('data-slide-animation-state', 'hidden');

  fireEvent.click(advance);
  expect(screen.getByText('1 / 2')).toBeVisible();
  expect(first).toHaveAttribute('data-slide-animation-state', 'playing');
  expect(first).toHaveAttribute(
    'data-slide-animation-effects',
    'fade fade-out',
  );
  expect(first?.style.animationName).toBe(
    'work-slide-animation-fade, work-slide-animation-fade-out',
  );
  expect(first?.style.animationDelay).toBe('0ms, 500ms');
  expect(second).toHaveAttribute('data-slide-animation-state', 'hidden');

  fireEvent.click(advance);
  expect(screen.getByText('1 / 2')).toBeVisible();
  expect(first).toHaveAttribute('data-slide-animation-state', 'hidden');
  expect(second).toHaveAttribute('data-slide-animation-state', 'playing');

  fireEvent.click(advance);
  expect(screen.getByText('2 / 2')).toBeVisible();
});

test('starts a leading automatic animation cue without a click', () => {
  const content = animatedPresentation();
  const firstAnimation = content.slides[0].animations?.[0];
  if (!firstAnimation) throw new Error('Expected the animated fixture cue.');
  firstAnimation.trigger = 'after-previous';
  render(<PresentationPlayer content={content} />);

  expect(
    screen
      .getByText('Animated one')
      .closest('[data-slide-preview-element-id="element-one"]'),
  ).toHaveAttribute('data-slide-animation-state', 'playing');
});

function presentation(): WorkPresentationContent {
  return {
    type: 'presentation',
    slides: [
      { id: 'slide-1', name: 'One', elements: [] },
      { id: 'slide-2', name: 'Two', elements: [] },
      { id: 'slide-3', name: 'Three', elements: [] },
    ],
  };
}

function animatedPresentation(): WorkPresentationContent {
  return {
    type: 'presentation',
    slides: [
      {
        id: 'animated-slide',
        name: 'Animated',
        elements: [
          {
            id: 'element-one',
            type: 'text',
            x: 10,
            y: 10,
            width: 30,
            height: 12,
            text: 'Animated one',
            fontSize: 20,
            color: '#111111',
            fill: 'transparent',
            bold: false,
            align: 'left',
          },
          {
            id: 'element-two',
            type: 'text',
            x: 10,
            y: 30,
            width: 30,
            height: 12,
            text: 'Animated two',
            fontSize: 20,
            color: '#111111',
            fill: 'transparent',
            bold: false,
            align: 'left',
          },
        ],
        animations: [
          {
            id: 'animation-one',
            elementId: 'element-one',
            effect: 'fade',
            trigger: 'on-click',
            durationMs: 500,
            delayMs: 0,
          },
          {
            id: 'animation-two',
            elementId: 'element-one',
            effect: 'fade-out',
            trigger: 'after-previous',
            durationMs: 400,
            delayMs: 0,
          },
          {
            id: 'animation-three',
            elementId: 'element-two',
            effect: 'fly-in',
            trigger: 'on-click',
            durationMs: 600,
            delayMs: 100,
            direction: 'right',
          },
        ],
      },
      { id: 'final-slide', name: 'Final', elements: [] },
    ],
  };
}
