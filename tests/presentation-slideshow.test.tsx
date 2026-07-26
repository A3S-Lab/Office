import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
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
