import {
  ChevronLeft,
  ChevronRight,
  Maximize2,
  Presentation,
  X,
} from 'lucide-react';
import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
} from 'react';
import { slideTransitionDurationMilliseconds } from '../work-presentation-transition';
import type { WorkPresentationContent } from '../work-types';
import {
  createPresentationTimerController,
  PresentationPresenterView,
  type PresentationTimerController,
} from './presentation-presenter-view';
import { SlideCanvas } from './presentation-slide-canvas';

interface PlaybackState {
  index: number;
  transitionKey: number;
}

export function PresentationPlayer({
  autoFullscreen = false,
  content,
  initialIndex = 0,
  onExit,
}: {
  autoFullscreen?: boolean;
  content: WorkPresentationContent;
  initialIndex?: number;
  onExit?: () => void;
}) {
  const [playback, setPlayback] = useState<PlaybackState>({
    index: presentationPlaybackIndex(initialIndex, content.slides.length),
    transitionKey: 0,
  });
  const [presenter, setPresenter] = useState(false);
  const playerRef = useRef<HTMLDivElement>(null);
  const presenterTimerRef = useRef<PresentationTimerController | null>(null);
  if (!presenterTimerRef.current) {
    presenterTimerRef.current = createPresentationTimerController();
  }
  const enteredFullscreenRef = useRef(false);
  const completedExitRef = useRef(false);
  const slide = content.slides[playback.index] ?? content.slides[0];
  const completeExit = useCallback(() => {
    if (completedExitRef.current) return;
    completedExitRef.current = true;
    onExit?.();
  }, [onExit]);
  const enterFullscreen = useCallback(
    (element: HTMLElement) => {
      void requestPresentationFullscreen(element).then((enteredFullscreen) => {
        if (!enteredFullscreen) return;
        enteredFullscreenRef.current = true;
        if (document.fullscreenElement !== element) completeExit();
      });
    },
    [completeExit],
  );
  const move = useCallback(
    (delta: number) => {
      setPlayback((current) => {
        const index = Math.min(
          Math.max(current.index + delta, 0),
          content.slides.length - 1,
        );
        return index === current.index
          ? current
          : { index, transitionKey: current.transitionKey + 1 };
      });
    },
    [content.slides.length],
  );
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        if (document.fullscreenElement && document.exitFullscreen) {
          void document.exitFullscreen();
        } else {
          completeExit();
        }
        return;
      }

      if (presentationTargetAcceptsTextInput(event.target)) return;

      if (
        (event.key === ' ' || event.key === 'Spacebar') &&
        presentationTargetActivatesWithSpace(event.target)
      ) {
        return;
      }

      if (
        event.key === 'ArrowRight' ||
        event.key === 'ArrowDown' ||
        event.key === 'PageDown' ||
        event.key === ' ' ||
        event.key === 'Spacebar'
      ) {
        event.preventDefault();
        move(1);
        return;
      }
      if (
        event.key === 'ArrowLeft' ||
        event.key === 'ArrowUp' ||
        event.key === 'PageUp'
      ) {
        event.preventDefault();
        move(-1);
        return;
      }
      if (event.key === 'Home') {
        event.preventDefault();
        move(-content.slides.length);
        return;
      }
      if (event.key === 'End') {
        event.preventDefault();
        move(content.slides.length);
      }
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [completeExit, content.slides.length, move]);
  useLayoutEffect(() => {
    const player = playerRef.current;
    if (!autoFullscreen || !player) return;
    player.focus({ preventScroll: true });
    enterFullscreen(player);
  }, [autoFullscreen, enterFullscreen]);
  useEffect(() => {
    const onFullscreenChange = () => {
      if (document.fullscreenElement === playerRef.current) {
        enteredFullscreenRef.current = true;
        return;
      }
      if (enteredFullscreenRef.current) completeExit();
    };
    document.addEventListener('fullscreenchange', onFullscreenChange);
    onFullscreenChange();
    return () =>
      document.removeEventListener('fullscreenchange', onFullscreenChange);
  }, [completeExit]);
  useEffect(() => {
    if (!onExit) return;
    const timer = window.setInterval(() => {
      if (
        enteredFullscreenRef.current &&
        document.fullscreenElement !== playerRef.current
      ) {
        completeExit();
      }
    }, 100);
    return () => window.clearInterval(timer);
  }, [completeExit, onExit]);
  useEffect(() => {
    const delay = slide?.transition?.advanceAfterMs;
    if (delay === undefined || playback.index >= content.slides.length - 1)
      return;
    const timer = window.setTimeout(() => move(1), delay);
    return () => window.clearTimeout(timer);
  }, [
    content.slides.length,
    move,
    playback.index,
    slide?.transition?.advanceAfterMs,
  ]);
  if (!slide) return null;
  const aspectRatio = `${content.width ?? 13.333} / ${content.height ?? 7.5}`;
  const transition = slide.transition;
  const transitionStyle = {
    '--work-slide-transition-duration': `${slideTransitionDurationMilliseconds(transition)}ms`,
  } as React.CSSProperties;
  return (
    <section
      className="work-presentation-player"
      data-player-mode={presenter ? 'presenter' : 'audience'}
      ref={playerRef}
      tabIndex={-1}
    >
      {presenter ? (
        <PresentationPresenterView
          content={content}
          slide={slide}
          nextSlide={content.slides[playback.index + 1]}
          index={playback.index}
          total={content.slides.length}
          aspectRatio={aspectRatio}
          timer={presenterTimerRef.current}
        />
      ) : (
        <div className="work-presentation-player-stage">
          <button
            type="button"
            className="work-presentation-player-advance"
            aria-label="单击换到下一张幻灯片"
            disabled={
              transition?.advanceOnClick === false ||
              playback.index === content.slides.length - 1
            }
            onClick={() => move(1)}
          />
          <div
            aria-live="polite"
            className="work-presentation-transition-layer"
            data-slide-index={playback.index}
            data-slide-transition={transition?.type ?? 'none'}
            data-transition-direction={transition?.direction}
            data-transition-orientation={transition?.orientation}
            data-transition-speed={transition?.speed ?? 'medium'}
            key={`${slide.id}-${playback.transitionKey}`}
            style={transitionStyle}
          >
            <SlideCanvas
              content={content}
              slide={slide}
              interactive={false}
              aspectRatio={aspectRatio}
            />
          </div>
        </div>
      )}
      <footer>
        <button
          type="button"
          aria-label="上一张"
          aria-keyshortcuts="ArrowLeft ArrowUp PageUp"
          disabled={playback.index === 0}
          onClick={() => move(-1)}
        >
          <ChevronLeft size={18} />
        </button>
        <span>
          {playback.index + 1} / {content.slides.length}
        </span>
        <button
          type="button"
          aria-label="下一张"
          aria-keyshortcuts="ArrowRight ArrowDown PageDown Space"
          disabled={playback.index === content.slides.length - 1}
          onClick={() => move(1)}
        >
          <ChevronRight size={18} />
        </button>
        <button
          type="button"
          className={presenter ? 'active' : ''}
          aria-label={presenter ? '退出演讲者视图' : '演讲者视图'}
          onClick={() => setPresenter((current) => !current)}
        >
          <Presentation size={16} />
        </button>
        <button
          type="button"
          className="work-presentation-player-fullscreen"
          aria-label="全屏放映"
          onClick={() => {
            if (playerRef.current) enterFullscreen(playerRef.current);
          }}
        >
          <Maximize2 size={16} />
        </button>
        {onExit && (
          <button
            type="button"
            className="work-presentation-player-exit"
            aria-label="退出放映"
            aria-keyshortcuts="Escape"
            onClick={() => {
              if (document.fullscreenElement && document.exitFullscreen) {
                void document.exitFullscreen().finally(completeExit);
              } else {
                completeExit();
              }
            }}
          >
            <X size={16} />
          </button>
        )}
      </footer>
    </section>
  );
}

function presentationTargetAcceptsTextInput(
  target: EventTarget | null,
): boolean {
  return (
    target instanceof HTMLInputElement ||
    target instanceof HTMLTextAreaElement ||
    target instanceof HTMLSelectElement ||
    (target instanceof HTMLElement && target.isContentEditable)
  );
}

function presentationTargetActivatesWithSpace(
  target: EventTarget | null,
): boolean {
  return (
    target instanceof HTMLButtonElement || target instanceof HTMLAnchorElement
  );
}

function presentationPlaybackIndex(index: number, slideCount: number): number {
  if (slideCount <= 1) return 0;
  return Math.min(slideCount - 1, Math.max(0, Math.trunc(index)));
}

async function requestPresentationFullscreen(
  element: HTMLElement,
): Promise<boolean> {
  if (!element.requestFullscreen) return false;
  try {
    await element.requestFullscreen();
    return true;
  } catch {
    // The in-page slideshow remains usable when the browser denies fullscreen.
    return false;
  }
}
