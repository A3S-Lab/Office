import { Pause, Play, RotateCcw } from 'lucide-react';
import { useEffect, useState } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import type { WorkPresentationContent, WorkSlide } from '../work-types';
import { useOfficeMessages } from './office-messages-context';
import { SlideCanvas } from './presentation-slide-canvas';

export interface PresentationTimerController {
  elapsedMilliseconds: number;
  runningSinceMilliseconds: number | null;
}

export function createPresentationTimerController(
  now = Date.now(),
): PresentationTimerController {
  return {
    elapsedMilliseconds: 0,
    runningSinceMilliseconds: now,
  };
}

export function PresentationPresenterView({
  animationCueIndex,
  content,
  slide,
  nextSlide,
  index,
  total,
  aspectRatio,
  timer,
}: {
  animationCueIndex?: number;
  content: WorkPresentationContent;
  slide: WorkSlide;
  nextSlide?: WorkSlide;
  index: number;
  total: number;
  aspectRatio: string;
  timer: PresentationTimerController;
}) {
  const messages = useOfficeMessages();
  const [, setTimerRevision] = useState(0);
  const now = Date.now();
  const running = timer.runningSinceMilliseconds !== null;
  const elapsedSeconds = presentationTimerElapsedSeconds(timer, now);

  useEffect(() => {
    if (!running) return;
    const interval = window.setInterval(
      () => setTimerRevision((current) => current + 1),
      1000,
    );
    return () => window.clearInterval(interval);
  }, [running]);

  const toggleTimer = () => {
    const currentTime = Date.now();
    if (timer.runningSinceMilliseconds === null) {
      timer.runningSinceMilliseconds = currentTime;
    } else {
      timer.elapsedMilliseconds = presentationTimerElapsedMilliseconds(
        timer,
        currentTime,
      );
      timer.runningSinceMilliseconds = null;
    }
    setTimerRevision((current) => current + 1);
  };

  const resetTimer = () => {
    const currentTime = Date.now();
    timer.elapsedMilliseconds = 0;
    if (timer.runningSinceMilliseconds !== null) {
      timer.runningSinceMilliseconds = currentTime;
    }
    setTimerRevision((current) => current + 1);
  };

  return (
    <section className="work-presentation-presenter" aria-label={officeMessage(messages, 'presentation.presenter.aria')}>
      <header>
        <div>
          <span>{officeMessage(messages, 'presentation.presenter.timer')}</span>
          <strong>
            <span className="sr-only">{officeMessage(messages, 'presentation.presenter.elapsedPrefix')}</span>
            <time dateTime={`PT${elapsedSeconds}S`}>
              {formatDuration(elapsedSeconds)}
            </time>
          </strong>
        </div>
        <div className="work-presentation-presenter-timer-actions">
          <button
            type="button"
            aria-label={running ? officeMessage(messages, 'presentation.presenter.pause') : officeMessage(messages, 'presentation.presenter.resume')}
            aria-pressed={!running}
            onClick={toggleTimer}
          >
            {running ? <Pause size={15} /> : <Play size={15} />}
          </button>
          <button type="button" aria-label={officeMessage(messages, 'presentation.presenter.reset')} onClick={resetTimer}>
            <RotateCcw size={15} />
          </button>
        </div>
        <output aria-live="polite">
          {officeMessage(messages, 'presentation.presenter.slideOf', { current: String(index + 1), total: String(total) })}
        </output>
      </header>

      <div className="work-presentation-presenter-grid">
        <section
          className="work-presentation-presenter-current"
          aria-label={officeMessage(messages, 'presentation.presenter.currentSlideAria')}
        >
          <h2>{slide.name}</h2>
          <SlideCanvas
            animationCueIndex={animationCueIndex}
            content={content}
            slide={slide}
            interactive={false}
            aspectRatio={aspectRatio}
          />
        </section>
        <section
          className="work-presentation-presenter-next"
          aria-label={officeMessage(messages, 'presentation.presenter.nextSlideAria')}
        >
          <h2>{officeMessage(messages, 'presentation.presenter.next')}</h2>
          {nextSlide ? (
            <>
              <SlideCanvas
                content={content}
                slide={nextSlide}
                interactive={false}
                aspectRatio={aspectRatio}
              />
              <span>{nextSlide.name}</span>
            </>
          ) : (
            <p>{officeMessage(messages, 'presentation.presenter.end')}</p>
          )}
        </section>
        <aside
          className="work-presentation-presenter-notes"
          aria-label={officeMessage(messages, 'presentation.presenter.notesAria')}
        >
          <h2>{officeMessage(messages, 'presentation.presenter.notes')}</h2>
          <p>{slide.notes?.trim() || officeMessage(messages, 'presentation.presenter.notesEmpty')}</p>
        </aside>
      </div>
    </section>
  );
}

function presentationTimerElapsedSeconds(
  timer: PresentationTimerController,
  now: number,
): number {
  return Math.floor(presentationTimerElapsedMilliseconds(timer, now) / 1000);
}

function presentationTimerElapsedMilliseconds(
  timer: PresentationTimerController,
  now: number,
): number {
  if (timer.runningSinceMilliseconds === null) {
    return timer.elapsedMilliseconds;
  }
  return (
    timer.elapsedMilliseconds +
    Math.max(0, now - timer.runningSinceMilliseconds)
  );
}

function formatDuration(seconds: number): string {
  const minutes = Math.floor(seconds / 60);
  const remainder = seconds % 60;
  return `${String(minutes).padStart(2, '0')}:${String(remainder).padStart(2, '0')}`;
}
