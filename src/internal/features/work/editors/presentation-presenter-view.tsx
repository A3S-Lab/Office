import { Pause, Play, RotateCcw } from 'lucide-react';
import { useEffect, useState } from 'react';
import type { WorkPresentationContent, WorkSlide } from '../work-types';
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
  content,
  slide,
  nextSlide,
  index,
  total,
  aspectRatio,
  timer,
}: {
  content: WorkPresentationContent;
  slide: WorkSlide;
  nextSlide?: WorkSlide;
  index: number;
  total: number;
  aspectRatio: string;
  timer: PresentationTimerController;
}) {
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
    <section className="work-presentation-presenter" aria-label="演讲者视图">
      <header>
        <div>
          <span>演讲计时</span>
          <strong>
            <span className="sr-only">已用时间：</span>
            <time dateTime={`PT${elapsedSeconds}S`}>
              {formatDuration(elapsedSeconds)}
            </time>
          </strong>
        </div>
        <div className="work-presentation-presenter-timer-actions">
          <button
            type="button"
            aria-label={running ? '暂停计时' : '继续计时'}
            aria-pressed={!running}
            onClick={toggleTimer}
          >
            {running ? <Pause size={15} /> : <Play size={15} />}
          </button>
          <button type="button" aria-label="重置计时" onClick={resetTimer}>
            <RotateCcw size={15} />
          </button>
        </div>
        <output aria-live="polite">
          幻灯片 {index + 1} / {total}
        </output>
      </header>

      <div className="work-presentation-presenter-grid">
        <section
          className="work-presentation-presenter-current"
          aria-label="当前幻灯片"
        >
          <h2>{slide.name}</h2>
          <SlideCanvas
            content={content}
            slide={slide}
            interactive={false}
            aspectRatio={aspectRatio}
          />
        </section>
        <section
          className="work-presentation-presenter-next"
          aria-label="下一张幻灯片"
        >
          <h2>下一张</h2>
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
            <p>演示结束</p>
          )}
        </section>
        <aside
          className="work-presentation-presenter-notes"
          aria-label="演讲者备注"
        >
          <h2>演讲者备注</h2>
          <p>{slide.notes?.trim() || '此页没有演讲者备注'}</p>
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
