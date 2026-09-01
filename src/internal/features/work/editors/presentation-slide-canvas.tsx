import {
  presentationSlideView,
  presentationSlideViewFromDesign,
  type WorkPresentationDesignContent,
} from '../work-presentation-layouts';
import {
  workSlideAnimationClass,
  workSlideAnimationCues,
  type WorkSlideAnimationCueItem,
} from '../work-presentation-animation';
import type {
  WorkPresentationContent,
  WorkSlide,
  WorkSlideAnimation,
  WorkSlideElement,
} from '../work-types';
import { OfficeTextArea } from './office-controls';
import { SlideChart } from './presentation-chart-canvas';

export function SlideCanvas({
  content,
  designContent,
  slide,
  interactive,
  aspectRatio,
  animationCueIndex,
  showPlaceholders = false,
}: {
  content?: WorkPresentationContent;
  designContent?: WorkPresentationDesignContent;
  slide: WorkSlide;
  interactive: boolean;
  aspectRatio: string;
  animationCueIndex?: number;
  showPlaceholders?: boolean;
}) {
  const view = designContent
    ? presentationSlideViewFromDesign(designContent, slide)
    : content
      ? presentationSlideView(content, slide)
      : undefined;
  const elements = [
    ...(view?.inheritedElements.map((element) => ({
      element,
      origin: 'inherited' as const,
    })) ?? []),
    ...slide.elements.map((element) => ({
      element,
      origin: 'slide' as const,
    })),
  ];
  const animationPlayback =
    animationCueIndex === undefined
      ? undefined
      : slideAnimationPlayback(slide.animations, animationCueIndex);
  return (
    <span
      className={`work-slide-canvas ${interactive ? 'interactive' : ''}`}
      style={{ background: view?.background ?? slide.background, aspectRatio }}
    >
      {elements.map(({ element, origin }) => (
        <SlideElementPreview
          element={element}
          key={`${origin}:${element.id}`}
          origin={origin}
          showPlaceholder={showPlaceholders}
          animationPlayback={
            origin === 'slide' ? animationPlayback?.get(element.id) : undefined
          }
        />
      ))}
    </span>
  );
}

export function SlideElementPreview({
  animationPlayback,
  element,
  origin,
  showPlaceholder = false,
}: {
  animationPlayback?: SlideElementAnimationPlayback;
  element: WorkSlideElement;
  origin: 'inherited' | 'slide';
  showPlaceholder?: boolean;
}) {
  const hasRichText = element.textRuns?.some((run) => run.text.length > 0);
  return (
    <span
      className={`work-slide-element ${element.type} ${origin} ${showPlaceholder && element.placeholder ? 'placeholder' : ''}`.trim()}
      data-slide-preview-element-id={element.id}
      data-slide-element-origin={origin}
      data-slide-animation-effect={
        animationPlayback?.animations.at(-1)?.animation.effect
      }
      data-slide-animation-effects={animationPlayback?.animations
        .map(({ animation }) => animation.effect)
        .join(' ')}
      data-slide-animation-final-visibility={
        animationPlayback
          ? animationPlayback.finallyVisible
            ? 'visible'
            : 'hidden'
          : undefined
      }
      data-slide-animation-initial-visibility={
        animationPlayback
          ? animationPlayback.initiallyVisible
            ? 'visible'
            : 'hidden'
          : undefined
      }
      data-slide-animation-state={animationPlayback?.state}
      style={{
        ...slideElementStyle(element),
        ...slideElementAnimationStyle(animationPlayback),
      }}
    >
      {element.type === 'image' && element.image ? (
        <img
          src={element.image.dataUrl}
          alt={element.altText ?? element.image.name}
        />
      ) : element.type === 'table' && element.table ? (
        <SlideTablePreview element={element} />
      ) : element.type === 'chart' && element.chart ? (
        <SlideChart
          chart={element.chart}
          label={element.altText ?? element.chart.title ?? '图表'}
        />
      ) : hasRichText ||
        element.text ||
        (showPlaceholder && element.placeholder?.prompt) ? (
        <SlideElementTextPreview
          element={element}
          showPlaceholder={showPlaceholder}
        />
      ) : null}
    </span>
  );
}

export function EditableSlideTable({
  element,
  onChange,
}: {
  element: WorkSlideElement;
  onChange: (rows: string[][]) => void;
}) {
  const table = element.table;
  if (!table) return null;
  return (
    <table
      className="work-slide-table editable"
      aria-label={element.altText ?? '幻灯片表格'}
      data-slide-editor
    >
      <tbody>
        {table.rows.map((row, rowIndex) => (
          <tr key={rowIndex}>
            {row.map((cell, columnIndex) => {
              const Cell = rowIndex < (table.headerRows ?? 0) ? 'th' : 'td';
              return (
                <Cell key={columnIndex}>
                  <OfficeTextArea
                    aria-label={`第 ${rowIndex + 1} 行第 ${columnIndex + 1} 列`}
                    value={cell}
                    onChange={(event) => {
                      const rows = table.rows.map((current) => [...current]);
                      rows[rowIndex][columnIndex] = event.target.value;
                      onChange(rows);
                    }}
                  />
                </Cell>
              );
            })}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

export function slideElementStyle(
  element: WorkSlideElement,
): React.CSSProperties {
  const shapeClipPath =
    element.shapeType === 'triangle'
      ? 'polygon(50% 0, 100% 100%, 0 100%)'
      : element.shapeType === 'diamond'
        ? 'polygon(50% 0, 100% 50%, 50% 100%, 0 50%)'
        : undefined;
  return {
    left: `${element.x}%`,
    top: `${element.y}%`,
    width: `${element.width}%`,
    height: `${element.height}%`,
    background: element.type === 'line' ? 'transparent' : element.fill,
    border:
      element.type === 'line'
        ? undefined
        : element.borderWidth
          ? `${element.borderWidth}px solid ${element.borderColor ?? element.fill}`
          : undefined,
    borderRadius:
      element.shapeType === 'ellipse' ? '50%' : `${element.radius ?? 0}%`,
    clipPath: shapeClipPath,
    opacity: element.opacity,
    '--work-slide-element-opacity': element.opacity ?? 1,
    '--work-slide-element-rotation': `${element.rotation ?? 0}deg`,
    transform: element.rotation ? `rotate(${element.rotation}deg)` : undefined,
    transformOrigin: 'center',
    ...(element.type === 'line'
      ? {
          height: `${Math.max(0.5, element.borderWidth ?? 1)}px`,
          borderTop: `${Math.max(0.5, element.borderWidth ?? 1)}px solid ${element.borderColor ?? element.color}`,
        }
      : {}),
  } as React.CSSProperties;
}

interface SlideElementAnimationPlayback {
  animations: WorkSlideAnimationCueItem[];
  finallyVisible: boolean;
  initiallyVisible: boolean;
  state: 'hidden' | 'playing' | 'visible';
}

function slideAnimationPlayback(
  animations: readonly WorkSlideAnimation[] | undefined,
  activeCueIndex: number,
): Map<string, SlideElementAnimationPlayback> {
  const cues = workSlideAnimationCues(animations);
  const playback = new Map<string, SlideElementAnimationPlayback>();
  for (const animation of animations ?? []) {
    if (playback.has(animation.elementId)) continue;
    const visible = workSlideAnimationClass(animation.effect) === 'exit';
    playback.set(animation.elementId, {
      animations: [],
      finallyVisible: visible,
      initiallyVisible: visible,
      state: visible ? 'visible' : 'hidden',
    });
  }

  if (activeCueIndex < 0) return playback;
  for (const [cueIndex, cue] of cues.entries()) {
    if (cueIndex > activeCueIndex) break;
    for (const item of cue.items) {
      const state = playback.get(item.animation.elementId);
      if (!state) continue;
      const finallyVisible =
        workSlideAnimationClass(item.animation.effect) === 'entrance';
      if (cueIndex < activeCueIndex) {
        state.finallyVisible = finallyVisible;
        state.initiallyVisible = finallyVisible;
        state.state = finallyVisible ? 'visible' : 'hidden';
        continue;
      }
      if (!state.animations.length) {
        state.initiallyVisible = state.finallyVisible;
      }
      state.animations.push(item);
      state.finallyVisible = finallyVisible;
      state.state = 'playing';
    }
  }
  return playback;
}

function slideElementAnimationStyle(
  playback: SlideElementAnimationPlayback | undefined,
): React.CSSProperties {
  if (!playback?.animations.length) return {};
  return {
    animationDelay: playback.animations
      .map(({ startOffsetMs }) => `${startOffsetMs}ms`)
      .join(', '),
    animationDuration: playback.animations
      .map(({ animation }) => `${animation.durationMs}ms`)
      .join(', '),
    animationFillMode: 'forwards',
    animationName: playback.animations
      .map(({ animation }) => slideAnimationName(animation))
      .join(', '),
    animationTimingFunction: playback.animations
      .map(({ animation }) => slideAnimationTimingFunction(animation))
      .join(', '),
  } as React.CSSProperties;
}

function slideAnimationName(animation: WorkSlideAnimation): string {
  if (animation.effect === 'fly-in' || animation.effect === 'fly-out') {
    return `work-slide-animation-${animation.effect}-${animation.direction ?? 'left'}`;
  }
  return `work-slide-animation-${animation.effect}`;
}

function slideAnimationTimingFunction(animation: WorkSlideAnimation): string {
  if (animation.effect === 'appear') return 'steps(1, start)';
  if (animation.effect === 'disappear') return 'steps(1, end)';
  return 'cubic-bezier(0.16, 1, 0.3, 1)';
}

export function slideTextStyle(element: WorkSlideElement): React.CSSProperties {
  return {
    color: element.color,
    fontFamily: element.fontFamily,
    fontSize: `clamp(6px, ${element.fontSize / 10}cqw, ${element.fontSize}px)`,
    fontWeight: element.bold ? 700 : 400,
    fontStyle: element.italic ? 'italic' : undefined,
    textDecoration: element.underline ? 'underline' : undefined,
    textAlign: element.align,
  };
}

export function SlideElementTextPreview({
  element,
  showPlaceholder = false,
}: {
  element: WorkSlideElement;
  showPlaceholder?: boolean;
}) {
  const placeholderText = showPlaceholder
    ? element.placeholder?.prompt?.trim()
    : undefined;
  const plainText = element.text || placeholderText || '';
  const hasRichText = element.textRuns?.some((run) => run.text.length > 0);
  const content = hasRichText ? (
    <span className="work-slide-rich-text" style={slideTextStyle(element)}>
      {element.textRuns?.map((run, index) => {
        const style: React.CSSProperties = {
          color: run.color,
          fontFamily: run.fontFamily,
          fontSize: run.fontSize
            ? `clamp(6px, ${run.fontSize / 10}cqw, ${run.fontSize}px)`
            : undefined,
          fontStyle:
            run.italic === undefined
              ? undefined
              : run.italic
                ? 'italic'
                : 'normal',
          fontWeight: run.bold === undefined ? undefined : run.bold ? 700 : 400,
          textDecoration:
            run.underline === undefined
              ? undefined
              : run.underline
                ? 'underline'
                : 'none',
        };
        return run.href ? (
          <a
            href={run.href}
            target="_blank"
            rel="noreferrer"
            key={`${index}-${run.text}`}
            style={style}
          >
            {run.text}
          </a>
        ) : (
          <span key={`${index}-${run.text}`} style={style}>
            {run.text}
          </span>
        );
      })}
    </span>
  ) : (
    <span
      className={
        !element.text && placeholderText
          ? 'work-slide-placeholder-text'
          : undefined
      }
      style={slideTextStyle(element)}
    >
      {plainText}
    </span>
  );
  return element.href ? (
    <a
      className="work-slide-element-link"
      href={element.href}
      target="_blank"
      rel="noreferrer"
    >
      {content}
    </a>
  ) : (
    content
  );
}

export function SlideTablePreview({ element }: { element: WorkSlideElement }) {
  const table = element.table;
  if (!table) return null;
  return (
    <table
      className="work-slide-table"
      aria-label={element.altText ?? '幻灯片表格'}
    >
      <tbody>
        {table.rows.map((row, rowIndex) => (
          <tr key={`${rowIndex}-${row.join('|')}`}>
            {row.map((cell, columnIndex) => {
              const Cell = rowIndex < (table.headerRows ?? 0) ? 'th' : 'td';
              return <Cell key={`${columnIndex}-${cell}`}>{cell}</Cell>;
            })}
          </tr>
        ))}
      </tbody>
    </table>
  );
}
