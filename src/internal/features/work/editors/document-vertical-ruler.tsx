import {
  type CSSProperties,
  type KeyboardEvent,
  type PointerEvent,
  useMemo,
  useRef,
} from 'react';
import {
  clampDocumentMargin,
  millimetersToPixels,
} from '../work-document-layout';
import type { WorkDocumentSectionLayout } from '../work-types';
import { documentRulerPageHeight } from './document-ruler';

type DocumentVerticalRulerHandleKind = 'page-top-margin' | 'page-bottom-margin';

export interface DocumentVerticalRulerProps {
  layout: WorkDocumentSectionLayout;
  onLayoutChange: (layout: WorkDocumentSectionLayout) => void;
}

interface DocumentVerticalRulerTick {
  label: string | null;
  position: number;
}

export function DocumentVerticalRuler({
  layout,
  onLayoutChange,
}: DocumentVerticalRulerProps) {
  const rulerRef = useRef<HTMLFieldSetElement>(null);
  const pageHeight = documentRulerPageHeight(layout);
  const marginTop = millimetersToPixels(layout.margins.top);
  const marginBottom = millimetersToPixels(layout.margins.bottom);
  const ticks = useMemo(
    () => documentVerticalRulerTicks(layout, pageHeight),
    [layout.margins.bottom, layout.margins.top, pageHeight],
  );

  const updateMargin = (side: 'top' | 'bottom', value: number): void => {
    const nextValue = clampDocumentMargin(value);
    if (nextValue === layout.margins[side]) return;
    onLayoutChange({
      ...layout,
      margins: {
        ...layout.margins,
        [side]: nextValue,
      },
    });
  };

  const updateFromPointer = (
    kind: DocumentVerticalRulerHandleKind,
    clientY: number,
  ): void => {
    const ruler = rulerRef.current;
    if (!ruler) return;
    const rectangle = ruler.getBoundingClientRect();
    if (rectangle.height <= 0) return;
    const position = Math.min(
      pageHeight,
      Math.max(0, ((clientY - rectangle.top) / rectangle.height) * pageHeight),
    );
    if (kind === 'page-top-margin') {
      updateMargin('top', pixelsToMillimeters(position));
      return;
    }
    updateMargin('bottom', pixelsToMillimeters(pageHeight - position));
  };

  const updateFromKeyboard = (
    kind: DocumentVerticalRulerHandleKind,
    event: KeyboardEvent<HTMLButtonElement>,
  ): void => {
    const direction =
      event.key === 'ArrowUp' || event.key === 'ArrowRight'
        ? 1
        : event.key === 'ArrowDown' || event.key === 'ArrowLeft'
          ? -1
          : event.key === 'PageUp'
            ? 1
            : event.key === 'PageDown'
              ? -1
              : 0;
    if (!direction && event.key !== 'Home' && event.key !== 'End') return;
    event.preventDefault();
    const side = kind === 'page-top-margin' ? 'top' : 'bottom';
    const current = layout.margins[side];
    const step =
      event.shiftKey || event.key === 'PageUp' || event.key === 'PageDown'
        ? 5
        : 1;
    const next =
      event.key === 'Home'
        ? 5
        : event.key === 'End'
          ? 60
          : current + direction * step;
    updateMargin(side, next);
  };

  return (
    <div
      className="work-document-vertical-ruler-shell"
      style={{ height: `${pageHeight}px` }}
    >
      <fieldset
        ref={rulerRef}
        className="work-document-vertical-ruler"
        style={
          {
            '--work-document-vertical-ruler-margin-top': `${marginTop}px`,
            '--work-document-vertical-ruler-margin-bottom': `${marginBottom}px`,
          } as CSSProperties
        }
      >
        <legend className="sr-only">垂直标尺</legend>
        <span
          aria-hidden="true"
          className="work-document-vertical-ruler-margin top"
        />
        <span
          aria-hidden="true"
          className="work-document-vertical-ruler-margin bottom"
        />
        <span aria-hidden="true" className="work-document-vertical-ruler-track">
          {ticks.map((tick) => (
            <span
              key={`${tick.position}-${tick.label ?? 'minor'}`}
              className={tick.label ? 'major' : 'minor'}
              style={{ top: `${tick.position}px` }}
            >
              {tick.label}
            </span>
          ))}
        </span>
        <DocumentVerticalRulerHandle
          kind="page-top-margin"
          label="上页边距"
          position={marginTop}
          value={layout.margins.top}
          onKeyboardChange={updateFromKeyboard}
          onPointerChange={updateFromPointer}
        />
        <DocumentVerticalRulerHandle
          kind="page-bottom-margin"
          label="下页边距"
          position={pageHeight - marginBottom}
          value={layout.margins.bottom}
          onKeyboardChange={updateFromKeyboard}
          onPointerChange={updateFromPointer}
        />
      </fieldset>
    </div>
  );
}

function DocumentVerticalRulerHandle({
  kind,
  label,
  position,
  value,
  onKeyboardChange,
  onPointerChange,
}: {
  kind: DocumentVerticalRulerHandleKind;
  label: string;
  position: number;
  value: number;
  onKeyboardChange: (
    kind: DocumentVerticalRulerHandleKind,
    event: KeyboardEvent<HTMLButtonElement>,
  ) => void;
  onPointerChange: (
    kind: DocumentVerticalRulerHandleKind,
    clientY: number,
  ) => void;
}) {
  const activePointer = useRef<number | null>(null);
  const startPointer = (event: PointerEvent<HTMLButtonElement>): void => {
    if (event.button !== 0) return;
    event.preventDefault();
    activePointer.current = event.pointerId;
    event.currentTarget.setPointerCapture?.(event.pointerId);
    onPointerChange(kind, event.clientY);
  };
  const movePointer = (event: PointerEvent<HTMLButtonElement>): void => {
    if (activePointer.current !== event.pointerId) return;
    event.preventDefault();
    onPointerChange(kind, event.clientY);
  };
  const finishPointer = (event: PointerEvent<HTMLButtonElement>): void => {
    if (activePointer.current !== event.pointerId) return;
    event.preventDefault();
    onPointerChange(kind, event.clientY);
    activePointer.current = null;
    if (event.currentTarget.hasPointerCapture?.(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId);
    }
  };

  return (
    <button
      type="button"
      role="slider"
      className={`work-document-vertical-ruler-handle ${kind}`}
      aria-label={label}
      aria-orientation="vertical"
      aria-valuemax={60}
      aria-valuemin={5}
      aria-valuenow={value}
      aria-valuetext={`${formatRulerValue(value)} 毫米`}
      title={`${label}：${formatRulerValue(value)} 毫米`}
      style={
        {
          '--work-document-vertical-ruler-handle-position': `${position}px`,
        } as CSSProperties
      }
      onKeyDown={(event) => onKeyboardChange(kind, event)}
      onPointerCancel={finishPointer}
      onPointerDown={startPointer}
      onPointerMove={movePointer}
      onPointerUp={finishPointer}
    >
      <span aria-hidden="true" />
    </button>
  );
}

function documentVerticalRulerTicks(
  layout: WorkDocumentSectionLayout,
  pageHeight: number,
): DocumentVerticalRulerTick[] {
  const top = millimetersToPixels(layout.margins.top);
  const textHeight = Math.max(
    0,
    pageHeight - top - millimetersToPixels(layout.margins.bottom),
  );
  const textHeightMillimeters = pixelsToMillimeters(textHeight);
  const ticks: DocumentVerticalRulerTick[] = [];
  for (
    let distance = 0;
    distance <= textHeightMillimeters + 0.01;
    distance += 5
  ) {
    ticks.push({
      label: distance > 0 && distance % 10 === 0 ? String(distance / 10) : null,
      position: top + millimetersToPixels(distance),
    });
  }
  return ticks;
}

function pixelsToMillimeters(value: number): number {
  return (value * 25.4) / 96;
}

function formatRulerValue(value: number): string {
  return Number.isInteger(value) ? String(value) : value.toFixed(1);
}
