import {
  type CSSProperties,
  type KeyboardEvent,
  type MouseEvent,
  type PointerEvent,
  type RefObject,
  useRef,
} from 'react';
import { documentTabLeaderLabel } from '../work-document-tab-node';
import {
  DOCUMENT_TAB_RULER_STEP_PX,
  type DocumentTabAlignment,
  type DocumentTabStop,
  nextDocumentTabAlignment,
  normalizeDocumentTabStops,
  normalizedTabPosition,
} from '../work-document-tab-stops';

export interface DocumentRulerTabStopsProps {
  marginLeft: number;
  pageWidth: number;
  rulerRef: RefObject<HTMLFieldSetElement | null>;
  tabStops: readonly DocumentTabStop[];
  textWidth: number;
  onChange: (tabStops: DocumentTabStop[]) => void;
}

export function DocumentRulerTabStops({
  marginLeft,
  pageWidth,
  rulerRef,
  tabStops,
  textWidth,
  onChange,
}: DocumentRulerTabStopsProps) {
  const normalized = constrainRulerTabStops(tabStops, textWidth);
  const positionFromClientX = (clientX: number): number | null => {
    const ruler = rulerRef.current;
    if (!ruler) return null;
    const rectangle = ruler.getBoundingClientRect();
    if (rectangle.width <= 0) return null;
    const rulerPosition =
      ((clientX - rectangle.left) / rectangle.width) * pageWidth;
    return Math.min(
      Math.floor(textWidth),
      Math.max(1, Math.round(rulerPosition - marginLeft)),
    );
  };
  const addAtPointer = (event: MouseEvent<HTMLButtonElement>): void => {
    const position = positionFromClientX(event.clientX);
    if (position === null) return;
    const existingIndex = normalized.findIndex(
      (stop) => Math.abs(stop.position - position) < 1,
    );
    if (existingIndex >= 0) {
      const existing = normalized[existingIndex];
      if (!existing) return;
      onChange(
        replaceTabStop(normalized, existingIndex, {
          ...existing,
          alignment: nextDocumentTabAlignment(existing.alignment),
        }),
      );
      return;
    }
    onChange(
      normalizeDocumentTabStops([
        ...normalized,
        { position, alignment: 'left', leader: 'none' },
      ]),
    );
  };
  const addFromKeyboard = (event: KeyboardEvent<HTMLButtonElement>): void => {
    if (event.key !== 'Enter' && event.key !== ' ') return;
    event.preventDefault();
    const last = normalized.at(-1)?.position ?? 0;
    const position = Math.min(
      Math.floor(textWidth),
      Math.max(1, Math.floor(last / 48 + 1) * 48),
    );
    onChange(
      normalizeDocumentTabStops([
        ...normalized,
        { position, alignment: 'left', leader: 'none' },
      ]),
    );
  };
  const updateTabStop = (index: number, next: DocumentTabStop): void => {
    onChange(
      replaceTabStop(normalized, index, {
        ...next,
        position: Math.min(
          Math.floor(textWidth),
          Math.max(1, normalizedTabPosition(next.position)),
        ),
      }),
    );
  };

  return (
    <>
      <button
        type="button"
        className="work-document-ruler-tab-track"
        aria-label="在标尺上添加制表位"
        title="单击添加制表位"
        style={
          {
            '--work-document-ruler-tab-track-left': `${marginLeft}px`,
            '--work-document-ruler-tab-track-width': `${textWidth}px`,
          } as CSSProperties
        }
        onClick={addAtPointer}
        onKeyDown={addFromKeyboard}
      />
      {normalized.map((tabStop, index) => (
        <DocumentRulerTabStopHandle
          key={`${tabStop.position}-${tabStop.alignment}-${tabStop.leader}`}
          index={index}
          marginLeft={marginLeft}
          maximum={textWidth}
          positionFromClientX={positionFromClientX}
          tabStop={tabStop}
          onChange={(next) => updateTabStop(index, next)}
          onRemove={() =>
            onChange(normalized.filter((_, stopIndex) => stopIndex !== index))
          }
        />
      ))}
    </>
  );
}

function DocumentRulerTabStopHandle({
  index,
  marginLeft,
  maximum,
  positionFromClientX,
  tabStop,
  onChange,
  onRemove,
}: {
  index: number;
  marginLeft: number;
  maximum: number;
  positionFromClientX: (clientX: number) => number | null;
  tabStop: DocumentTabStop;
  onChange: (tabStop: DocumentTabStop) => void;
  onRemove: () => void;
}) {
  const activePointer = useRef<number | null>(null);
  const updateFromPointer = (clientX: number): void => {
    const position = positionFromClientX(clientX);
    if (position !== null) onChange({ ...tabStop, position });
  };
  const startPointer = (event: PointerEvent<HTMLButtonElement>): void => {
    if (event.button !== 0) return;
    event.preventDefault();
    event.stopPropagation();
    activePointer.current = event.pointerId;
    event.currentTarget.setPointerCapture?.(event.pointerId);
    updateFromPointer(event.clientX);
  };
  const movePointer = (event: PointerEvent<HTMLButtonElement>): void => {
    if (activePointer.current !== event.pointerId) return;
    event.preventDefault();
    updateFromPointer(event.clientX);
  };
  const finishPointer = (event: PointerEvent<HTMLButtonElement>): void => {
    if (activePointer.current !== event.pointerId) return;
    event.preventDefault();
    updateFromPointer(event.clientX);
    activePointer.current = null;
    if (event.currentTarget.hasPointerCapture?.(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId);
    }
  };
  const updateFromKeyboard = (
    event: KeyboardEvent<HTMLButtonElement>,
  ): void => {
    if (event.key === 'Delete' || event.key === 'Backspace') {
      event.preventDefault();
      onRemove();
      return;
    }
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      onChange({
        ...tabStop,
        alignment: nextDocumentTabAlignment(tabStop.alignment),
      });
      return;
    }
    const direction =
      event.key === 'ArrowLeft' || event.key === 'ArrowDown'
        ? -1
        : event.key === 'ArrowRight' || event.key === 'ArrowUp'
          ? 1
          : 0;
    if (!direction && event.key !== 'Home' && event.key !== 'End') return;
    event.preventDefault();
    const position =
      event.key === 'Home'
        ? 1
        : event.key === 'End'
          ? maximum
          : tabStop.position +
            direction * (event.shiftKey ? 1 : DOCUMENT_TAB_RULER_STEP_PX);
    onChange({ ...tabStop, position });
  };
  const label = `${tabAlignmentLabel(tabStop.alignment)}制表位 ${index + 1}`;
  const valueText = `${formatCentimeters(tabStop.position)} 厘米，${documentTabLeaderLabel(tabStop.leader)}`;

  return (
    <button
      type="button"
      role="slider"
      className={`work-document-ruler-tab-stop ${tabStop.alignment}`}
      data-leader={tabStop.leader}
      aria-label={label}
      aria-orientation="horizontal"
      aria-valuemax={Math.floor(maximum)}
      aria-valuemin={1}
      aria-valuenow={Math.round(tabStop.position)}
      aria-valuetext={valueText}
      title={`${label}：${valueText}；按 Enter 切换对齐，按 Delete 删除`}
      style={
        {
          '--work-document-ruler-tab-position': `${marginLeft + tabStop.position}px`,
        } as CSSProperties
      }
      onDoubleClick={() =>
        onChange({
          ...tabStop,
          alignment: nextDocumentTabAlignment(tabStop.alignment),
        })
      }
      onKeyDown={updateFromKeyboard}
      onPointerCancel={finishPointer}
      onPointerDown={startPointer}
      onPointerMove={movePointer}
      onPointerUp={finishPointer}
    >
      <span aria-hidden="true" />
    </button>
  );
}

export function constrainRulerTabStops(
  tabStops: readonly DocumentTabStop[],
  textWidth: number,
): DocumentTabStop[] {
  return normalizeDocumentTabStops(tabStops)
    .filter((stop) => stop.position <= Math.floor(textWidth))
    .map((stop) => ({ ...stop }));
}

function replaceTabStop(
  tabStops: readonly DocumentTabStop[],
  index: number,
  next: DocumentTabStop,
): DocumentTabStop[] {
  return normalizeDocumentTabStops(
    tabStops.map((tabStop, stopIndex) =>
      stopIndex === index ? next : tabStop,
    ),
  );
}

function tabAlignmentLabel(alignment: DocumentTabAlignment): string {
  if (alignment === 'center') return '居中';
  if (alignment === 'right') return '右对齐';
  if (alignment === 'decimal') return '小数点对齐';
  return '左对齐';
}

function formatCentimeters(value: number): string {
  return ((value * 25.4) / 96 / 10).toFixed(2);
}
