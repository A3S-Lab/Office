import { createContext, useContext, useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import {
  paintTextCaretInBox,
  type FrameCaretPaint,
  type TextCaretBox,
  type WorkOfficeCollaborationFrameCaret,
} from '../../../collaboration/office-collaboration-frame-caret';
import type { WorkSlideElement } from '../work-types';
import {
  officePresenceColorStyle,
  useOfficeRemoteParticipants,
} from './office-collaboration-presence-ui';

type PresentationFrameCaret = Extract<
  WorkOfficeCollaborationFrameCaret,
  { kind: 'presentation' }
>;

const PresentationFrameCaretContext = createContext<PresentationFrameCaret | null>(
  null,
);

export function PresentationFrameCaretProvider({
  caret,
  children,
}: {
  caret: PresentationFrameCaret | null;
  children: ReactNode;
}) {
  return (
    <PresentationFrameCaretContext.Provider value={caret}>
      {children}
    </PresentationFrameCaretContext.Provider>
  );
}

export function presentationFrameCaretPaint(input: {
  readonly box: TextCaretBox;
  readonly caret: PresentationFrameCaret;
  readonly element: WorkSlideElement;
  readonly layoutSettled: boolean;
  readonly measure: (slice: string) => number;
}): FrameCaretPaint | null {
  if (input.caret.elementId !== input.element.id) return null;
  if (input.caret.indexUtf16 === undefined) return null;
  const runs = input.element.textRuns?.filter((run) => run.text.length) ?? [];
  const text = runs.length
    ? runs.map((run) => run.text).join('')
    : input.element.text;
  return paintTextCaretInBox({
    align: input.element.align,
    box: input.box,
    fontSize: input.element.fontSize,
    indexUtf16: input.caret.indexUtf16,
    layoutSettled: input.layoutSettled,
    measure: input.measure,
    text,
    verticalAlign: input.element.verticalAlign ?? 'top',
  });
}

export function PresentationCollaborationPresenceLayer({
  elements,
  measure,
  slideId,
  textBox,
}: {
  elements: readonly WorkSlideElement[];
  measure?: (slice: string) => number;
  slideId: string;
  textBox?: TextCaretBox;
}) {
  const frameCaret = useContext(PresentationFrameCaretContext);
  const participants = useOfficeRemoteParticipants().filter(
    (participant) =>
      participant.location?.kind === 'presentation' &&
      participant.location.slideId === slideId,
  );
  const frameElement =
    frameCaret?.indexUtf16 !== undefined
      ? elements.find((element) => element.id === frameCaret.elementId)
      : undefined;
  if (!participants.length && !frameElement) return null;

  return (
    <div className="work-presentation-remote-presence-layer" aria-hidden="true">
      {frameCaret && frameElement ? (
        <PresentationFrameCaretMark
          caret={frameCaret}
          element={frameElement}
          measure={measure}
          textBox={textBox}
        />
      ) : null}
      {participants.flatMap((participant, participantIndex) => {
        const location = participant.location;
        if (location?.kind !== 'presentation') return [];
        const selected = location.elementIds.flatMap((elementId) => {
          const element = elements.find(
            (candidate) => candidate.id === elementId,
          );
          return element ? [element] : [];
        });
        if (!selected.length) {
          return [
            <span
              className="work-presentation-remote-slide-marker"
              data-activity={participant.activity}
              data-participant-id={participant.actor.id}
              key={`${participant.presenceId}:slide`}
              style={{
                ...officePresenceColorStyle(participant),
                top: 8 + participantIndex * 22,
              }}
            >
              {participant.actor.name}
            </span>,
          ];
        }
        return selected.map((element, index) => (
          <span
            className="work-presentation-remote-element"
            data-activity={participant.activity}
            data-participant-id={participant.actor.id}
            data-remote-slide-element-id={element.id}
            key={`${participant.presenceId}:${element.id}`}
            style={{
              ...officePresenceColorStyle(participant),
              left: `${element.x}%`,
              top: `${element.y}%`,
              width: `${element.width}%`,
              height: `${element.height}%`,
              transform: element.rotation
                ? `rotate(${element.rotation}deg)`
                : undefined,
            }}
          >
            {index === 0 && <span>{participant.actor.name}</span>}
          </span>
        ));
      })}
    </div>
  );
}

function PresentationFrameCaretMark({
  caret,
  element,
  measure,
  textBox,
}: {
  caret: PresentationFrameCaret;
  element: WorkSlideElement;
  measure?: (slice: string) => number;
  textBox?: TextCaretBox;
}) {
  const hostRef = useRef<HTMLSpanElement>(null);
  const [measured, setMeasured] = useState<TextCaretBox | null>(textBox ?? null);
  useLayoutEffect(() => {
    if (textBox) {
      setMeasured(textBox);
      return;
    }
    const node = hostRef.current;
    if (!node) return;
    const rect = node.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) return;
    setMeasured({ height: rect.height, left: 0, top: 0, width: rect.width });
  }, [caret, element, textBox]);
  const paint = measured
    ? presentationFrameCaretPaint({
        box: {
          height: measured.height,
          left: 0,
          top: 0,
          width: measured.width,
        },
        caret,
        element,
        layoutSettled: true,
        measure: measure ?? ((slice) => slice.length * element.fontSize),
      })
    : null;
  return (
    <span
      className="work-presentation-frame-caret-host"
      data-frame-caret-element-id={element.id}
      ref={hostRef}
      style={{
        left: `${element.x}%`,
        top: `${element.y}%`,
        width: `${element.width}%`,
        height: `${element.height}%`,
      }}
    >
      {paint ? (
        <i
          className="work-presentation-frame-caret"
          data-frame-caret-index={paint.head}
          style={{
            left: paint.left,
            top: paint.top,
            width: paint.width,
            height: paint.height,
          }}
        />
      ) : null}
    </span>
  );
}
