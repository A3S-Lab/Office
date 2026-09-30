import { type PointerEvent as ReactPointerEvent, useRef } from 'react';

const ROTATE_COMMIT_DEGREES = 1;

export function PresentationRotateHandle({
  label,
  onRotate,
}: {
  label: string;
  onRotate: (degrees: number) => void;
}) {
  const dragRef = useRef<{
    centerX: number;
    centerY: number;
    pointerId: number;
    startAngle: number;
  } | null>(null);

  const begin = (event: ReactPointerEvent<HTMLButtonElement>) => {
    if (event.button !== 0) return;
    const frame = event.currentTarget.parentElement?.getBoundingClientRect();
    if (!frame || frame.width <= 0 || frame.height <= 0) return;
    event.preventDefault();
    event.stopPropagation();
    event.currentTarget.setPointerCapture(event.pointerId);
    const centerX = frame.left + frame.width / 2;
    const centerY = frame.top + frame.height / 2;
    dragRef.current = {
      centerX,
      centerY,
      pointerId: event.pointerId,
      startAngle: pointerAngle(centerX, centerY, event.clientX, event.clientY),
    };
  };

  const finish = (event: ReactPointerEvent<HTMLButtonElement>) => {
    const drag = dragRef.current;
    if (!drag || drag.pointerId !== event.pointerId) return;
    dragRef.current = null;
    const next = pointerAngle(
      drag.centerX,
      drag.centerY,
      event.clientX,
      event.clientY,
    );
    const delta = signedDegrees(next - drag.startAngle);
    if (Math.abs(delta) < ROTATE_COMMIT_DEGREES) return;
    onRotate(delta);
  };

  return (
    <button
      type="button"
      aria-label={label}
      className="work-slide-selection-rotate-handle"
      data-presentation-rotate-handle
      onClick={(event) => event.stopPropagation()}
      onPointerDown={begin}
      onPointerUp={finish}
    />
  );
}

function pointerAngle(
  centerX: number,
  centerY: number,
  x: number,
  y: number,
): number {
  return (Math.atan2(y - centerY, x - centerX) * 180) / Math.PI;
}

function signedDegrees(delta: number): number {
  const wrapped = ((((delta + 180) % 360) + 360) % 360) - 180;
  return Math.round(wrapped * 100) / 100;
}
