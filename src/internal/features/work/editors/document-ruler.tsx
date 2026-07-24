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
import {
  DOCUMENT_INDENT_STEP_PX,
  DOCUMENT_RULER_INDENT_STEP_PX,
  MAX_DOCUMENT_INDENT_LEVEL,
  normalizeDocumentParagraphIndent,
  type DocumentParagraphIndent,
} from '../work-document-paragraph-formatting';
import type { DocumentTabStop } from '../work-document-tab-stops';
import type { WorkDocumentSectionLayout } from '../work-types';
import { DocumentRulerTabStops } from './document-ruler-tab-stops';

type DocumentRulerHandleKind =
  | 'page-left-margin'
  | 'page-right-margin'
  | 'left-indent'
  | 'first-line-indent'
  | 'right-indent';

export interface DocumentRulerProps {
  layout: WorkDocumentSectionLayout;
  paragraphIndent: DocumentParagraphIndent;
  tabStops: readonly DocumentTabStop[];
  onLayoutChange: (layout: WorkDocumentSectionLayout) => void;
  onParagraphIndentChange: (indent: DocumentParagraphIndent) => void;
  onTabStopsChange: (tabStops: DocumentTabStop[]) => void;
}

interface DocumentRulerTick {
  label: string | null;
  position: number;
}

const MINIMUM_TEXT_WIDTH_PX = 48;

export function DocumentRuler({
  layout,
  paragraphIndent,
  tabStops,
  onLayoutChange,
  onParagraphIndentChange,
  onTabStopsChange,
}: DocumentRulerProps) {
  const rulerRef = useRef<HTMLFieldSetElement>(null);
  const pageWidth = documentRulerPageWidth(layout);
  const marginLeft = millimetersToPixels(layout.margins.left);
  const marginRight = millimetersToPixels(layout.margins.right);
  const textWidth = Math.max(1, pageWidth - marginLeft - marginRight);
  const indent = constrainRulerParagraphIndent(paragraphIndent, textWidth);
  const leftIndentPosition = marginLeft + indent.left;
  const firstLineIndentPosition = leftIndentPosition + indent.firstLine;
  const rightIndentPosition = pageWidth - marginRight - indent.right;
  const maximumLeftIndent = maximumLeftIndentPixels(indent, textWidth);
  const maximumRightIndent = maximumRightIndentPixels(indent, textWidth);
  const maximumFirstLineIndent = Math.max(
    0,
    textWidth - indent.left - indent.right - MINIMUM_TEXT_WIDTH_PX,
  );
  const ticks = useMemo(
    () => documentRulerTicks(layout, pageWidth),
    [layout.margins.left, layout.margins.right, pageWidth],
  );

  const updateMargin = (side: 'left' | 'right', value: number): void => {
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

  const updateParagraphIndent = (next: DocumentParagraphIndent): void => {
    const normalized = constrainRulerParagraphIndent(next, textWidth);
    if (
      normalized.left === indent.left &&
      normalized.right === indent.right &&
      normalized.firstLine === indent.firstLine
    )
      return;
    onParagraphIndentChange(normalized);
  };

  const updateFromPointer = (
    kind: DocumentRulerHandleKind,
    clientX: number,
  ): void => {
    const ruler = rulerRef.current;
    if (!ruler) return;
    const rectangle = ruler.getBoundingClientRect();
    if (rectangle.width <= 0) return;
    const position = Math.min(
      pageWidth,
      Math.max(0, ((clientX - rectangle.left) / rectangle.width) * pageWidth),
    );
    if (kind === 'page-left-margin') {
      updateMargin('left', pixelsToMillimeters(position));
      return;
    }
    if (kind === 'page-right-margin') {
      updateMargin('right', pixelsToMillimeters(pageWidth - position));
      return;
    }

    const textPosition = snapIndentPosition(position - marginLeft);
    if (kind === 'left-indent') {
      updateParagraphIndent({ ...indent, left: textPosition });
      return;
    }
    if (kind === 'first-line-indent') {
      updateParagraphIndent({
        ...indent,
        firstLine: textPosition - indent.left,
      });
      return;
    }
    updateParagraphIndent({
      ...indent,
      right: textWidth - textPosition,
    });
  };

  const updateFromKeyboard = (
    kind: DocumentRulerHandleKind,
    event: KeyboardEvent<HTMLButtonElement>,
  ): void => {
    const direction =
      event.key === 'ArrowLeft' || event.key === 'ArrowDown'
        ? -1
        : event.key === 'ArrowRight' || event.key === 'ArrowUp'
          ? 1
          : 0;
    if (!direction && event.key !== 'Home' && event.key !== 'End') return;
    event.preventDefault();

    if (kind === 'left-indent') {
      updateParagraphIndent({
        ...indent,
        left: keyboardIndentValue(
          indent.left,
          0,
          maximumLeftIndent,
          direction,
          event,
        ),
      });
      return;
    }
    if (kind === 'first-line-indent') {
      updateParagraphIndent({
        ...indent,
        firstLine: keyboardIndentValue(
          indent.firstLine,
          -indent.left,
          maximumFirstLineIndent,
          direction,
          event,
        ),
      });
      return;
    }
    if (kind === 'right-indent') {
      updateParagraphIndent({
        ...indent,
        right: keyboardIndentValue(
          indent.right,
          0,
          maximumRightIndent,
          direction,
          event,
        ),
      });
      return;
    }

    const side = kind === 'page-left-margin' ? 'left' : 'right';
    const current = layout.margins[side];
    const step = event.shiftKey ? 5 : 1;
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
      className="work-document-ruler-shell"
      style={{ width: `${pageWidth}px` }}
    >
      <fieldset
        ref={rulerRef}
        className="work-document-ruler"
        style={
          {
            '--work-document-ruler-margin-left': `${marginLeft}px`,
            '--work-document-ruler-margin-right': `${marginRight}px`,
            '--work-document-ruler-text-width': `${textWidth}px`,
          } as CSSProperties
        }
      >
        <legend className="sr-only">水平标尺</legend>
        <span aria-hidden="true" className="work-document-ruler-margin left" />
        <span aria-hidden="true" className="work-document-ruler-margin right" />
        <span aria-hidden="true" className="work-document-ruler-track">
          {ticks.map((tick) => (
            <span
              key={`${tick.position}-${tick.label ?? 'minor'}`}
              className={tick.label ? 'major' : 'minor'}
              style={{ left: `${tick.position}px` }}
            >
              {tick.label}
            </span>
          ))}
        </span>
        <DocumentRulerTabStops
          marginLeft={marginLeft}
          pageWidth={pageWidth}
          rulerRef={rulerRef}
          tabStops={tabStops}
          textWidth={textWidth}
          onChange={onTabStopsChange}
        />
        <DocumentRulerHandle
          kind="page-left-margin"
          label="左页边距"
          maximum={60}
          minimum={5}
          position={marginLeft}
          value={layout.margins.left}
          valueText={`${formatRulerValue(layout.margins.left)} 毫米`}
          onKeyboardChange={updateFromKeyboard}
          onPointerChange={updateFromPointer}
        />
        <DocumentRulerHandle
          kind="page-right-margin"
          label="右页边距"
          maximum={60}
          minimum={5}
          position={pageWidth - marginRight}
          value={layout.margins.right}
          valueText={`${formatRulerValue(layout.margins.right)} 毫米`}
          onKeyboardChange={updateFromKeyboard}
          onPointerChange={updateFromPointer}
        />
        <DocumentRulerHandle
          kind="left-indent"
          label="左缩进"
          maximum={maximumLeftIndent}
          minimum={0}
          position={leftIndentPosition}
          value={indent.left}
          valueText={formatParagraphIndentValue(indent.left)}
          onKeyboardChange={updateFromKeyboard}
          onPointerChange={updateFromPointer}
        />
        <DocumentRulerHandle
          kind="first-line-indent"
          label="首行/悬挂缩进"
          maximum={maximumFirstLineIndent}
          minimum={-indent.left}
          position={firstLineIndentPosition}
          value={indent.firstLine}
          valueText={formatFirstLineIndentValue(indent.firstLine)}
          onKeyboardChange={updateFromKeyboard}
          onPointerChange={updateFromPointer}
        />
        <DocumentRulerHandle
          kind="right-indent"
          label="右缩进"
          maximum={maximumRightIndent}
          minimum={0}
          position={rightIndentPosition}
          value={indent.right}
          valueText={formatParagraphIndentValue(indent.right)}
          onKeyboardChange={updateFromKeyboard}
          onPointerChange={updateFromPointer}
        />
      </fieldset>
    </div>
  );
}

function DocumentRulerHandle({
  kind,
  label,
  maximum,
  minimum,
  position,
  value,
  valueText,
  onKeyboardChange,
  onPointerChange,
}: {
  kind: DocumentRulerHandleKind;
  label: string;
  maximum: number;
  minimum: number;
  position: number;
  value: number;
  valueText: string;
  onKeyboardChange: (
    kind: DocumentRulerHandleKind,
    event: KeyboardEvent<HTMLButtonElement>,
  ) => void;
  onPointerChange: (kind: DocumentRulerHandleKind, clientX: number) => void;
}) {
  const activePointer = useRef<number | null>(null);
  const startPointer = (event: PointerEvent<HTMLButtonElement>): void => {
    if (event.button !== 0) return;
    event.preventDefault();
    activePointer.current = event.pointerId;
    event.currentTarget.setPointerCapture?.(event.pointerId);
    onPointerChange(kind, event.clientX);
  };
  const movePointer = (event: PointerEvent<HTMLButtonElement>): void => {
    if (activePointer.current !== event.pointerId) return;
    event.preventDefault();
    onPointerChange(kind, event.clientX);
  };
  const finishPointer = (event: PointerEvent<HTMLButtonElement>): void => {
    if (activePointer.current !== event.pointerId) return;
    event.preventDefault();
    onPointerChange(kind, event.clientX);
    activePointer.current = null;
    if (event.currentTarget.hasPointerCapture?.(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId);
    }
  };

  return (
    <button
      type="button"
      role="slider"
      className={`work-document-ruler-handle ${kind}`}
      aria-label={label}
      aria-orientation="horizontal"
      aria-valuemax={Math.round(maximum)}
      aria-valuemin={Math.round(minimum)}
      aria-valuenow={Math.round(value)}
      aria-valuetext={valueText}
      title={`${label}：${valueText}`}
      style={
        {
          '--work-document-ruler-handle-position': `${position}px`,
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

export function documentRulerPageWidth(
  layout: WorkDocumentSectionLayout,
): number {
  const portrait = documentRulerPortraitPage(layout);
  return layout.orientation === 'landscape' ? portrait.height : portrait.width;
}

export function documentRulerPageHeight(
  layout: WorkDocumentSectionLayout,
): number {
  const portrait = documentRulerPortraitPage(layout);
  return layout.orientation === 'landscape' ? portrait.width : portrait.height;
}

function documentRulerPortraitPage(layout: WorkDocumentSectionLayout): {
  width: number;
  height: number;
} {
  return layout.pageSize === 'letter'
    ? { width: 816, height: 1056 }
    : {
        width: millimetersToPixels(210),
        height: millimetersToPixels(297),
      };
}

export function documentRulerMaximumIndentLevel(
  layout: WorkDocumentSectionLayout,
): number {
  const textWidth =
    documentRulerPageWidth(layout) -
    millimetersToPixels(layout.margins.left + layout.margins.right);
  return Math.min(
    MAX_DOCUMENT_INDENT_LEVEL,
    Math.max(
      0,
      Math.floor(
        (Math.max(0, textWidth) - MINIMUM_TEXT_WIDTH_PX) /
          DOCUMENT_INDENT_STEP_PX,
      ),
    ),
  );
}

function documentRulerTicks(
  layout: WorkDocumentSectionLayout,
  pageWidth: number,
): DocumentRulerTick[] {
  const left = millimetersToPixels(layout.margins.left);
  const textWidth = Math.max(
    0,
    pageWidth - left - millimetersToPixels(layout.margins.right),
  );
  const textWidthMillimeters = pixelsToMillimeters(textWidth);
  const ticks: DocumentRulerTick[] = [];
  for (
    let distance = 0;
    distance <= textWidthMillimeters + 0.01;
    distance += 5
  ) {
    ticks.push({
      label: distance > 0 && distance % 10 === 0 ? String(distance / 10) : null,
      position: left + millimetersToPixels(distance),
    });
  }
  return ticks;
}

function constrainRulerParagraphIndent(
  value: DocumentParagraphIndent,
  textWidth: number,
): DocumentParagraphIndent {
  const normalized = normalizeDocumentParagraphIndent(value);
  const right = Math.min(
    normalized.right,
    Math.max(0, textWidth - MINIMUM_TEXT_WIDTH_PX),
  );
  const left = Math.min(
    normalized.left,
    Math.max(
      0,
      textWidth -
        right -
        Math.max(0, normalized.firstLine) -
        MINIMUM_TEXT_WIDTH_PX,
    ),
  );
  const firstLine = Math.max(
    -left,
    Math.min(
      normalized.firstLine,
      Math.max(0, textWidth - left - right - MINIMUM_TEXT_WIDTH_PX),
    ),
  );
  return { left, right, firstLine };
}

function maximumLeftIndentPixels(
  indent: DocumentParagraphIndent,
  textWidth: number,
): number {
  return Math.max(
    0,
    textWidth -
      indent.right -
      Math.max(0, indent.firstLine) -
      MINIMUM_TEXT_WIDTH_PX,
  );
}

function maximumRightIndentPixels(
  indent: DocumentParagraphIndent,
  textWidth: number,
): number {
  return Math.max(
    0,
    textWidth -
      indent.left -
      Math.max(0, indent.firstLine) -
      MINIMUM_TEXT_WIDTH_PX,
  );
}

function keyboardIndentValue(
  current: number,
  minimum: number,
  maximum: number,
  direction: number,
  event: KeyboardEvent<HTMLButtonElement>,
): number {
  if (event.key === 'Home') return minimum;
  if (event.key === 'End') return maximum;
  const step = event.shiftKey
    ? DOCUMENT_INDENT_STEP_PX
    : DOCUMENT_RULER_INDENT_STEP_PX;
  return Math.min(maximum, Math.max(minimum, current + direction * step));
}

function snapIndentPosition(value: number): number {
  return (
    Math.round(value / DOCUMENT_RULER_INDENT_STEP_PX) *
    DOCUMENT_RULER_INDENT_STEP_PX
  );
}

function pixelsToMillimeters(value: number): number {
  return (value * 25.4) / 96;
}

function formatRulerValue(value: number): string {
  return Number.isInteger(value) ? String(value) : value.toFixed(1);
}

function formatParagraphIndentValue(value: number): string {
  return value ? `${formatCentimeters(value)} 厘米` : '无缩进';
}

function formatFirstLineIndentValue(value: number): string {
  if (value > 0) return `首行缩进 ${formatCentimeters(value)} 厘米`;
  if (value < 0) return `悬挂缩进 ${formatCentimeters(-value)} 厘米`;
  return '无首行缩进';
}

function formatCentimeters(value: number): string {
  return (pixelsToMillimeters(value) / 10).toFixed(2);
}
