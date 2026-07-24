import type { TabStopDefinition } from 'docx';
import {
  type DocumentTabAlignment,
  type DocumentTabLeader,
  normalizeDocumentTabStops,
} from './work-document-tab-stops';

export function paragraphAlignment(
  element: HTMLElement,
  docx: typeof import('docx'),
) {
  const alignment = element.style.textAlign;
  if (alignment === 'center') return docx.AlignmentType.CENTER;
  if (alignment === 'right' || alignment === 'end')
    return docx.AlignmentType.RIGHT;
  if (alignment === 'justify') return docx.AlignmentType.JUSTIFIED;
  if (alignment === 'left' || alignment === 'start')
    return docx.AlignmentType.LEFT;
  return undefined;
}

export function paragraphDirectionOptions(element: HTMLElement): {
  bidirectional?: boolean;
} {
  const bidirectional = paragraphBidirectional(element);
  return bidirectional === undefined ? {} : { bidirectional };
}

export function paragraphBidirectional(
  element: HTMLElement,
): boolean | undefined {
  const direction = domDirection(element);
  return direction === undefined ? undefined : direction === 'rtl';
}

export function domDirection(element: HTMLElement): 'ltr' | 'rtl' | undefined {
  let current: HTMLElement | null = element;
  while (current) {
    const attribute = current.getAttribute('dir')?.trim().toLowerCase();
    if (attribute === 'ltr' || attribute === 'rtl') return attribute;
    if (attribute === 'auto') return undefined;
    const css = current.style.direction.trim().toLowerCase();
    if (css === 'ltr' || css === 'rtl') return css;
    current = current.parentElement;
  }
  return undefined;
}

function paragraphLineSpacing(element: HTMLElement): number | undefined {
  const value = element.style.lineHeight.trim();
  if (!value || value === 'normal') return undefined;
  const percentage = /^(\d+(?:\.\d+)?)%$/.exec(value);
  if (percentage) return Math.round((Number(percentage[1]) / 100) * 240);
  const unitless = /^(\d+(?:\.\d+)?)$/.exec(value);
  if (unitless) return Math.round(Number(unitless[1]) * 240);
  const points = cssLengthToPoints(value);
  return points ? Math.round(points * 20) : undefined;
}

export function paragraphSpacingOptions(
  element: HTMLElement,
  heading: boolean,
  docx: typeof import('docx'),
) {
  const before =
    dataPointLengthToTwips(element.dataset.officeSpaceBefore) ??
    cssSpacingLengthToTwips(element.style.marginTop);
  const after =
    dataPointLengthToTwips(element.dataset.officeSpaceAfter) ??
    cssSpacingLengthToTwips(element.style.marginBottom);
  const lineRule = paragraphLineRule(element, docx);
  return {
    after: after ?? (heading ? 180 : 120),
    ...(before !== undefined ? { before } : {}),
    line: paragraphLineSpacing(element) ?? 320,
    ...(lineRule ? { lineRule } : {}),
  };
}

function paragraphLineRule(element: HTMLElement, docx: typeof import('docx')) {
  if (element.dataset.officeLineRule === 'auto') return docx.LineRuleType.AUTO;
  if (element.dataset.officeLineRule === 'exact')
    return docx.LineRuleType.EXACT;
  if (element.dataset.officeLineRule === 'atLeast')
    return docx.LineRuleType.AT_LEAST;
  return undefined;
}

export function paragraphIndent(
  element: HTMLElement,
  tag: string,
):
  | {
      left?: number;
      right?: number;
      hanging?: number;
      firstLine?: number;
    }
  | undefined {
  const blockquoteIndent = tag === 'blockquote' ? 540 : 0;
  const level = Number(element.dataset.officeIndentLevel);
  const explicitIndent =
    Number.isFinite(level) && level > 0
      ? Math.round(level * 360)
      : cssLengthToTwips(element.style.marginLeft);
  const left = blockquoteIndent + explicitIndent;
  const right =
    dataPixelLengthToTwips(element.dataset.officeIndentRight) ??
    cssLengthToTwips(element.style.marginRight);
  const firstLine =
    dataPixelLengthToTwips(element.dataset.officeIndentFirstLine, true) ??
    cssSignedLengthToTwips(element.style.textIndent);
  if (!left && !right && !firstLine) return undefined;
  return {
    ...(left > 0 ? { left } : {}),
    ...(right > 0 ? { right } : {}),
    ...(firstLine > 0 ? { firstLine } : {}),
    ...(firstLine < 0 ? { hanging: Math.abs(firstLine) } : {}),
  };
}

export function paragraphPaginationOptions(element: HTMLElement) {
  return {
    keepLines: dataBoolean(element.dataset.officeKeepLines),
    keepNext: dataBoolean(element.dataset.officeKeepWithNext),
    pageBreakBefore: dataBoolean(element.dataset.officePageBreakBefore),
    widowControl: dataBoolean(element.dataset.officeWidowControl),
  };
}

export function paragraphTabStops(
  element: HTMLElement,
  docx: typeof import('docx'),
): TabStopDefinition[] | undefined {
  const tabStops = normalizeDocumentTabStops(
    element.dataset.officeTabStops ?? '',
  ).map((tabStop) => ({
    type: docxTabStopType(tabStop.alignment, docx),
    position: Math.round(tabStop.position * 15),
    ...(tabStop.leader === 'none'
      ? {}
      : { leader: docxTabLeader(tabStop.leader, docx) }),
  }));
  return tabStops.length ? tabStops : undefined;
}

function docxTabStopType(
  alignment: DocumentTabAlignment,
  docx: typeof import('docx'),
): TabStopDefinition['type'] {
  if (alignment === 'center') return docx.TabStopType.CENTER;
  if (alignment === 'right') return docx.TabStopType.RIGHT;
  if (alignment === 'decimal') return docx.TabStopType.DECIMAL;
  return docx.TabStopType.LEFT;
}

function docxTabLeader(
  leader: DocumentTabLeader,
  docx: typeof import('docx'),
): NonNullable<TabStopDefinition['leader']> {
  if (leader === 'dot') return docx.LeaderType.DOT;
  if (leader === 'hyphen') return docx.LeaderType.HYPHEN;
  if (leader === 'underscore') return docx.LeaderType.UNDERSCORE;
  if (leader === 'middleDot') return docx.LeaderType.MIDDLE_DOT;
  return docx.LeaderType.NONE;
}

export function dataBoolean(value: string | undefined): boolean | undefined {
  if (value === 'true' || value === '1') return true;
  if (value === 'false' || value === '0') return false;
  return undefined;
}

export function cssFontFamily(value: string): string | undefined {
  const family = value
    .split(',')[0]
    ?.trim()
    .replace(/^(['"])(.*)\1$/, '$2');
  return family || undefined;
}

export function cssFontSize(value: string): number | undefined {
  const points = cssLengthToPoints(value);
  return points ? Math.max(1, Math.round(points * 2)) : undefined;
}

function cssLengthToTwips(value: string): number {
  const points = cssLengthToPoints(value);
  return points ? Math.max(0, Math.round(points * 20)) : 0;
}

function cssSignedLengthToTwips(value: string): number {
  const match = /^(-?\d+(?:\.\d+)?)(px|pt|in|cm|mm)$/i.exec(value.trim());
  if (!match) return 0;
  const amount = Number(match[1]);
  if (!Number.isFinite(amount) || !amount) return 0;
  const unit = match[2]?.toLowerCase();
  const points =
    unit === 'px'
      ? amount * 0.75
      : unit === 'in'
        ? amount * 72
        : unit === 'cm'
          ? (amount * 72) / 2.54
          : unit === 'mm'
            ? (amount * 72) / 25.4
            : amount;
  return Math.round(points * 20);
}

function dataPixelLengthToTwips(
  value: string | undefined,
  signed = false,
): number | undefined {
  if (value === undefined) return undefined;
  const pixels = Number(value);
  if (!Number.isFinite(pixels)) return undefined;
  const twips = Math.round(pixels * 15);
  return signed ? twips : Math.max(0, twips);
}

function dataPointLengthToTwips(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  const points = Number(value);
  return Number.isFinite(points) && points >= 0
    ? Math.round(points * 20)
    : undefined;
}

function cssSpacingLengthToTwips(value: string): number | undefined {
  if (!value.trim()) return undefined;
  if (/^0(?:\.0+)?(?:px|pt|in|cm|mm)?$/i.test(value.trim())) return 0;
  const points = cssLengthToPoints(value);
  return points === undefined ? undefined : Math.round(points * 20);
}

function cssLengthToPoints(value: string): number | undefined {
  const match = /^(-?\d+(?:\.\d+)?)(px|pt|in|cm|mm)$/i.exec(value.trim());
  if (!match) return undefined;
  const amount = Number(match[1]);
  if (!Number.isFinite(amount) || amount <= 0) return undefined;
  const unit = match[2]?.toLowerCase();
  if (unit === 'px') return amount * 0.75;
  if (unit === 'in') return amount * 72;
  if (unit === 'cm') return (amount * 72) / 2.54;
  if (unit === 'mm') return (amount * 72) / 25.4;
  return amount;
}

export function cssColorToHex(source: string): string | undefined {
  const value = source.trim();
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(value)?.[1];
  if (hex) {
    return (
      hex.length === 3
        ? [...hex].map((character) => character.repeat(2)).join('')
        : hex
    ).toUpperCase();
  }
  const rgb =
    /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})(?:\s*,[^)]*)?\)$/i.exec(
      value,
    );
  if (!rgb) return undefined;
  return rgb
    .slice(1, 4)
    .map((channel) =>
      Math.min(255, Number(channel)).toString(16).padStart(2, '0'),
    )
    .join('')
    .toUpperCase();
}
