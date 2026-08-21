import type { Cell } from '@fortune-sheet/core';

export const spreadsheetTextOrientationIds = [
  'horizontal',
  'angleCounterclockwise',
  'angleClockwise',
  'vertical',
  'rotateUp',
  'rotateDown',
] as const;

export type SpreadsheetTextOrientationId =
  (typeof spreadsheetTextOrientationIds)[number];

export type SpreadsheetTextOrientation =
  | { kind: 'rotation'; angle: number }
  | { kind: 'stacked' };

export type SpreadsheetTextOrientationCellStyle = Pick<Cell, 'rt' | 'tr'>;

const presetAngles: Readonly<
  Record<Exclude<SpreadsheetTextOrientationId, 'vertical'>, number>
> = {
  horizontal: 0,
  angleCounterclockwise: 45,
  angleClockwise: -45,
  rotateUp: 90,
  rotateDown: -90,
};

const legacyFortunePresetAngles: Readonly<Record<string, number>> = {
  '0': 0,
  '1': 45,
  '2': -45,
  '4': 90,
  '5': -90,
};

export function isSpreadsheetTextOrientationId(
  value: unknown,
): value is SpreadsheetTextOrientationId {
  return spreadsheetTextOrientationIds.includes(
    value as SpreadsheetTextOrientationId,
  );
}

export function spreadsheetTextOrientationFromChoice(
  choice: SpreadsheetTextOrientationId,
): SpreadsheetTextOrientation {
  return choice === 'vertical'
    ? { kind: 'stacked' }
    : { kind: 'rotation', angle: presetAngles[choice] };
}

export function spreadsheetTextOrientationFromAngle(
  value: unknown,
): SpreadsheetTextOrientation | null {
  const angle = integer(value);
  return angle !== null && angle >= -90 && angle <= 90
    ? { kind: 'rotation', angle }
    : null;
}

export function spreadsheetTextOrientationFromXlsx(
  value: unknown,
): SpreadsheetTextOrientation | null {
  const rotation = integer(value);
  if (rotation === 255) return { kind: 'stacked' };
  if (rotation === null || rotation < 0 || rotation > 180) return null;
  return {
    kind: 'rotation',
    angle: rotation <= 90 ? rotation : 90 - rotation,
  };
}

export function spreadsheetExplicitTextOrientationFromCell(
  cell: Pick<Cell, 'rt' | 'tr'> | null | undefined,
): SpreadsheetTextOrientation | null {
  if (cell?.rt !== undefined && cell.rt !== null) {
    const direct = spreadsheetTextOrientationFromXlsx(cell.rt);
    if (direct) return direct;
  }

  const legacy = String(cell?.tr ?? '').trim();
  if (!legacy) return null;
  if (legacy === '3' || legacy === '255') return { kind: 'stacked' };
  const presetAngle = legacyFortunePresetAngles[legacy];
  if (presetAngle !== undefined) {
    return { kind: 'rotation', angle: presetAngle };
  }
  const raw = integer(legacy);
  return raw !== null && raw >= 6
    ? spreadsheetTextOrientationFromXlsx(raw)
    : null;
}

export function spreadsheetTextOrientationFromCell(
  cell: Pick<Cell, 'rt' | 'tr'> | null | undefined,
): SpreadsheetTextOrientation {
  return (
    spreadsheetExplicitTextOrientationFromCell(cell) ?? {
      kind: 'rotation',
      angle: 0,
    }
  );
}

export function spreadsheetTextOrientationCellStyle(
  orientation: SpreadsheetTextOrientation | null | undefined,
): SpreadsheetTextOrientationCellStyle | null {
  if (!orientation) return null;
  if (orientation.kind === 'stacked') return { tr: '3' };
  return {
    rt: orientation.angle < 0 ? 90 - orientation.angle : orientation.angle,
  };
}

export function spreadsheetTextOrientationXlsxValue(
  orientation: SpreadsheetTextOrientation | null | undefined,
): number | null {
  if (!orientation) return null;
  if (orientation.kind === 'stacked') return 255;
  return orientation.angle < 0 ? 90 - orientation.angle : orientation.angle;
}

export function spreadsheetTextOrientationXlsxValueFromCell(
  cell: Pick<Cell, 'rt' | 'tr'> | null | undefined,
): number | null {
  return spreadsheetTextOrientationXlsxValue(
    spreadsheetExplicitTextOrientationFromCell(cell),
  );
}

export function spreadsheetVisibleTextRotationFromCell(
  cell: Pick<Cell, 'rt' | 'tr'> | null | undefined,
): number {
  const orientation = spreadsheetTextOrientationFromCell(cell);
  return orientation.kind === 'rotation' ? orientation.angle : 0;
}

export function spreadsheetTextOrientationChoiceFromCell(
  cell: Pick<Cell, 'rt' | 'tr'> | null | undefined,
): SpreadsheetTextOrientationId | null {
  const orientation = spreadsheetTextOrientationFromCell(cell);
  if (orientation.kind === 'stacked') return 'vertical';
  return (
    (Object.entries(presetAngles).find(
      ([, angle]) => angle === orientation.angle,
    )?.[0] as Exclude<SpreadsheetTextOrientationId, 'vertical'> | undefined) ??
    null
  );
}

function integer(value: unknown): number | null {
  if (typeof value === 'string' && !/^-?\d+$/.test(value.trim())) return null;
  const parsed = typeof value === 'number' ? value : Number(value);
  return Number.isSafeInteger(parsed) ? parsed : null;
}
