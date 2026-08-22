export interface WorkSpreadsheetDiagonalBorderLine {
  color: string;
  style: string;
}

export interface WorkSpreadsheetDiagonalBorder {
  down: boolean;
  line: WorkSpreadsheetDiagonalBorderLine;
  up: boolean;
}

type UnknownRecord = Record<string, unknown>;

const A3S_DIAGONAL_BORDER_KEY = 'a3sDiagonal';

export function spreadsheetDiagonalBorderFromCellValue(
  value: unknown,
): WorkSpreadsheetDiagonalBorder | null | undefined {
  if (!isRecord(value)) return undefined;
  const metadata = value[A3S_DIAGONAL_BORDER_KEY];
  if (isRecord(metadata)) {
    const up = metadata.up;
    const down = metadata.down;
    const line = spreadsheetDiagonalBorderLine(metadata.line);
    if (typeof up === 'boolean' && typeof down === 'boolean') {
      if (!up && !down) return null;
      if (line) return { down, line, up };
    }
  }

  if (value.s === null) return null;
  const legacyLine = spreadsheetDiagonalBorderLine(value.s);
  return legacyLine ? { down: true, line: legacyLine, up: false } : undefined;
}

export function spreadsheetCellValueWithDiagonalBorder(
  value: UnknownRecord,
  border: WorkSpreadsheetDiagonalBorder | null,
): UnknownRecord {
  const next = { ...value };
  delete next.s;
  delete next[A3S_DIAGONAL_BORDER_KEY];
  if (!border) return next;
  const line = { ...border.line };
  next[A3S_DIAGONAL_BORDER_KEY] = {
    down: border.down,
    line,
    up: border.up,
  };
  if (border.down) next.s = line;
  return next;
}

function spreadsheetDiagonalBorderLine(
  value: unknown,
): WorkSpreadsheetDiagonalBorderLine | null {
  if (!isRecord(value)) return null;
  return typeof value.color === 'string' && typeof value.style === 'string'
    ? { color: value.color, style: value.style }
    : null;
}

function isRecord(value: unknown): value is UnknownRecord {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value));
}
