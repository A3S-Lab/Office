import type { Cell } from '@fortune-sheet/core';
import { cloneSparseMatrix } from '../spreadsheet-sparse';
import { spreadsheetPivotOutputContains } from '../work-spreadsheet-pivots';
import { sheetProtectionAuthority } from '../work-spreadsheet-protection';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../work-types';
import { resolveSpreadsheetGoToTarget } from './spreadsheet-go-to';
import { spreadsheetSheetBounds } from './spreadsheet-keyboard-navigation';

const MAXIMUM_SPREADSHEET_ROWS = 1_048_576;
const MAXIMUM_SPREADSHEET_COLUMNS = 16_384;
const MAXIMUM_SPREADSHEET_HYPERLINK_ADDRESS_LENGTH = 2_048;
const MAXIMUM_SPREADSHEET_CELL_TEXT_LENGTH = 32_767;

export type SpreadsheetHyperlinkType = 'webpage' | 'cellrange' | 'sheet';

export interface SpreadsheetHyperlinkTarget {
  linkType: SpreadsheetHyperlinkType;
  linkAddress: string;
}

export interface SpreadsheetHyperlinkCell {
  sheetId: string;
  row: number;
  column: number;
}

export interface SpreadsheetHyperlinkRequest extends SpreadsheetHyperlinkCell {
  linkType: SpreadsheetHyperlinkType;
  linkAddress: string;
  displayText?: string;
}

export type SpreadsheetHyperlinkDialogValue = Pick<
  SpreadsheetHyperlinkRequest,
  'linkType' | 'linkAddress' | 'displayText'
>;

export type SpreadsheetHyperlinkErrorCode =
  | 'empty-address'
  | 'formula-display-text'
  | 'invalid-cell-range'
  | 'invalid-display-text'
  | 'invalid-web-address'
  | 'pivot-cell'
  | 'protected-cell'
  | 'source-out-of-bounds'
  | 'source-sheet-not-found'
  | 'target-out-of-bounds'
  | 'target-sheet-hidden'
  | 'target-sheet-not-found'
  | 'unsupported-link-type';

export type SpreadsheetHyperlinkValidation =
  | {
      ok: true;
      target: SpreadsheetHyperlinkTarget;
      displayText?: string;
    }
  | {
      ok: false;
      code: SpreadsheetHyperlinkErrorCode;
      message: string;
    };

export interface SpreadsheetHyperlinkDialogSource
  extends SpreadsheetHyperlinkCell {
  sheetName: string;
  cellReference: string;
  displayText: string;
  displayTextEditable: boolean;
  hasHyperlink: boolean;
  link: SpreadsheetHyperlinkTarget | null;
  sheetOptions: Array<{ id: string; name: string }>;
}

export function validateSpreadsheetHyperlinkRequest(
  content: WorkSpreadsheetContent,
  request: SpreadsheetHyperlinkRequest,
): SpreadsheetHyperlinkValidation {
  const source = validateSpreadsheetHyperlinkCell(content, request);
  if (!source.ok) return source;
  const target = normalizeSpreadsheetHyperlinkTarget(
    content,
    request.sheetId,
    request.linkType,
    request.linkAddress,
  );
  if (!target.ok) return target;
  if (request.displayText === undefined) return target;
  const displayText = request.displayText.trim();
  if (
    !displayText ||
    displayText.length > MAXIMUM_SPREADSHEET_CELL_TEXT_LENGTH
  ) {
    return hyperlinkError('invalid-display-text');
  }
  const cell = spreadsheetHyperlinkCellAt(
    source.sheet,
    request.row,
    request.column,
  );
  if (typeof cell?.f === 'string' && cell.f.trim()) {
    return hyperlinkError('formula-display-text');
  }
  return { ...target, displayText };
}

export function applySpreadsheetHyperlink(
  content: WorkSpreadsheetContent,
  request: SpreadsheetHyperlinkRequest,
): WorkSpreadsheetContent | null {
  const validation = validateSpreadsheetHyperlinkRequest(content, request);
  if (!validation.ok) return null;
  const sheetIndex = content.sheets.findIndex(
    (sheet) => sheet.id === request.sheetId,
  );
  const sheet = content.sheets[sheetIndex];
  if (!sheet || sheetIndex < 0) return null;

  const cell = {
    ...(spreadsheetHyperlinkCellAt(sheet, request.row, request.column) ?? {}),
    hl: { r: request.row, c: request.column, id: request.sheetId },
  } satisfies Cell;
  if (validation.displayText !== undefined) {
    cell.v = validation.displayText;
    cell.m = validation.displayText;
    if (cell.f !== undefined) delete cell.f;
    if (cell.ct) {
      const cellType = { ...cell.ct, t: 's' };
      if ('s' in cellType) delete cellType.s;
      cell.ct = cellType;
    }
  }
  const key = spreadsheetHyperlinkKey(request.row, request.column);
  const nextSheet = setSpreadsheetHyperlinkCell(sheet, request, cell);
  nextSheet.hyperlink = {
    ...(sheet.hyperlink ?? {}),
    [key]: validation.target,
  };

  const sheets = [...content.sheets];
  sheets[sheetIndex] = nextSheet;
  return { ...content, sheets };
}

export function removeSpreadsheetHyperlink(
  content: WorkSpreadsheetContent,
  target: SpreadsheetHyperlinkCell,
): WorkSpreadsheetContent | null {
  if (!canRemoveSpreadsheetHyperlink(content, target)) return null;
  const sheetIndex = content.sheets.findIndex(
    (sheet) => sheet.id === target.sheetId,
  );
  const sheet = content.sheets[sheetIndex];
  if (!sheet || sheetIndex < 0) return null;
  const key = spreadsheetHyperlinkKey(target.row, target.column);
  const sourceCell = spreadsheetHyperlinkCellAt(
    sheet,
    target.row,
    target.column,
  );
  const cell = sourceCell ? { ...sourceCell } : null;
  if (cell?.hl !== undefined) delete cell.hl;
  const nextSheet = setSpreadsheetHyperlinkCell(
    sheet,
    target,
    cell && Object.keys(cell).length ? cell : null,
  );
  const hyperlink = { ...(sheet.hyperlink ?? {}) };
  delete hyperlink[key];
  if (Object.keys(hyperlink).length) nextSheet.hyperlink = hyperlink;
  else if (nextSheet.hyperlink !== undefined) delete nextSheet.hyperlink;

  const sheets = [...content.sheets];
  sheets[sheetIndex] = nextSheet;
  return { ...content, sheets };
}

export function canRemoveSpreadsheetHyperlink(
  content: WorkSpreadsheetContent,
  target: SpreadsheetHyperlinkCell,
): boolean {
  const source = validateSpreadsheetHyperlinkCell(content, target);
  if (!source.ok) return false;
  const key = spreadsheetHyperlinkKey(target.row, target.column);
  return Boolean(
    source.sheet.hyperlink?.[key] ||
      spreadsheetHyperlinkCellAt(source.sheet, target.row, target.column)?.hl,
  );
}

export function createSpreadsheetHyperlinkDialogSource(
  content: WorkSpreadsheetContent,
  target: SpreadsheetHyperlinkCell,
): SpreadsheetHyperlinkDialogSource | null {
  const validation = validateSpreadsheetHyperlinkCell(content, target);
  if (!validation.ok) return null;
  const sheet = validation.sheet;
  const cell = spreadsheetHyperlinkCellAt(sheet, target.row, target.column);
  const key = spreadsheetHyperlinkKey(target.row, target.column);
  const rawLink = sheet.hyperlink?.[key];
  const link = spreadsheetHyperlinkTarget(rawLink);
  const display = cell?.m ?? cell?.v;
  return {
    ...target,
    sheetName: sheet.name,
    cellReference: spreadsheetHyperlinkCellReference(target.row, target.column),
    displayText:
      display === null || display === undefined ? '' : String(display),
    displayTextEditable: !(typeof cell?.f === 'string' && cell.f.trim()),
    hasHyperlink: Boolean(rawLink || cell?.hl),
    link,
    sheetOptions: content.sheets.flatMap((candidate) =>
      candidate.id && candidate.hide !== 1
        ? [{ id: candidate.id, name: candidate.name }]
        : [],
    ),
  };
}

function validateSpreadsheetHyperlinkCell(
  content: WorkSpreadsheetContent,
  target: SpreadsheetHyperlinkCell,
):
  | { ok: true; sheet: WorkSpreadsheetSheet }
  | {
      ok: false;
      code: SpreadsheetHyperlinkErrorCode;
      message: string;
    } {
  const sheet = content.sheets.find(
    (candidate) => candidate.id === target.sheetId,
  );
  if (!sheet) return hyperlinkError('source-sheet-not-found');
  const bounds = spreadsheetSheetBounds(sheet);
  if (
    !Number.isSafeInteger(target.row) ||
    !Number.isSafeInteger(target.column) ||
    target.row < 0 ||
    target.column < 0 ||
    target.row >= MAXIMUM_SPREADSHEET_ROWS ||
    target.column >= MAXIMUM_SPREADSHEET_COLUMNS ||
    target.row > bounds.lastRow ||
    target.column > bounds.lastColumn
  ) {
    return hyperlinkError('source-out-of-bounds');
  }
  if (sheetProtectionAuthority(sheet).sheet === 1) {
    return hyperlinkError('protected-cell');
  }
  if (spreadsheetPivotOutputContains(sheet, target.row, target.column)) {
    return hyperlinkError('pivot-cell');
  }
  return { ok: true, sheet };
}

function normalizeSpreadsheetHyperlinkTarget(
  content: WorkSpreadsheetContent,
  sourceSheetId: string,
  linkType: SpreadsheetHyperlinkType,
  linkAddress: string,
): SpreadsheetHyperlinkValidation {
  if (!['webpage', 'cellrange', 'sheet'].includes(linkType)) {
    return hyperlinkError('unsupported-link-type');
  }
  const address = linkAddress.trim();
  if (!address) return hyperlinkError('empty-address');
  if (address.length > MAXIMUM_SPREADSHEET_HYPERLINK_ADDRESS_LENGTH) {
    return hyperlinkError(
      linkType === 'webpage' ? 'invalid-web-address' : 'invalid-cell-range',
    );
  }
  if (linkType === 'webpage') {
    const normalized = /^[a-z][a-z\d+.-]*:/i.test(address)
      ? address
      : `https://${address}`;
    try {
      const url = new URL(normalized);
      if (
        !['http:', 'https:'].includes(url.protocol) ||
        !url.hostname ||
        /\s/.test(normalized)
      ) {
        return hyperlinkError('invalid-web-address');
      }
    } catch {
      return hyperlinkError('invalid-web-address');
    }
    return {
      ok: true,
      target: { linkType: 'webpage', linkAddress: normalized },
    };
  }
  if (linkType === 'sheet') {
    const targetSheet = spreadsheetHyperlinkSheetByName(content, address);
    if (!targetSheet) return hyperlinkError('target-sheet-not-found');
    if (targetSheet.hide === 1) return hyperlinkError('target-sheet-hidden');
    return {
      ok: true,
      target: { linkType: 'sheet', linkAddress: targetSheet.name },
    };
  }

  const resolution = resolveSpreadsheetGoToTarget(
    content,
    sourceSheetId,
    address,
  );
  if (!resolution.ok) {
    if (resolution.code === 'sheet-not-found') {
      return hyperlinkError('target-sheet-not-found');
    }
    if (resolution.code === 'hidden-sheet') {
      return hyperlinkError('target-sheet-hidden');
    }
    if (resolution.code === 'out-of-bounds') {
      return hyperlinkError('target-out-of-bounds');
    }
    return hyperlinkError('invalid-cell-range');
  }
  if (resolution.target.source !== 'reference') {
    return hyperlinkError('invalid-cell-range');
  }
  return {
    ok: true,
    target: {
      linkType: 'cellrange',
      linkAddress: resolution.target.displayReference,
    },
  };
}

function setSpreadsheetHyperlinkCell(
  sheet: WorkSpreadsheetSheet,
  target: SpreadsheetHyperlinkCell,
  cell: Cell | null,
): WorkSpreadsheetSheet {
  if (sheet.data !== undefined) {
    const data = cloneSparseMatrix(sheet.data);
    const row = data[target.row] ?? [];
    data[target.row] = row;
    if (cell) row[target.column] = cell;
    else delete row[target.column];
    return { ...sheet, data };
  }

  let found = false;
  const celldata = (sheet.celldata ?? []).flatMap((entry) => {
    if (entry.r !== target.row || entry.c !== target.column) return [entry];
    if (found) return [entry];
    found = true;
    return cell ? [{ ...entry, v: cell }] : [];
  });
  if (!found && cell) {
    celldata.push({ r: target.row, c: target.column, v: cell });
    celldata.sort((left, right) => left.r - right.r || left.c - right.c);
  }
  return { ...sheet, celldata };
}

function spreadsheetHyperlinkCellAt(
  sheet: WorkSpreadsheetSheet,
  row: number,
  column: number,
): Cell | null {
  return (
    sheet.data?.[row]?.[column] ??
    sheet.celldata?.find((entry) => entry.r === row && entry.c === column)?.v ??
    null
  );
}

function spreadsheetHyperlinkTarget(
  value: unknown,
): SpreadsheetHyperlinkTarget | null {
  if (!value || typeof value !== 'object') return null;
  const link = value as Record<string, unknown>;
  return ['webpage', 'cellrange', 'sheet'].includes(String(link.linkType)) &&
    typeof link.linkAddress === 'string'
    ? {
        linkType: link.linkType as SpreadsheetHyperlinkType,
        linkAddress: link.linkAddress,
      }
    : null;
}

function spreadsheetHyperlinkSheetByName(
  content: WorkSpreadsheetContent,
  name: string,
): WorkSpreadsheetSheet | undefined {
  const normalized = name.trim().toLocaleLowerCase();
  return content.sheets.find(
    (sheet) => sheet.name.trim().toLocaleLowerCase() === normalized,
  );
}

function spreadsheetHyperlinkKey(row: number, column: number): string {
  return `${row}_${column}`;
}

function spreadsheetHyperlinkCellReference(
  row: number,
  column: number,
): string {
  let value = column + 1;
  let label = '';
  while (value > 0) {
    value -= 1;
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26);
  }
  return `${label}${row + 1}`;
}

function hyperlinkError(code: SpreadsheetHyperlinkErrorCode): {
  ok: false;
  code: SpreadsheetHyperlinkErrorCode;
  message: string;
} {
  const messages: Record<SpreadsheetHyperlinkErrorCode, string> = {
    'empty-address': '请输入超链接地址。',
    'formula-display-text': '公式单元格只能更新链接目标，不能替换显示文本。',
    'invalid-cell-range': '请输入一个有效的 A1 单元格或连续区域。',
    'invalid-display-text': '显示文本不能为空，且最多包含 32767 个字符。',
    'invalid-web-address': '请输入有效的 HTTP 或 HTTPS 地址。',
    'pivot-cell': '不能在数据透视表结果区域中插入超链接。',
    'protected-cell': '不能在受保护的工作表中更改超链接。',
    'source-out-of-bounds': '当前单元格超出了工作表的有效边界。',
    'source-sheet-not-found': '找不到当前工作表。',
    'target-out-of-bounds': '链接目标超出了目标工作表的有效边界。',
    'target-sheet-hidden': '不能链接到隐藏工作表。',
    'target-sheet-not-found': '找不到链接目标工作表。',
    'unsupported-link-type': '不支持此超链接类型。',
  };
  return { ok: false, code, message: messages[code] };
}
