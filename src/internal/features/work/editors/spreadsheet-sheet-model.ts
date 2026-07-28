import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../work-types';

export type SpreadsheetSheetMoveDirection = -1 | 1;

export function addSpreadsheetSheet(
  content: WorkSpreadsheetContent,
): WorkSpreadsheetContent {
  const identity = nextSpreadsheetSheetIdentity(content.sheets);
  const nextSheet: WorkSpreadsheetSheet = {
    id: identity.id,
    name: identity.name,
    order: content.sheets.length,
    status: 1,
    hide: 0,
    row: 60,
    column: 26,
    defaultRowHeight: 24,
    defaultColWidth: 96,
    data: emptySpreadsheetData(60, 26),
    celldata: [],
    luckysheet_select_save: [
      {
        row: [0, 0],
        column: [0, 0],
        row_focus: 0,
        column_focus: 0,
      },
    ],
  };
  return withSpreadsheetSheets(content, [
    ...content.sheets.map((sheet) => ({ ...sheet, status: 0 })),
    nextSheet,
  ]);
}

function emptySpreadsheetData(
  rows: number,
  columns: number,
): NonNullable<WorkSpreadsheetSheet['data']> {
  return Array.from({ length: rows }, () =>
    Array.from({ length: columns }, () => null),
  );
}

export function activateSpreadsheetSheet(
  content: WorkSpreadsheetContent,
  sheetId: string,
): WorkSpreadsheetContent | null {
  const target = content.sheets.find((sheet) => sheet.id === sheetId);
  if (!target) return null;
  if (target.status === 1 && !isSpreadsheetSheetHidden(target)) return null;
  return withSpreadsheetSheets(
    content,
    content.sheets.map((sheet) => ({
      ...sheet,
      hide: sheet.id === sheetId ? 0 : sheet.hide,
      status: sheet.id === sheetId ? 1 : 0,
    })),
  );
}

export function deleteSpreadsheetSheet(
  content: WorkSpreadsheetContent,
  sheetId: string,
): WorkSpreadsheetContent | null {
  const targetIndex = content.sheets.findIndex((sheet) => sheet.id === sheetId);
  if (targetIndex < 0 || content.sheets.length <= 1) return null;
  const targetWasActive = content.sheets[targetIndex]?.status === 1;
  const remaining = content.sheets
    .filter((sheet) => sheet.id !== sheetId)
    .map((sheet) => ({ ...sheet }));
  if (targetWasActive || !remaining.some((sheet) => sheet.status === 1)) {
    const nextActive =
      remaining.find(
        (sheet) =>
          !isSpreadsheetSheetHidden(sheet) &&
          (sheet.order ?? 0) >=
            (content.sheets[targetIndex]?.order ?? targetIndex),
      ) ??
      [...remaining]
        .reverse()
        .find((sheet) => !isSpreadsheetSheetHidden(sheet)) ??
      remaining[0];
    for (const sheet of remaining) {
      sheet.status = sheet === nextActive ? 1 : 0;
      if (sheet === nextActive) sheet.hide = 0;
    }
  }
  return withSpreadsheetSheets(content, remaining);
}

export function duplicateSpreadsheetSheet(
  content: WorkSpreadsheetContent,
  sheetId: string,
): WorkSpreadsheetContent | null {
  const source = content.sheets.find((sheet) => sheet.id === sheetId);
  if (!source) return null;
  const identity = nextSpreadsheetSheetIdentity(content.sheets);
  const copyName = uniqueSpreadsheetSheetName(
    content.sheets,
    `${source.name} 副本`,
  );
  const duplicate: WorkSpreadsheetSheet = {
    ...structuredClone(source),
    id: identity.id,
    name: copyName,
    order: content.sheets.length,
    status: 1,
    hide: 0,
  };
  return withSpreadsheetSheets(content, [
    ...content.sheets.map((sheet) => ({ ...sheet, status: 0 })),
    duplicate,
  ]);
}

export function renameSpreadsheetSheet(
  content: WorkSpreadsheetContent,
  sheetId: string,
  name: string,
): WorkSpreadsheetContent | null {
  const normalized = name.trim();
  const target = content.sheets.find((sheet) => sheet.id === sheetId);
  if (
    !target ||
    !normalized ||
    normalized.length > 31 ||
    /[\\/?*[\]:]/u.test(normalized) ||
    target.name === normalized ||
    content.sheets.some(
      (sheet) =>
        sheet.id !== sheetId &&
        sheet.name.localeCompare(normalized, undefined, {
          sensitivity: 'accent',
        }) === 0,
    )
  ) {
    return null;
  }
  return withSpreadsheetSheets(
    content,
    content.sheets.map((sheet) =>
      sheet.id === sheetId ? { ...sheet, name: normalized } : sheet,
    ),
  );
}

export function setSpreadsheetSheetColor(
  content: WorkSpreadsheetContent,
  sheetId: string,
  color: string | null,
): WorkSpreadsheetContent | null {
  const normalized = color?.trim().toLowerCase() || undefined;
  if (normalized && !/^#[0-9a-f]{6}$/u.test(normalized)) return null;
  const target = content.sheets.find((sheet) => sheet.id === sheetId);
  if (!target || target.color?.toLowerCase() === normalized) return null;
  return withSpreadsheetSheets(
    content,
    content.sheets.map((sheet) =>
      sheet.id === sheetId ? { ...sheet, color: normalized } : sheet,
    ),
  );
}

export function hideSpreadsheetSheet(
  content: WorkSpreadsheetContent,
  sheetId: string,
  hidden: boolean,
): WorkSpreadsheetContent | null {
  const target = content.sheets.find((sheet) => sheet.id === sheetId);
  if (!target || isSpreadsheetSheetHidden(target) === hidden) return null;
  const visible = content.sheets.filter(
    (sheet) => !isSpreadsheetSheetHidden(sheet),
  );
  if (hidden && visible.length <= 1) return null;
  const nextActiveId =
    hidden && target.status === 1
      ? adjacentSpreadsheetSheetId(content, sheetId, 1)
      : null;
  return withSpreadsheetSheets(
    content,
    content.sheets.map((sheet) => ({
      ...sheet,
      hide:
        sheet.id === sheetId
          ? hidden
            ? 1
            : 0
          : nextActiveId && sheet.id === nextActiveId
            ? 0
            : sheet.hide,
      status:
        sheet.id === sheetId && hidden
          ? 0
          : nextActiveId && sheet.id === nextActiveId
            ? 1
            : nextActiveId
              ? 0
              : sheet.status,
    })),
  );
}

export function moveSpreadsheetSheet(
  content: WorkSpreadsheetContent,
  sheetId: string,
  direction: SpreadsheetSheetMoveDirection,
): WorkSpreadsheetContent | null {
  const sheets = orderedSpreadsheetSheets(content.sheets);
  const visibleIndexes = sheets.flatMap((sheet, index) =>
    isSpreadsheetSheetHidden(sheet) ? [] : [index],
  );
  const index = sheets.findIndex((sheet) => sheet.id === sheetId);
  const visibleIndex = visibleIndexes.indexOf(index);
  const targetIndex = visibleIndexes[visibleIndex + direction];
  if (
    index < 0 ||
    visibleIndex < 0 ||
    targetIndex === undefined ||
    targetIndex < 0
  ) {
    return null;
  }
  const reordered = [...sheets];
  [reordered[index], reordered[targetIndex]] = [
    reordered[targetIndex],
    reordered[index],
  ];
  return withSpreadsheetSheets(content, reordered);
}

export function adjacentSpreadsheetSheetId(
  content: WorkSpreadsheetContent,
  activeSheetId: string,
  direction: SpreadsheetSheetMoveDirection,
): string | null {
  const visible = orderedSpreadsheetSheets(content.sheets).filter(
    (sheet) => !isSpreadsheetSheetHidden(sheet) && sheet.id,
  );
  if (!visible.length) return null;
  const index = visible.findIndex((sheet) => sheet.id === activeSheetId);
  if (index < 0) return visible[0]?.id ?? null;
  const next = (index + direction + visible.length) % visible.length;
  return visible[next]?.id ?? null;
}

export function isSpreadsheetSheetHidden(sheet: WorkSpreadsheetSheet): boolean {
  return sheet.hide === 1;
}

function withSpreadsheetSheets(
  content: WorkSpreadsheetContent,
  sheets: WorkSpreadsheetSheet[],
): WorkSpreadsheetContent {
  return {
    ...content,
    sheets: sheets.map((sheet, order) => ({ ...sheet, order })),
  };
}

function orderedSpreadsheetSheets(
  sheets: WorkSpreadsheetSheet[],
): WorkSpreadsheetSheet[] {
  return sheets
    .map((sheet, index) => ({ sheet, index }))
    .sort(
      (left, right) =>
        (left.sheet.order ?? left.index) - (right.sheet.order ?? right.index) ||
        left.index - right.index,
    )
    .map(({ sheet }) => sheet);
}

function nextSpreadsheetSheetIdentity(sheets: WorkSpreadsheetSheet[]): {
  id: string;
  name: string;
} {
  const ids = new Set(sheets.map((sheet) => sheet.id));
  const names = new Set(sheets.map((sheet) => sheet.name));
  let ordinal = sheets.length + 1;
  while (ids.has(`sheet-${ordinal}`) || names.has(`工作表 ${ordinal}`)) {
    ordinal += 1;
  }
  return { id: `sheet-${ordinal}`, name: `工作表 ${ordinal}` };
}

function uniqueSpreadsheetSheetName(
  sheets: WorkSpreadsheetSheet[],
  preferred: string,
): string {
  const names = new Set(sheets.map((sheet) => sheet.name));
  if (!names.has(preferred)) return preferred;
  let suffix = 2;
  while (names.has(`${preferred} ${suffix}`)) suffix += 1;
  return `${preferred} ${suffix}`;
}
