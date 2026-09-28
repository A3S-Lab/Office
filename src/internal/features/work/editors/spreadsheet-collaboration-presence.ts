import type { Cell, Presence } from '@fortune-sheet/core';
import type { WorkbookInstance } from '@fortune-sheet/react';
import { useEffect } from 'react';
import {
  paintTextCaretInBox,
  type FrameCaretPaint,
  type TextCaretBox,
} from '../../../collaboration/office-collaboration-frame-caret';
import type { WorkOfficeCollaborationParticipant } from '../../../collaboration/office-collaboration-presence';
import type { WorkSpreadsheetContent } from '../work-types';
import {
  officePresenceColor,
  useOfficeRemoteParticipants,
} from './office-collaboration-presence-ui';
import { spreadsheetCellAt } from './spreadsheet-editor-support';

export function useSpreadsheetCollaborationPresenceProjection({
  content,
  frameCell = null,
  workbook,
  /** Bumped after in-place sheet sync so Fortune re-draws overlays it cleared. */
  refreshKey = 0,
}: {
  content: WorkSpreadsheetContent;
  /** Cell named by the collaboration frame that just wrote it. */
  frameCell?: SpreadsheetFrameCell | null;
  workbook: WorkbookInstance | null;
  refreshKey?: number;
}): void {
  const participants = useOfficeRemoteParticipants();
  // Keyed on the projected presences, not on content identity: every content
  // update (e.g. a remote peer typing into a cell) would otherwise remove and
  // re-add the same presence and blink the remote cursor.
  const projectionKey = JSON.stringify([
    ...spreadsheetPresenceProjection(content, participants),
    ...spreadsheetWrittenCellProjection(content, frameCell),
  ]);

  useEffect(() => {
    if (!workbook) return;
    const projected = JSON.parse(projectionKey) as Presence[];
    if (projected.length) workbook.addPresences(projected);
    return () => {
      if (!projected.length) return;
      workbook.removePresences(
        projected.map(({ userId, username }) => ({ userId, username })),
      );
    };
  }, [projectionKey, workbook]);

  // `updateSheet` can wipe presence chrome without changing projectionKey.
  // Re-add the same overlays without a remove cycle so the caret does not blink.
  useEffect(() => {
    if (!workbook || refreshKey === 0) return;
    const projected = JSON.parse(projectionKey) as Presence[];
    if (projected.length) workbook.addPresences(projected);
  }, [projectionKey, refreshKey, workbook]);
}

export interface SpreadsheetFrameCell {
  readonly sheetId: string;
  readonly row: number;
  readonly column: number;
  /** Present only for a plain-text splice. Field writes omit it. */
  readonly indexUtf16?: number;
}

/**
 * Insertion edge of a plain-text splice inside the cell box layout measured.
 * Formula cells, numeric cells, and writes that only name the cell paint nothing.
 */
export function spreadsheetFrameCaretPaint(input: {
  readonly box: TextCaretBox;
  readonly cell: SpreadsheetFrameCell;
  readonly content: WorkSpreadsheetContent;
  readonly layoutSettled: boolean;
  readonly measure: (slice: string) => number;
}): FrameCaretPaint | null {
  if (input.cell.indexUtf16 === undefined) return null;
  const sheet = input.content.sheets.find(
    (candidate) => candidate.id === input.cell.sheetId,
  );
  const source = spreadsheetCellAt(sheet, input.cell.row, input.cell.column);
  const text = spreadsheetSpliceText(source);
  if (text === null) return null;
  return paintTextCaretInBox({
    align: spreadsheetHorizontalAlign(source?.ht),
    box: input.box,
    fontSize:
      typeof source?.fs === 'number' && source.fs > 0
        ? source.fs
        : 11,
    indexUtf16: input.cell.indexUtf16,
    layoutSettled: input.layoutSettled,
    measure: input.measure,
    text,
    verticalAlign: spreadsheetVerticalAlign(source?.vt),
  });
}

function spreadsheetSpliceText(cell: Cell | null): string | null {
  if (typeof cell?.f === 'string' && cell.f.length > 0) return null;
  if (typeof cell?.v === 'number') return null;
  if (typeof cell?.m === 'string') return cell.m;
  if (typeof cell?.v === 'string') return cell.v;
  return '';
}

function spreadsheetHorizontalAlign(
  value: Cell['ht'] | undefined,
): 'left' | 'center' | 'right' {
  if (String(value) === '0') return 'center';
  if (String(value) === '2') return 'right';
  return 'left';
}

function spreadsheetVerticalAlign(
  value: Cell['vt'] | undefined,
): 'top' | 'middle' | 'bottom' {
  if (Number(value) === 0) return 'middle';
  if (Number(value) === 2) return 'bottom';
  return 'top';
}

export function spreadsheetCellLayoutBox(input: {
  readonly column: number;
  readonly columnHeaderHeight: number;
  readonly row: number;
  readonly rowHeaderWidth: number;
  readonly scrollLeft?: number;
  readonly scrollTop?: number;
  readonly sheet: WorkSpreadsheetContent['sheets'][number] | undefined;
}): TextCaretBox | null {
  const sheet = input.sheet;
  if (!sheet || input.row < 0 || input.column < 0) return null;
  const zoom = sheet.zoomRatio && sheet.zoomRatio > 0 ? sheet.zoomRatio : 1;
  const defaultColumn = positiveSize(sheet.defaultColWidth) ?? 96;
  const defaultRow = positiveSize(sheet.defaultRowHeight) ?? 24;
  let left = input.rowHeaderWidth;
  for (let column = 0; column < input.column; column += 1) {
    left += columnSize(sheet, column, defaultColumn);
  }
  let top = input.columnHeaderHeight;
  for (let row = 0; row < input.row; row += 1) {
    top += rowSize(sheet, row, defaultRow);
  }
  const width = columnSize(sheet, input.column, defaultColumn);
  const height = rowSize(sheet, input.row, defaultRow);
  const scrollLeft = input.scrollLeft ?? 0;
  const scrollTop = input.scrollTop ?? 0;
  return {
    left: left * zoom - scrollLeft,
    top: top * zoom - scrollTop,
    width: width * zoom,
    height: height * zoom,
  };
}

function columnSize(
  sheet: WorkSpreadsheetContent['sheets'][number],
  column: number,
  fallback: number,
): number {
  return positiveSize(sheet.config?.columnlen?.[column]) ?? fallback;
}

function rowSize(
  sheet: WorkSpreadsheetContent['sheets'][number],
  row: number,
  fallback: number,
): number {
  return positiveSize(sheet.config?.rowlen?.[row]) ?? fallback;
}

function positiveSize(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) && value > 0
    ? value
    : undefined;
}

/**
 * Highlight for the cell a collaboration frame just wrote. A follow-up
 * location is ignored. The same cell produces the same presence, so an
 * unchanged update does not remove and re-add the overlay.
 */
export function spreadsheetWrittenCellProjection(
  content: WorkSpreadsheetContent,
  cell: SpreadsheetFrameCell | null,
  followUpLocation?: unknown,
): Presence[] {
  void followUpLocation;
  if (!cell) return [];
  const sheet = content.sheets.find(
    (candidate) => candidate.id === cell.sheetId && candidate.hide !== 1,
  );
  if (!sheet) return [];
  const rowCount = Math.max(sheet.row ?? 60, sheet.data?.length ?? 0);
  const columnCount = Math.max(
    sheet.column ?? 26,
    sheet.data?.reduce(
      (maximum, row) => Math.max(maximum, row?.length ?? 0),
      0,
    ) ?? 0,
  );
  if (
    cell.row < 0 ||
    cell.column < 0 ||
    cell.row >= rowCount ||
    cell.column >= columnCount
  ) {
    return [];
  }
  return [
    {
      sheetId: cell.sheetId,
      username: 'Agent',
      userId: 'a3s-office:frame-cell',
      color: '#6d28d9',
      selection: { r: cell.row, c: cell.column },
    },
  ];
}

export function spreadsheetPresenceProjection(
  content: WorkSpreadsheetContent,
  participants: readonly WorkOfficeCollaborationParticipant[],
): Presence[] {
  return participants.flatMap((participant) => {
    const location = participant.location;
    if (location?.kind !== 'spreadsheet') return [];
    const sheet = content.sheets.find(
      (candidate) => candidate.id === location.sheetId && candidate.hide !== 1,
    );
    if (!sheet) return [];
    const selection = location.activeCell ?? {
      row: location.ranges[0]?.startRow ?? -1,
      column: location.ranges[0]?.startColumn ?? -1,
    };
    const rowCount = Math.max(sheet.row ?? 60, sheet.data?.length ?? 0);
    const columnCount = Math.max(
      sheet.column ?? 26,
      sheet.data?.reduce(
        (maximum, row) => Math.max(maximum, row?.length ?? 0),
        0,
      ) ?? 0,
    );
    if (
      selection.row < 0 ||
      selection.column < 0 ||
      selection.row >= rowCount ||
      selection.column >= columnCount
    ) {
      return [];
    }
    return [
      {
        sheetId: location.sheetId,
        username: participant.actor.name,
        userId: `a3s-office:${participant.presenceId}`,
        color: officePresenceColor(participant),
        selection: { r: selection.row, c: selection.column },
      },
    ];
  });
}
