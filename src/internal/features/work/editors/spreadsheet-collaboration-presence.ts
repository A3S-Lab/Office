import type { Presence } from '@fortune-sheet/core';
import type { WorkbookInstance } from '@fortune-sheet/react';
import { useEffect } from 'react';
import type { WorkOfficeCollaborationParticipant } from '../../../collaboration/office-collaboration-presence';
import type { WorkSpreadsheetContent } from '../work-types';
import {
  officePresenceColor,
  useOfficeRemoteParticipants,
} from './office-collaboration-presence-ui';

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
