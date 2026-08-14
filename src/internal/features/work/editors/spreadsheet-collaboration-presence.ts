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
  workbook,
}: {
  content: WorkSpreadsheetContent;
  workbook: WorkbookInstance | null;
}): void {
  const participants = useOfficeRemoteParticipants();

  useEffect(() => {
    if (!workbook) return;
    const projected = spreadsheetPresenceProjection(content, participants);
    if (projected.length) workbook.addPresences(projected);
    return () => {
      if (!projected.length) return;
      workbook.removePresences(
        projected.map(({ userId, username }) => ({ userId, username })),
      );
    };
  }, [content, participants, workbook]);
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
