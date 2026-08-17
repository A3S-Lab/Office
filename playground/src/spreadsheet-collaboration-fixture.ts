import {
  createOfficeCollaborationSession,
  type OfficeCollaborationSession,
  readOfficeSpreadsheetCollaboration,
} from '@a3s-lab/office/core';
import { useCallback, useEffect, useState } from 'react';
import * as Y from 'yjs';
import browserSpreadsheetFixtureBase64 from '../../tests/fixtures/browser-spreadsheet-collaboration-update.base64';
import {
  NATIVE_SPREADSHEET_BATCH_CELLS_BASE64,
  NATIVE_SPREADSHEET_CREATE_CELL_BASE64,
  NATIVE_SPREADSHEET_DELETE_CELL_BASE64,
  NATIVE_SPREADSHEET_SET_CELL_BASE64,
} from '../../tests/fixtures/native-spreadsheet-cell-updates';

export type PlaygroundSpreadsheetCellStage =
  | 'ready'
  | 'batched'
  | 'updated'
  | 'created'
  | 'deleted';

export interface PlaygroundSpreadsheetCollaborationFixture {
  readonly collaboration: OfficeCollaborationSession;
  readonly cellStage: PlaygroundSpreadsheetCellStage;
  advanceCell(): void;
}

export function usePlaygroundSpreadsheetCollaborationFixture(
  enabled: boolean,
): PlaygroundSpreadsheetCollaborationFixture | undefined {
  const [ownedFixture, setOwnedFixture] =
    useState<OwnedPlaygroundSpreadsheetCollaborationFixture>();
  const [cellStage, setCellStage] =
    useState<PlaygroundSpreadsheetCellStage>('ready');

  useEffect(() => {
    if (!enabled) return;
    const nextFixture = createSpreadsheetCollaborationFixture();
    setOwnedFixture(nextFixture);
    return () => nextFixture.destroy();
  }, [enabled]);

  const advanceCell = useCallback(() => {
    if (!ownedFixture) return;
    setCellStage(ownedFixture.advanceCell());
  }, [ownedFixture]);

  return ownedFixture
    ? {
        collaboration: ownedFixture.collaboration,
        cellStage,
        advanceCell,
      }
    : undefined;
}

interface OwnedPlaygroundSpreadsheetCollaborationFixture {
  readonly collaboration: OfficeCollaborationSession;
  advanceCell(): PlaygroundSpreadsheetCellStage;
  destroy(): void;
}

function createSpreadsheetCollaborationFixture(): OwnedPlaygroundSpreadsheetCollaborationFixture {
  const document = new Y.Doc();
  Y.applyUpdate(document, decodeBase64(browserSpreadsheetFixtureBase64.trim()));
  const collaboration = createOfficeCollaborationSession({
    actor: { id: 'playground-user', name: 'Lin Cheng', color: '#047857' },
    artifactId: 'fixture-spreadsheet',
    document,
    kind: 'spreadsheet',
  });
  const updates = [
    NATIVE_SPREADSHEET_BATCH_CELLS_BASE64,
    NATIVE_SPREADSHEET_SET_CELL_BASE64,
    NATIVE_SPREADSHEET_CREATE_CELL_BASE64,
    NATIVE_SPREADSHEET_DELETE_CELL_BASE64,
  ];
  const stages: PlaygroundSpreadsheetCellStage[] = [
    'batched',
    'updated',
    'created',
    'deleted',
  ];
  let updateIndex = 0;

  return {
    collaboration,
    advanceCell() {
      const encoded = updates[updateIndex];
      const stage = stages[updateIndex];
      if (!encoded || !stage) return 'deleted';
      Y.applyUpdate(document, decodeBase64(encoded), 'playground-native-agent');
      assertNativeUpdateApplied(collaboration, stage);
      updateIndex += 1;
      return stage;
    },
    destroy() {
      collaboration.destroy();
      document.destroy();
    },
  };
}

function assertNativeUpdateApplied(
  collaboration: OfficeCollaborationSession,
  stage: PlaygroundSpreadsheetCellStage,
): void {
  const content = readOfficeSpreadsheetCollaboration(collaboration);
  if (stage === 'batched') {
    const data = content.sheets.find(({ id }) => id === 'sheet-data')?.data;
    if (
      data?.[0]?.[0] !== null ||
      data?.[1]?.[0]?.bl !== 1 ||
      data?.[3]?.[4]?.v !== 'Batched'
    ) {
      throw new Error(
        'The native Spreadsheet batch-cell update did not project atomically.',
      );
    }
    return;
  }
  if (stage === 'updated') {
    const data = content.sheets.find(({ id }) => id === 'sheet-data')?.data;
    const cell = data?.[1]?.[0];
    if (cell?.v !== 12 || cell.m !== '12' || cell.f !== '=6*2') {
      throw new Error(
        'The native Spreadsheet set-cell update did not project.',
      );
    }
    return;
  }
  if (stage === 'created') {
    const cells = content.sheets.find(
      ({ id }) => id === 'sheet-empty',
    )?.celldata;
    const cell = cells?.find(({ r, c }) => r === 100 && c === 5)?.v;
    if (
      cell?.v !== 'sparse native' ||
      cell.m !== 'sparse native' ||
      cell.ps?.value !== 'Agent note'
    ) {
      throw new Error(
        'The native Spreadsheet sparse cell creation did not project.',
      );
    }
    return;
  }
  if (stage === 'deleted') {
    const cells = content.sheets.find(
      ({ id }) => id === 'sheet-sparse',
    )?.celldata;
    if (cells?.some(({ r, c }) => r === 5 && c === 3)) {
      throw new Error(
        'The native Spreadsheet exact cell deletion did not project.',
      );
    }
  }
}

function decodeBase64(value: string): Uint8Array {
  const decoded = atob(value);
  return Uint8Array.from(decoded, (character) => character.charCodeAt(0));
}
