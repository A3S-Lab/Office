import type { Selection } from '@fortune-sheet/core';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { WorkOfficeCollaborationSession } from '../../../collaboration/office-collaboration';
import {
  createWorkOfficeSpreadsheetCollaborationBinding,
  readWorkOfficeSpreadsheetCollaboration,
  replaceWorkOfficeSpreadsheetCollaboration,
  type WorkOfficeSpreadsheetCollaborationBinding,
} from '../../../collaboration/office-spreadsheet-collaboration';
import type { WorkSpreadsheetContent } from '../work-types';
import { finiteSpreadsheetSelection } from './spreadsheet-editor-support';

export interface SpreadsheetCollaborationHistory {
  canRedo: boolean;
  canUndo: boolean;
  redo: () => boolean;
  undo: () => boolean;
}

export interface SpreadsheetCollaborationViewController {
  activateSheet: (sheetId: string) => boolean;
  select: (sheetId: string, selection: Selection) => void;
  setZoom: (sheetId: string, zoomRatio: number) => void;
}

interface SpreadsheetCollaborationViewState {
  activeSheetId?: string;
  selections: Map<string, Selection>;
  zoomBySheet: Map<string, number>;
}

interface UseSpreadsheetCollaborationOptions {
  initialContent: WorkSpreadsheetContent;
  onChange: (content: WorkSpreadsheetContent) => void;
  session: WorkOfficeCollaborationSession;
}

export function useSpreadsheetCollaboration({
  initialContent,
  onChange,
  session,
}: UseSpreadsheetCollaborationOptions) {
  const sessionRef = useRef(session);
  if (sessionRef.current !== session) {
    throw new Error(
      'SpreadsheetEditor collaboration sessions cannot be replaced while mounted. Remount the editor for another shared workbook.',
    );
  }
  const viewStateRef = useRef<SpreadsheetCollaborationViewState>(
    spreadsheetCollaborationViewState(initialContent),
  );
  const [derivedOrigin] = useState(() =>
    session.createOrigin('system', 'spreadsheet-calculation'),
  );
  const [content, setContent] = useState(() =>
    spreadsheetContentWithCollaborationViewState(
      readWorkOfficeSpreadsheetCollaboration(session),
      viewStateRef.current,
    ),
  );
  const contentRef = useRef(content);
  const bindingRef = useRef<
    WorkOfficeSpreadsheetCollaborationBinding | undefined
  >(undefined);
  const [historyRevision, refreshHistory] = useState(0);
  const onChangeRef = useRef(onChange);
  onChangeRef.current = onChange;
  contentRef.current = content;

  const project = useCallback((sharedContent: WorkSpreadsheetContent) => {
    const projected = spreadsheetContentWithCollaborationViewState(
      sharedContent,
      viewStateRef.current,
    );
    contentRef.current = projected;
    setContent(projected);
    return projected;
  }, []);

  useEffect(() => {
    const binding = createWorkOfficeSpreadsheetCollaborationBinding(session);
    bindingRef.current = binding;
    const unsubscribeContent = binding.subscribe(({ content: next }) => {
      onChangeRef.current(project(next));
    });
    const unsubscribeError = binding.subscribeError((error) => {
      queueMicrotask(() => {
        throw error;
      });
    });
    const unsubscribeHistory = binding.subscribeHistory(() =>
      refreshHistory((value) => value + 1),
    );
    project(binding.content());
    return () => {
      unsubscribeContent();
      unsubscribeError();
      unsubscribeHistory();
      if (bindingRef.current === binding) bindingRef.current = undefined;
      binding.destroy();
    };
  }, [project, session]);

  const commit = useCallback((next: WorkSpreadsheetContent) => {
    const previous = contentRef.current;
    updateSpreadsheetCollaborationViewStateForContentChange(
      viewStateRef.current,
      previous,
      next,
    );
    const binding = bindingRef.current;
    if (binding?.replace(previous, next)) {
      contentRef.current = spreadsheetContentWithCollaborationViewState(
        binding.content(),
        viewStateRef.current,
      );
      return;
    }
    const projected = spreadsheetContentWithCollaborationViewState(
      binding?.content() ?? previous,
      viewStateRef.current,
    );
    contentRef.current = projected;
    setContent(projected);
  }, []);
  const commitDerived = useCallback(
    (next: WorkSpreadsheetContent) => {
      const previous = contentRef.current;
      const binding = bindingRef.current;
      if (
        binding &&
        replaceWorkOfficeSpreadsheetCollaboration(
          session,
          previous,
          next,
          derivedOrigin,
        )
      ) {
        contentRef.current = spreadsheetContentWithCollaborationViewState(
          binding.content(),
          viewStateRef.current,
        );
        return;
      }
      project(binding?.content() ?? previous);
    },
    [derivedOrigin, project, session],
  );
  const view = useMemo<SpreadsheetCollaborationViewController>(
    () => ({
      activateSheet: (sheetId) => {
        const sharedContent =
          bindingRef.current?.content() ?? contentRef.current;
        const sheet = sharedContent.sheets.find(
          (candidate) => candidate.id === sheetId && candidate.hide !== 1,
        );
        if (!sheet) return false;
        if (viewStateRef.current.activeSheetId === sheetId) return true;
        viewStateRef.current.activeSheetId = sheetId;
        project(sharedContent);
        return true;
      },
      select: (sheetId, selection) => {
        viewStateRef.current.selections.set(
          sheetId,
          finiteSpreadsheetSelection(selection),
        );
      },
      setZoom: (sheetId, zoomRatio) => {
        const sharedContent =
          bindingRef.current?.content() ?? contentRef.current;
        if (!sharedContent.sheets.some((sheet) => sheet.id === sheetId)) return;
        if (viewStateRef.current.zoomBySheet.get(sheetId) === zoomRatio) return;
        viewStateRef.current.zoomBySheet.set(sheetId, zoomRatio);
        project(sharedContent);
      },
    }),
    [project],
  );
  const history = useMemo<SpreadsheetCollaborationHistory>(
    () => ({
      canRedo: bindingRef.current?.canRedo() ?? false,
      canUndo: bindingRef.current?.canUndo() ?? false,
      redo: () => bindingRef.current?.redo() ?? false,
      undo: () => bindingRef.current?.undo() ?? false,
    }),
    [historyRevision],
  );

  return {
    content,
    history,
    onChange: commit,
    onDerivedChange: commitDerived,
    readOnly: session.mode !== 'edit',
    view,
  };
}

function spreadsheetCollaborationViewState(
  content: WorkSpreadsheetContent,
): SpreadsheetCollaborationViewState {
  const selections = new Map<string, Selection>();
  const zoomBySheet = new Map<string, number>();
  for (const sheet of content.sheets) {
    if (!sheet.id) continue;
    const selection = sheet.luckysheet_select_save?.at(-1);
    if (selection)
      selections.set(sheet.id, finiteSpreadsheetSelection(selection));
    if (sheet.zoomRatio !== undefined)
      zoomBySheet.set(sheet.id, sheet.zoomRatio);
  }
  return {
    activeSheetId:
      content.sheets.find((sheet) => sheet.status === 1)?.id ??
      content.sheets.find((sheet) => sheet.hide !== 1)?.id,
    selections,
    zoomBySheet,
  };
}

function updateSpreadsheetCollaborationViewStateForContentChange(
  view: SpreadsheetCollaborationViewState,
  previous: WorkSpreadsheetContent,
  next: WorkSpreadsheetContent,
): void {
  pruneSpreadsheetCollaborationViewState(view, next);
  if (!spreadsheetSheetVisibilityChanged(previous, next)) return;
  const active = next.sheets.find(
    (sheet) => sheet.status === 1 && sheet.hide !== 1,
  );
  if (active?.id) view.activeSheetId = active.id;
}

function spreadsheetSheetVisibilityChanged(
  previous: WorkSpreadsheetContent,
  next: WorkSpreadsheetContent,
): boolean {
  if (previous.sheets.length !== next.sheets.length) return true;
  const previousById = new Map(
    previous.sheets.flatMap((sheet) =>
      sheet.id ? [[sheet.id, sheet.hide] as const] : [],
    ),
  );
  return next.sheets.some(
    (sheet) =>
      !sheet.id ||
      !previousById.has(sheet.id) ||
      previousById.get(sheet.id) !== sheet.hide,
  );
}

function spreadsheetContentWithCollaborationViewState(
  content: WorkSpreadsheetContent,
  view: SpreadsheetCollaborationViewState,
): WorkSpreadsheetContent {
  pruneSpreadsheetCollaborationViewState(view, content);
  const visibleIds = new Set(
    content.sheets
      .filter((sheet) => sheet.hide !== 1 && sheet.id)
      .map((sheet) => sheet.id as string),
  );
  const activeSheetId =
    (view.activeSheetId && visibleIds.has(view.activeSheetId)
      ? view.activeSheetId
      : content.sheets.find((sheet) => sheet.hide !== 1)?.id) ??
    content.sheets[0]?.id;
  if (activeSheetId) view.activeSheetId = activeSheetId;
  return {
    ...content,
    sheets: content.sheets.map((sheet) => {
      const id = sheet.id;
      const selection = id ? view.selections.get(id) : undefined;
      const zoomRatio = id ? view.zoomBySheet.get(id) : undefined;
      return {
        ...sheet,
        status: id === activeSheetId ? 1 : 0,
        ...(selection ? { luckysheet_select_save: [selection] } : {}),
        ...(zoomRatio === undefined ? {} : { zoomRatio }),
      };
    }),
  };
}

function pruneSpreadsheetCollaborationViewState(
  view: SpreadsheetCollaborationViewState,
  content: WorkSpreadsheetContent,
): void {
  const sheetIds = new Set(
    content.sheets.flatMap((sheet) => (sheet.id ? [sheet.id] : [])),
  );
  for (const sheetId of view.selections.keys()) {
    if (!sheetIds.has(sheetId)) view.selections.delete(sheetId);
  }
  for (const sheetId of view.zoomBySheet.keys()) {
    if (!sheetIds.has(sheetId)) view.zoomBySheet.delete(sheetId);
  }
}
