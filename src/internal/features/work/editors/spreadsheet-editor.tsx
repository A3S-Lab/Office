import type { Hooks, Op, Selection } from '@fortune-sheet/core';
import { Workbook, type WorkbookInstance } from '@fortune-sheet/react';
import { Cloud, Grid3X3 } from 'lucide-react';
import {
  memo,
  type ClipboardEvent as ReactClipboardEvent,
  useCallback,
  useEffect,
  useId,
  useMemo,
  useRef,
  useState,
} from 'react';
import type { WorkOfficeCollaborationSession } from '../../../collaboration/office-collaboration';
import type {
  WorkOfficeCollaborationParticipant,
  WorkOfficeSpreadsheetPresenceCell,
  WorkOfficeSpreadsheetPresenceRange,
} from '../../../collaboration/office-collaboration-presence';
import { showToast } from '../../../state/app-state';
import {
  isWorkspaceContextMenuKeyboardEvent,
  WorkspaceContextMenu,
  type WorkspaceContextMenuEvent,
  workspaceContextMenuPosition,
} from '../../workspace/components/workspace-context-menu';
import { spreadsheetAgentMenuItems } from '../components/work-editor-agent-menus';
import { spreadsheetGridSize } from '../spreadsheet-sparse';
import { applySpreadsheetAgentProposalChanges } from '../work-agent-proposal-apply';
import type { WorkEditorAgentRequest } from '../work-agent-request';
import {
  spreadsheetAgentSelection,
  type WorkSpreadsheetAgentSelection,
} from '../work-spreadsheet-agent-context';
import {
  reconcileSpreadsheetChartPreviews,
  spreadsheetSheetsWithChartPreviews,
} from '../work-spreadsheet-charts';
import {
  drawSpreadsheetCommentMarker,
  drawSpreadsheetConditionalDataBar,
} from '../work-spreadsheet-conditional-canvas';
import { spreadsheetConditionalFormatStyles } from '../work-spreadsheet-conditional-format';
import { drawSpreadsheetConditionalIcon } from '../work-spreadsheet-conditional-icons';
import { spreadsheetMatrixProfile } from '../work-spreadsheet-matrix-profile';
import {
  reconcileSpreadsheetPivots,
  refreshSpreadsheetPivotTables,
  spreadsheetPivotIntersects,
  spreadsheetPivotOutputContains,
} from '../work-spreadsheet-pivots';
import { spreadsheetProtectionKey } from '../work-spreadsheet-protection';
import type { WorkSpreadsheetContent } from '../work-types';
import { useOfficeCollaborationLocationNavigator } from './office-collaboration-presence-context';
import { useOfficePublishPresenceLocation } from './office-collaboration-presence-ui';
import { useOfficeDialog } from './office-dialog';
import { useOfficeEditorFocusOrigin } from './office-editor-focus-handoff';
import { spreadsheetCellBordersAt } from './spreadsheet-cell-border';
import { useSpreadsheetCollaborationPresenceProjection } from './spreadsheet-collaboration-presence';
import {
  createSpreadsheetEditorExtensions,
  type SpreadsheetCommandRange,
  type SpreadsheetEditorCommands,
  type SpreadsheetFormatCellsOpenRequest,
  type SpreadsheetStructureAxis,
} from './spreadsheet-command-controller';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import {
  browserSpreadsheetClipboard,
  copySpreadsheetSelection,
  parseSpreadsheetClipboardText,
  pasteSpreadsheetSelection,
  spreadsheetCoreContextMenuItems,
  spreadsheetSortContextMenuItems,
  spreadsheetStructureContextMenuItems,
} from './spreadsheet-context-menu';
import {
  SpreadsheetEditorRibbon,
  type SpreadsheetRibbonTabId,
} from './spreadsheet-editor-ribbon';
import {
  finiteSpreadsheetSelection,
  isSpreadsheetCellEditingTarget,
  isSpreadsheetNativeTextUndoTarget,
  sameSpreadsheetHistoryContent,
  sameSpreadsheetWorkbookState,
  sameSpreadsheetWorkbookStateAfterOperations,
  spreadsheetCellAt,
  spreadsheetContentWithSelection,
  spreadsheetSelectionReference,
  spreadsheetSelectionSummary,
  spreadsheetSheetsForFortune,
  spreadsheetSheetsFromFortune,
  spreadsheetSheetsWithFiniteSelections,
  spreadsheetSingleRange,
} from './spreadsheet-editor-support';
import type { SpreadsheetFindMatch } from './spreadsheet-find';
import { SpreadsheetFindBar } from './spreadsheet-find-bar';
import { SpreadsheetFormatCellsDialog } from './spreadsheet-format-cells-dialog';
import {
  createSpreadsheetFormatCellsDialogSource,
  type SpreadsheetFormatCellsDialogSource,
} from './spreadsheet-format-cells-dialog-model';
import { spreadsheetFreezePanesStatus } from './spreadsheet-freeze-panes';
import {
  resolveSpreadsheetGoToTarget,
  spreadsheetGoToValidationMessage,
} from './spreadsheet-go-to';
import { synchronizeSpreadsheetWorkbookInPlace } from './spreadsheet-in-place-workbook-sync';
import { spreadsheetSelectionContainsFocus } from './spreadsheet-keyboard-navigation';
import { SpreadsheetSheetBar } from './spreadsheet-sheet-bar';
import {
  SpreadsheetWorkbookPanel,
  type SpreadsheetWorkbookPanelView,
} from './spreadsheet-workbook-panel';
import { useOfficeEditorKeyboardShortcuts } from './use-office-editor-keyboard-shortcuts';
import { useOfficeEditorRuntime } from './use-office-editor-runtime';
import {
  stepOfficeZoom,
  useOfficeEditorWheelZoom,
} from './use-office-editor-wheel-zoom';
import { useOfficeHistory } from './use-office-history';
import { useSpreadsheetAutoFilter } from './use-spreadsheet-auto-filter';
import { useSpreadsheetCalculation } from './use-spreadsheet-calculation';
import {
  type SpreadsheetCollaborationHistory,
  type SpreadsheetCollaborationViewController,
  useSpreadsheetCollaboration,
} from './use-spreadsheet-collaboration';
import {
  type SpreadsheetFormatPainterMode,
  useSpreadsheetFormatPainter,
} from './use-spreadsheet-format-painter';
import { useSpreadsheetWorkbookSync } from './use-spreadsheet-workbook-sync';
import {
  type WorkOfficeFileAction,
  WorkOfficePreviewBar,
  WorkOfficeStatusBar,
  WorkOfficeZoomControls,
} from './work-office-chrome';

export interface SpreadsheetEditorProps {
  autoFocus?: boolean;
  collaboration?: WorkOfficeCollaborationSession;
  content: WorkSpreadsheetContent;
  kernelWasmUrl?: string;
  preview: boolean;
  saveStatus?: string;
  fileActions?: readonly WorkOfficeFileAction[];
  onChange: (content: WorkSpreadsheetContent) => void;
  onAgentRequest?: (request: WorkEditorAgentRequest) => void | Promise<void>;
}

const spreadsheetFocusRetryFrames = 12;
const spreadsheetFocusObservationMs = 5_000;
const spreadsheetFocusCleanups = new WeakMap<HTMLElement, () => void>();
const ControlledWorkbook = memo(Workbook, (previous, next) => {
  const previousProps = previous as Record<string, unknown>;
  const nextProps = next as Record<string, unknown>;
  const previousKeys = Object.keys(previousProps).filter(
    (key) => key !== 'data',
  );
  const nextKeys = Object.keys(nextProps).filter((key) => key !== 'data');
  return (
    previousKeys.length === nextKeys.length &&
    previousKeys.every(
      (key) =>
        Object.hasOwn(nextProps, key) && previousProps[key] === nextProps[key],
    )
  );
});

interface SpreadsheetSelectionState {
  sheetId: string;
  selection: Selection;
}

interface SpreadsheetContextMenuState {
  kind: 'cells' | SpreadsheetStructureAxis;
  range: SpreadsheetCommandRange;
  x: number;
  y: number;
  selection: WorkSpreadsheetAgentSelection;
}

export function SpreadsheetEditor(props: SpreadsheetEditorProps) {
  if (props.collaboration) {
    return (
      <CollaborativeSpreadsheetEditor
        {...props}
        collaboration={props.collaboration}
      />
    );
  }
  return <SpreadsheetEditorSurface {...props} />;
}

function CollaborativeSpreadsheetEditor(
  props: SpreadsheetEditorProps & {
    collaboration: WorkOfficeCollaborationSession;
  },
) {
  const { collaboration } = props;
  const shared = useSpreadsheetCollaboration({
    initialContent: props.content,
    onChange: props.onChange,
    session: collaboration,
  });
  return (
    <SpreadsheetEditorSurface
      {...props}
      content={shared.content}
      collaborationHistory={shared.history}
      collaborationView={shared.view}
      onChange={shared.onChange}
      onDerivedChange={shared.onDerivedChange}
      preview={props.preview || shared.readOnly}
    />
  );
}

interface SpreadsheetEditorSurfaceProps extends SpreadsheetEditorProps {
  collaborationHistory?: SpreadsheetCollaborationHistory;
  collaborationView?: SpreadsheetCollaborationViewController;
  onDerivedChange?: (content: WorkSpreadsheetContent) => void;
}

function SpreadsheetEditorSurface({
  autoFocus = true,
  content,
  collaborationHistory,
  collaborationView,
  kernelWasmUrl,
  preview,
  saveStatus = '已自动保存',
  fileActions,
  onChange,
  onDerivedChange,
  onAgentRequest,
}: SpreadsheetEditorSurfaceProps) {
  const materializedContent = useMemo(
    () => refreshSpreadsheetPivotTables(content),
    [content],
  );
  const contentRef = useRef(materializedContent);
  const spreadsheetCommandsRef = useRef<SpreadsheetEditorCommands | null>(null);
  const previewRef = useRef(preview);
  const spreadsheetRootRef = useRef<HTMLElement>(null);
  const spreadsheetCanvasRef = useRef<HTMLDivElement>(null);
  const editorFocusOrigin = useOfficeEditorFocusOrigin();
  const workbookRef = useRef<WorkbookInstance>(null);
  const projectedWorkbookSheetsRef = useRef<WorkSpreadsheetContent['sheets']>(
    [],
  );
  const spreadsheetZoomRef = useRef(100);
  const [workbookInstance, setWorkbookInstance] =
    useState<WorkbookInstance | null>(null);
  const bindWorkbookInstance = useCallback(
    (instance: WorkbookInstance | null) => {
      workbookRef.current = instance;
      setWorkbookInstance((current) =>
        Object.is(current, instance) ? current : instance,
      );
    },
    [],
  );
  const synchronizeExternalWorkbook = useCallback(
    (previous: WorkSpreadsheetContent, next: WorkSpreadsheetContent) =>
      synchronizeSpreadsheetWorkbookInPlace(
        workbookRef.current,
        previous,
        next,
        projectedWorkbookSheetsRef.current,
        previewRef.current,
      ),
    [],
  );
  const {
    acceptContent: acceptWorkbookContent,
    ignoreChangeDuringExternalSync,
    mountRevision: workbookMountRevision,
    recordOperations: recordWorkbookOperations,
    takeOperations: takeWorkbookOperations,
  } = useSpreadsheetWorkbookSync(
    materializedContent,
    synchronizeExternalWorkbook,
  );
  const calculation = useSpreadsheetCalculation({
    content: materializedContent,
    kernelWasmUrl,
    workbookRef,
  });
  const panelId = useId();
  const [ribbonTab, setRibbonTab] = useState<SpreadsheetRibbonTabId>('home');
  const [panel, setPanel] = useState<SpreadsheetWorkbookPanelView | null>(null);
  const panelTriggerRef = useRef<HTMLButtonElement | null>(null);
  const [selectionState, setSelectionState] =
    useState<SpreadsheetSelectionState | null>(null);
  const formatPainterSelectionHandlerRef = useRef(
    (_sheetId: string, _selection: Selection) => undefined,
  );
  const [contextMenu, setContextMenu] =
    useState<SpreadsheetContextMenuState | null>(null);
  const [formatCellsSource, setFormatCellsSource] =
    useState<SpreadsheetFormatCellsDialogSource | null>(null);
  const formatCellsInvokerRef = useRef<HTMLElement | null>(null);
  const formatCellsSelectionRef = useRef<SpreadsheetSelectionState | null>(
    null,
  );
  const formatCellsApplyingRef = useRef(false);
  const officeDialog = useOfficeDialog();
  const [findOpen, setFindOpen] = useState(false);
  const [findFocusRequest, setFindFocusRequest] = useState(0);
  const [navigationActiveSheetId, setNavigationActiveSheetId] = useState<
    string | null
  >(null);
  const [previewZoom, setPreviewZoom] = useState(100);
  const [previewActiveSheetId, setPreviewActiveSheetId] = useState<
    string | null
  >(null);
  const contentActiveSheetId =
    content.sheets.find((sheet) => sheet.status === 1)?.id ??
    content.sheets.find((sheet) => !sheet.hide)?.id ??
    '';
  const viewActiveSheetId = preview
    ? previewActiveSheetId
    : navigationActiveSheetId;
  const activeSheetId =
    viewActiveSheetId &&
    content.sheets.some(
      (sheet) => sheet.id === viewActiveSheetId && sheet.hide !== 1,
    )
      ? viewActiveSheetId
      : contentActiveSheetId;
  const activeSheetIdRef = useRef(activeSheetId);
  const focusedSheetIdRef = useRef<string | null>(null);
  contentRef.current = materializedContent;
  previewRef.current = preview;
  const controlledHistory = useOfficeHistory({
    content,
    onChange: (next) => {
      const liveSelection = workbookRef.current?.getSelection()?.at(-1);
      onChange(
        spreadsheetContentWithSelection(
          next,
          selectionState?.sheetId ?? activeSheetIdRef.current,
          liveSelection ?? selectionState?.selection,
        ),
      );
    },
    sameValue: collaborationHistory
      ? ignoreSpreadsheetControlledHistory
      : sameSpreadsheetHistoryContent,
  });
  const history = collaborationHistory ?? controlledHistory;
  const conditionalStylesBySheet = useMemo(
    () =>
      new Map(
        materializedContent.sheets.flatMap((sheet) =>
          sheet.id
            ? [[sheet.id, spreadsheetConditionalFormatStyles(sheet)] as const]
            : [],
        ),
      ),
    [materializedContent.sheets],
  );
  const conditionalStylesBySheetRef = useRef(conditionalStylesBySheet);
  conditionalStylesBySheetRef.current = conditionalStylesBySheet;
  useEffect(() => {
    setPreviewActiveSheetId(preview ? contentActiveSheetId : null);
  }, [contentActiveSheetId, preview]);
  useEffect(() => {
    setNavigationActiveSheetId(null);
  }, [contentActiveSheetId]);
  useEffect(() => {
    activeSheetIdRef.current = activeSheetId;
  }, [activeSheetId]);
  useEffect(() => {
    if (
      preview ||
      !activeSheetId ||
      focusedSheetIdRef.current === activeSheetId
    )
      return;
    const initialFocus = focusedSheetIdRef.current === null;
    if (initialFocus && !autoFocus) {
      focusedSheetIdRef.current = activeSheetId;
      return;
    }
    focusedSheetIdRef.current = activeSheetId;
    const frame = requestAnimationFrame(() =>
      focusSpreadsheetGrid(spreadsheetCanvasRef.current, {
        focusOrigin: initialFocus ? editorFocusOrigin : document.activeElement,
        forceInitial: !initialFocus,
      }),
    );
    return () => cancelAnimationFrame(frame);
  }, [activeSheetId, autoFocus, editorFocusOrigin, preview]);
  useEffect(() => {
    const container = spreadsheetCanvasRef.current;
    if (!container || preview) return;
    const enhanceTriggers = () => {
      for (const trigger of container.querySelectorAll<HTMLElement>(
        '.header-arrow',
      )) {
        trigger.setAttribute('role', 'button');
        trigger.setAttribute('aria-label', '列操作');
        trigger.setAttribute('aria-haspopup', 'menu');
        trigger.setAttribute('title', '列操作');
      }
    };
    enhanceTriggers();
    if (typeof MutationObserver === 'undefined') return;
    const observer = new MutationObserver(enhanceTriggers);
    observer.observe(container, { childList: true, subtree: true });
    return () => observer.disconnect();
  }, [preview, workbookMountRevision]);
  useEffect(() => {
    if (!preview) return;
    panelTriggerRef.current = null;
    setPanel(null);
    setContextMenu(null);
    setFormatCellsSource(null);
    formatCellsInvokerRef.current = null;
    formatCellsSelectionRef.current = null;
    formatCellsApplyingRef.current = false;
  }, [preview]);
  const closeWorkbookPanel = useCallback(() => {
    const trigger = panelTriggerRef.current;
    setPanel(null);
    requestAnimationFrame(() => {
      if (trigger?.isConnected) trigger.focus({ preventScroll: true });
      if (panelTriggerRef.current === trigger) panelTriggerRef.current = null;
    });
  }, []);
  const workbookHooks = useMemo<Hooks>(
    () => ({
      afterActivateSheet: (id) => {
        activeSheetIdRef.current = id;
        collaborationView?.activateSheet(id);
        if (previewRef.current) setPreviewActiveSheetId(id);
        else setNavigationActiveSheetId(id);
        setSelectionState(null);
      },
      afterSelectionChange: (sheetId, selection) => {
        collaborationView?.select(sheetId, selection);
        setSelectionState({ sheetId, selection });
        formatPainterSelectionHandlerRef.current(sheetId, selection);
      },
      beforeUpdateCell: (row, column) => {
        const sheet = contentRef.current.sheets.find(
          (candidate) => candidate.id === activeSheetIdRef.current,
        );
        return !sheet || !spreadsheetPivotOutputContains(sheet, row, column);
      },
      beforePaste: (selections) => {
        const sheet = contentRef.current.sheets.find(
          (candidate) => candidate.id === activeSheetIdRef.current,
        );
        if (!sheet) return true;
        return !(selections ?? []).some(
          (selection) =>
            spreadsheetPivotIntersects(sheet, {
              startRow: Math.min(selection.row[0], selection.row[1]),
              endRow: Math.max(selection.row[0], selection.row[1]),
              startColumn: Math.min(selection.column[0], selection.column[1]),
              endColumn: Math.max(selection.column[0], selection.column[1]),
            }).length,
        );
      },
      beforeRenderCell: (_cell, cellInfo, context) => {
        const style = conditionalStylesBySheetRef.current
          .get(activeSheetIdRef.current)
          ?.get(`${cellInfo.row}_${cellInfo.column}`);
        if (style?.cellColor) context.fillStyle = style.cellColor;
        return true;
      },
      afterRenderCell: (cell, cellInfo, context) => {
        const style = conditionalStylesBySheetRef.current
          .get(activeSheetIdRef.current)
          ?.get(`${cellInfo.row}_${cellInfo.column}`);
        if (!style?.icon && !style?.dataBar) return;
        const background =
          style.cellColor ??
          (typeof cell?.bg === 'string' ? cell.bg : '#ffffff');
        if (style.dataBar) {
          drawSpreadsheetConditionalDataBar(
            context,
            cellInfo,
            {
              ...style.dataBar,
              showValue:
                style.dataBar.showValue && (style.icon?.showValue ?? true),
            },
            cell,
            background,
            style.textColor,
          );
        }
        if (style.icon) {
          drawSpreadsheetConditionalIcon(
            context,
            cellInfo,
            style.icon,
            background,
            style.dataBar ? false : !style.icon.showValue,
          );
        }
        if (cell?.ps) drawSpreadsheetCommentMarker(context, cellInfo);
      },
    }),
    [collaborationView],
  );
  const conditionalFormatKey = useMemo(
    () =>
      content.sheets
        .map(
          (sheet) =>
            `${sheet.id}:${JSON.stringify(sheet.luckysheet_conditionformat_save ?? [])}`,
        )
        .join('|'),
    [content.sheets],
  );
  const protectionKey = useMemo(
    () => spreadsheetProtectionKey(content.sheets),
    [content.sheets],
  );
  const renderedWorkbookSheets = useMemo(
    () =>
      spreadsheetSheetsWithFiniteSelections(
        spreadsheetSheetsWithChartPreviews(materializedContent),
      ),
    [materializedContent],
  );
  const workbookSheets = useMemo(
    () => spreadsheetSheetsForFortune(renderedWorkbookSheets),
    [renderedWorkbookSheets],
  );
  const displayedWorkbookSheets = useMemo(
    () =>
      preview
        ? workbookSheets.map((sheet) => ({
            ...sheet,
            status: sheet.id === activeSheetId ? 1 : 0,
            zoomRatio: previewZoom / 100,
          }))
        : workbookSheets,
    [activeSheetId, preview, previewZoom, workbookSheets],
  );
  projectedWorkbookSheetsRef.current = displayedWorkbookSheets;
  const handleWorkbookChange = useCallback(
    (sheets: WorkSpreadsheetContent['sheets']) => {
      const operations = takeWorkbookOperations();
      const matchesRenderedWorkbook =
        sameSpreadsheetWorkbookStateAfterOperations(
          sheets,
          projectedWorkbookSheetsRef.current,
          operations,
        ) ??
        sameSpreadsheetWorkbookState(
          sheets,
          projectedWorkbookSheetsRef.current,
        );
      if (ignoreChangeDuringExternalSync(matchesRenderedWorkbook)) {
        return;
      }
      const hasPendingResultPatches = calculation.hasPendingResultPatches();
      if (
        previewRef.current ||
        (matchesRenderedWorkbook && !hasPendingResultPatches)
      )
        return;
      const controlledSheets = spreadsheetSheetsFromFortune(
        sheets,
        contentRef.current.sheets,
        operations,
      );
      const withCharts = reconcileSpreadsheetChartPreviews(
        contentRef.current,
        controlledSheets,
      );
      const next = reconcileSpreadsheetPivots(
        contentRef.current,
        withCharts.sheets,
      );
      contentRef.current = next;
      acceptWorkbookContent(next);
      calculation.synchronizeWorkbook(next, operations);
      if (matchesRenderedWorkbook && onDerivedChange) onDerivedChange(next);
      else spreadsheetCommandsRef.current?.setSpreadsheetContent(next);
    },
    [
      acceptWorkbookContent,
      calculation,
      ignoreChangeDuringExternalSync,
      onDerivedChange,
      takeWorkbookOperations,
    ],
  );
  const handleWorkbookOperations = useCallback(
    (operations: Op[]) => {
      recordWorkbookOperations(operations);
      calculation.notifyWorkbookOperations(operations);
    },
    [calculation, recordWorkbookOperations],
  );
  const chartPreviewKey = useMemo(
    () =>
      workbookSheets
        .flatMap((sheet) =>
          (sheet.images ?? [])
            .filter((image) => image.id.startsWith('work-chart-preview-'))
            .map(
              (image) =>
                `${image.id}:${image.src}:${image.left}:${image.top}:${image.width}:${image.height}`,
            ),
        )
        .join('|'),
    [workbookSheets],
  );
  const panelSheetId =
    selectionState?.sheetId ?? activeSheetIdRef.current ?? activeSheetId;
  const activeSheet = materializedContent.sheets.find(
    (sheet) => sheet.id === activeSheetId,
  );
  const activeSheetProfile = spreadsheetMatrixProfile(activeSheet?.data);
  const zoom = Math.round((activeSheet?.zoomRatio ?? 1) * 100);
  spreadsheetZoomRef.current = preview ? previewZoom : zoom;
  useEffect(() => {
    if (preview) setPreviewZoom(zoom);
  }, [preview, zoom]);
  const gridLinesVisible =
    activeSheet?.showGridLines !== false && activeSheet?.showGridLines !== 0;
  const toolbarSheetId = selectionState?.sheetId ?? activeSheetId;
  const targetSheetGridSize = useMemo(() => {
    const sheet = materializedContent.sheets.find(
      (candidate) => candidate.id === toolbarSheetId,
    );
    const profile = spreadsheetMatrixProfile(sheet?.data);
    return profile
      ? {
          columnCount: Math.max(sheet?.column ?? 0, profile.columnCount),
          rowCount: Math.max(sheet?.row ?? 0, profile.rowCount),
        }
      : spreadsheetGridSize(sheet);
  }, [materializedContent.sheets, toolbarSheetId]);
  const toolbarSheet = workbookSheets.find(
    (sheet) => sheet.id === toolbarSheetId,
  );
  const toolbarSelection = finiteSpreadsheetSelection(
    selectionState?.selection ?? toolbarSheet?.luckysheet_select_save?.at(-1),
  );
  const spreadsheetPresenceLocation = useMemo(() => {
    if (!toolbarSheetId) return null;
    const liveSelections = workbookRef.current?.getSelection();
    const ranges = (
      liveSelections?.length ? liveSelections : [toolbarSelection]
    )
      .slice(0, 64)
      .map((selection) => ({
        startRow: Math.min(selection.row[0] ?? 0, selection.row[1] ?? 0),
        startColumn: Math.min(
          selection.column[0] ?? 0,
          selection.column[1] ?? 0,
        ),
        endRow: Math.max(selection.row[0] ?? 0, selection.row[1] ?? 0),
        endColumn: Math.max(selection.column[0] ?? 0, selection.column[1] ?? 0),
      }));
    const activeCell = {
      row: toolbarSelection.row_focus ?? toolbarSelection.row[0] ?? 0,
      column: toolbarSelection.column_focus ?? toolbarSelection.column[0] ?? 0,
    };
    return {
      kind: 'spreadsheet' as const,
      sheetId: toolbarSheetId,
      ranges,
      activeCell,
    };
  }, [selectionState, toolbarSelection, toolbarSheetId]);
  useOfficePublishPresenceLocation(spreadsheetPresenceLocation);
  useSpreadsheetCollaborationPresenceProjection({
    content: materializedContent,
    workbook: workbookInstance,
  });
  const toolbarRow = toolbarSelection.row_focus ?? toolbarSelection.row[0];
  const toolbarColumn =
    toolbarSelection.column_focus ?? toolbarSelection.column[0];
  const toolbarCell = spreadsheetCellAt(
    toolbarSheet,
    toolbarRow,
    toolbarColumn,
  );
  const toolbarCellBorders = spreadsheetCellBordersAt(
    toolbarSheet,
    toolbarRow,
    toolbarColumn,
  );
  const selectedRange = spreadsheetSingleRange(toolbarSelection);
  const multipleCellsSelected =
    selectedRange.row[0] !== selectedRange.row[1] ||
    selectedRange.column[0] !== selectedRange.column[1];
  const selectionSummary = multipleCellsSelected
    ? spreadsheetSelectionSummary(toolbarSheet, toolbarSelection)
    : null;
  const { commandPort: formatPainter, mode: formatPainterMode } =
    useSpreadsheetFormatPainter({
      content: materializedContent,
      editable: !preview,
      onError: (message) => showToast(message, 'error'),
      sourceRange: selectedRange,
      sourceSheetId: toolbarSheetId,
      workbook: workbookInstance,
    });
  const {
    active: autoFilterActive,
    commandPort: autoFilter,
    reserveAltKey: reserveAutoFilterAltKey,
    status: autoFilterStatus,
  } = useSpreadsheetAutoFilter({
    canvasRef: spreadsheetCanvasRef,
    content: materializedContent,
    editable: !preview,
    mountRevision: workbookMountRevision,
    onChange: (next) => {
      contentRef.current = next;
      onChange(next);
    },
    selection: toolbarSelection,
    sheetId: toolbarSheetId,
    workbook: workbookInstance,
  });
  const currentClipboardSelection = () => {
    if (previewRef.current) return null;
    const sheetId = selectionState?.sheetId ?? activeSheetIdRef.current;
    const sheet = contentRef.current.sheets.find(
      (candidate) => candidate.id === sheetId,
    );
    const selection =
      selectionState?.selection ?? sheet?.luckysheet_select_save?.at(-1);
    if (!selection) return null;
    const range = spreadsheetSingleRange(selection);
    const maximumCells =
      (range.row[1] - range.row[0] + 1) *
      (range.column[1] - range.column[0] + 1);
    return spreadsheetAgentSelection(
      contentRef.current,
      sheetId,
      selection,
      maximumCells,
    );
  };
  const runSpreadsheetClipboardCopy = (cut: boolean): boolean => {
    const selection = currentClipboardSelection();
    const commands = spreadsheetCommandsRef.current;
    if (!selection || (cut && !commands)) return false;
    void copySpreadsheetSelection(
      browserSpreadsheetClipboard,
      selection.clipboard,
      cut,
    ).then((copied) => {
      if (copied && cut && !commands?.clearSelectedCells()) {
        showToast('选区已复制，但无法清除原内容。', 'error');
      }
    });
    return true;
  };
  const runSpreadsheetClipboardPaste = (): boolean => {
    const commands = spreadsheetCommandsRef.current;
    if (!commands) return false;
    void pasteSpreadsheetSelection(
      browserSpreadsheetClipboard,
      commands.pasteCells,
    );
    return true;
  };
  const openSpreadsheetFormatCells = useCallback(
    (request: SpreadsheetFormatCellsOpenRequest): boolean => {
      if (previewRef.current) return false;
      const source = createSpreadsheetFormatCellsDialogSource(
        contentRef.current,
        request.sheetId,
        request.range,
        request.cells,
        request.activeCell,
      );
      if (!source) return false;
      formatCellsInvokerRef.current =
        document.activeElement instanceof HTMLElement
          ? document.activeElement
          : spreadsheetGridFocusTarget(spreadsheetCanvasRef.current);
      formatCellsSelectionRef.current = {
        sheetId: request.sheetId,
        selection: {
          row: [...request.range.row],
          column: [...request.range.column],
          row_focus: request.activeCell.row,
          column_focus: request.activeCell.column,
        },
      };
      setFormatCellsSource(source);
      return true;
    },
    [],
  );
  const closeSpreadsheetFormatCells = useCallback(() => {
    const invoker = formatCellsInvokerRef.current;
    const canvas = spreadsheetCanvasRef.current;
    const openedFromGrid = Boolean(invoker && canvas?.contains(invoker));
    setFormatCellsSource(null);
    requestAnimationFrame(() => {
      if (openedFromGrid || !invoker?.isConnected) {
        focusSpreadsheetGrid(canvas, { focusOrigin: invoker });
      } else {
        invoker.focus({ preventScroll: true });
      }
      if (formatCellsInvokerRef.current === invoker) {
        formatCellsInvokerRef.current = null;
      }
      formatCellsSelectionRef.current = null;
    });
  }, []);
  const activateLocalSpreadsheetSheet = useCallback(
    (sheetId: string) => {
      if (!previewRef.current && !collaborationView) return false;
      const sheet = contentRef.current.sheets.find(
        (candidate) => candidate.id === sheetId && candidate.hide !== 1,
      );
      if (!sheet) return false;
      activeSheetIdRef.current = sheetId;
      if (collaborationView && !collaborationView.activateSheet(sheetId)) {
        return false;
      }
      setSelectionState(null);
      if (previewRef.current) setPreviewActiveSheetId(sheetId);
      try {
        workbookRef.current?.activateSheet({ id: sheetId });
      } catch {
        return false;
      }
      return true;
    },
    [collaborationView],
  );
  const navigateToSpreadsheetRange = useCallback(
    (
      sheetId: string,
      ranges: SpreadsheetCommandRange[],
      activeCell: { column: number; row: number },
      focusGrid = true,
    ): boolean => {
      const workbook = workbookRef.current;
      const sheet = contentRef.current.sheets.find(
        (candidate) => candidate.id === sheetId && candidate.hide !== 1,
      );
      if (!workbook || !sheet || !ranges.length) return false;
      const focusedRange =
        ranges.find((range) =>
          spreadsheetSelectionContainsFocus(range, activeCell),
        ) ?? ranges[0];
      if (!focusedRange) return false;
      try {
        if (collaborationView && !collaborationView.activateSheet(sheetId)) {
          return false;
        }
        activeSheetIdRef.current = sheetId;
        if (previewRef.current) setPreviewActiveSheetId(sheetId);
        else setNavigationActiveSheetId(sheetId);
        workbook.activateSheet({ id: sheetId });
        workbook.setSelection(ranges, { id: sheetId });
        workbook.scroll({
          targetRow: activeCell.row,
          targetColumn: activeCell.column,
        });
        setSelectionState({
          sheetId,
          selection: {
            row: [...focusedRange.row],
            column: [...focusedRange.column],
            row_focus: activeCell.row,
            column_focus: activeCell.column,
          },
        });
      } catch {
        return false;
      }
      if (focusGrid) {
        requestAnimationFrame(() =>
          focusSpreadsheetGrid(spreadsheetCanvasRef.current),
        );
      }
      return true;
    },
    [collaborationView],
  );
  const openSpreadsheetFind = useCallback((): boolean => {
    if (previewRef.current) return false;
    setFindOpen(true);
    setFindFocusRequest((current) => current + 1);
    return true;
  }, []);
  const openSpreadsheetGoTo = useCallback((): boolean => {
    if (previewRef.current) return false;
    const sourceSheetId = activeSheetIdRef.current;
    const sourceSheet = contentRef.current.sheets.find(
      (sheet) => sheet.id === sourceSheetId && sheet.hide !== 1,
    );
    if (!sourceSheet) return false;
    const invoker =
      document.activeElement instanceof HTMLElement
        ? document.activeElement
        : spreadsheetGridFocusTarget(spreadsheetCanvasRef.current);
    const selection = workbookRef.current?.getSelection()?.at(-1) ??
      sourceSheet.luckysheet_select_save?.at(-1) ?? {
        row: [0, 0],
        column: [0, 0],
      };
    let navigated = false;
    void officeDialog
      .prompt({
        title: spreadsheetCommandCatalog.goTo.label,
        description:
          '输入单元格、连续区域或已定义名称；可用工作表名称限定目标。',
        fieldLabel: '引用位置',
        initialValue: spreadsheetSelectionReference(selection),
        placeholder: '例如 A1、B2:D8 或 MyRange',
        confirmLabel: spreadsheetCommandCatalog.goTo.label,
        required: '请输入单元格、连续区域或已定义名称。',
        validate: (candidate) =>
          spreadsheetGoToValidationMessage(
            contentRef.current,
            sourceSheetId,
            candidate,
          ),
        restoreFocusTarget: () =>
          navigated
            ? spreadsheetGridFocusTarget(spreadsheetCanvasRef.current)
            : invoker?.isConnected
              ? invoker
              : spreadsheetGridFocusTarget(spreadsheetCanvasRef.current),
      })
      .then((value) => {
        if (value === null) return;
        const resolution = resolveSpreadsheetGoToTarget(
          contentRef.current,
          sourceSheetId,
          value,
        );
        if (!resolution.ok) {
          showToast(resolution.message, 'error');
          return;
        }
        navigated = navigateToSpreadsheetRange(
          resolution.target.sheetId,
          [resolution.target.selection],
          {
            row: resolution.target.selection.row_focus,
            column: resolution.target.selection.column_focus,
          },
        );
        if (!navigated) showToast('无法定位到所选区域。', 'error');
      });
    return true;
  }, [navigateToSpreadsheetRange, officeDialog]);
  const spreadsheetExtensions = useMemo(createSpreadsheetEditorExtensions, []);
  const spreadsheetEditor = useOfficeEditorRuntime(
    {
      activeSheetId,
      autoFilter,
      calculation,
      clipboard: {
        canCopySelection: !preview && Boolean(toolbarSheet),
        canCutSelection: !preview && Boolean(toolbarSheet && workbookInstance),
        canPasteSelection:
          !preview && Boolean(toolbarSheet && workbookInstance),
        copySelection: () => runSpreadsheetClipboardCopy(false),
        cutSelection: () => runSpreadsheetClipboardCopy(true),
        pasteSelection: runSpreadsheetClipboardPaste,
      },
      content: contentRef.current,
      editable: !preview,
      fallbackRange: selectedRange,
      formulaBar: {
        setValue: (value) => {
          const formulaBar =
            spreadsheetCanvasRef.current?.querySelector<HTMLElement>(
              '.fortune-fx-input',
            );
          if (formulaBar)
            formulaBar.textContent = value == null ? '' : String(value);
        },
      },
      formatPainter,
      formatCells: {
        canOpen: !preview && formatCellsSource === null,
        open: openSpreadsheetFormatCells,
      },
      history,
      navigation: {
        canOpenFind: !preview,
        canOpenGoTo: !preview,
        openFind: openSpreadsheetFind,
        openGoTo: openSpreadsheetGoTo,
      },
      onChange: (next) => {
        const formatCellsSelection = formatCellsApplyingRef.current
          ? formatCellsSelectionRef.current
          : null;
        const controlled =
          formatCellsSelection && !collaborationView
            ? spreadsheetContentWithSelection(
                next,
                formatCellsSelection.sheetId,
                formatCellsSelection.selection,
              )
            : next;
        contentRef.current = controlled;
        onChange(controlled);
      },
      selection: selectionState,
      targetSheetGridSize,
      targetSheetId: toolbarSheetId,
      toolbarCell,
      view:
        preview || collaborationView
          ? { activateSheet: activateLocalSpreadsheetSheet }
          : null,
      workbook: workbookInstance,
    },
    spreadsheetExtensions,
  );
  const spreadsheetCommands = spreadsheetEditor.commands;
  spreadsheetCommandsRef.current = spreadsheetCommands;
  formatPainterSelectionHandlerRef.current = (sheetId, selection) => {
    spreadsheetCommandsRef.current?.applyFormatPainter({
      sheetId,
      selection,
    });
  };
  const spreadsheetCan = spreadsheetEditor.can();
  const navigateToSpreadsheetParticipant = useCallback(
    (participant: WorkOfficeCollaborationParticipant): boolean => {
      const location = participant.location;
      const workbook = workbookRef.current;
      if (!workbook || location?.kind !== 'spreadsheet') return false;
      const sheet = materializedContent.sheets.find(
        (candidate) =>
          candidate.id === location.sheetId && candidate.hide !== 1,
      );
      if (!sheet) return false;
      const ranges = location.ranges
        .filter((range) => spreadsheetPresenceRangeExists(sheet, range))
        .map((range) => ({
          row: [range.startRow, range.endRow],
          column: [range.startColumn, range.endColumn],
        }));
      if (!ranges.length) return false;
      const activeCell = location.activeCell ?? {
        row: ranges[0]?.row[0] ?? 0,
        column: ranges[0]?.column[0] ?? 0,
      };
      if (!spreadsheetPresenceCellExists(sheet, activeCell)) return false;
      return navigateToSpreadsheetRange(location.sheetId, ranges, activeCell);
    },
    [materializedContent.sheets, navigateToSpreadsheetRange],
  );
  useOfficeCollaborationLocationNavigator(navigateToSpreadsheetParticipant);
  const restoreSpreadsheetGridFocus = useCallback(
    () => focusSpreadsheetGrid(spreadsheetCanvasRef.current),
    [],
  );
  const closeSpreadsheetFind = useCallback(() => {
    setFindOpen(false);
    focusSpreadsheetGrid(spreadsheetCanvasRef.current);
  }, []);
  const selectSpreadsheetFindMatch = useCallback(
    (match: SpreadsheetFindMatch) => {
      navigateToSpreadsheetRange(
        match.sheetId,
        [
          {
            row: [match.row, match.row],
            column: [match.column, match.column],
          },
        ],
        { row: match.row, column: match.column },
        false,
      );
    },
    [navigateToSpreadsheetRange],
  );
  const spreadsheetRibbonCommands = useMemo(
    () =>
      spreadsheetCommandsWithGridFocus(
        spreadsheetCommands,
        restoreSpreadsheetGridFocus,
      ),
    [restoreSpreadsheetGridFocus, spreadsheetCommands],
  );
  const restoreSpreadsheetShortcutFocus = useCallback(
    (event: KeyboardEvent) => {
      event.stopPropagation();
      if (spreadsheetNavigationShortcutOwnsFocus(event)) return;
      if (
        (event.metaKey || event.ctrlKey) &&
        !event.altKey &&
        !event.shiftKey &&
        event.key === '1'
      ) {
        return;
      }
      if (!isSpreadsheetCellEditingTarget(event.target)) {
        focusSpreadsheetGrid(spreadsheetCanvasRef.current);
      }
    },
    [],
  );
  useOfficeEditorKeyboardShortcuts(spreadsheetEditor, {
    capture: true,
    onHandled: restoreSpreadsheetShortcutFocus,
    scopeRef: spreadsheetRootRef,
  });
  const changeSpreadsheetWheelZoom = (direction: 'in' | 'out') => {
    const current = spreadsheetZoomRef.current;
    const next = stepOfficeZoom(current, direction);
    if (next === current) return;
    spreadsheetZoomRef.current = next;
    const sheetId = activeSheetIdRef.current;
    if (sheetId) collaborationView?.setZoom(sheetId, next / 100);
    if (previewRef.current) setPreviewZoom(next);
    else spreadsheetCommandsRef.current?.setZoom(next);
  };
  useOfficeEditorWheelZoom({
    scopeRef: spreadsheetRootRef,
    onZoomIn: () => changeSpreadsheetWheelZoom('in'),
    onZoomOut: () => changeSpreadsheetWheelZoom('out'),
  });
  const handleSpreadsheetEditingEscape = (
    event: React.KeyboardEvent<HTMLElement>,
  ) => {
    if (
      event.key === 'Escape' &&
      isSpreadsheetCellEditingTarget(event.target)
    ) {
      requestAnimationFrame(() =>
        focusSpreadsheetGrid(spreadsheetCanvasRef.current),
      );
    }
  };
  const handleSpreadsheetKeyDownCapture = (
    event: React.KeyboardEvent<HTMLElement>,
  ) => {
    if (
      event.key === 'Alt' &&
      !event.repeat &&
      event.target instanceof Element &&
      event.target.closest('.fortune-sheet-overlay') &&
      reserveAutoFilterAltKey()
    ) {
      // Fortune moves focus into its hidden cell editor on a bare Alt keydown.
      // Keep the header grid focused so the following ArrowDown reaches the
      // WPS-compatible AutoFilter shortcut.
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    handleSpreadsheetEditingEscape(event);
  };
  const handleSpreadsheetCopy = (
    event: ReactClipboardEvent<HTMLElement>,
    cut: boolean,
  ) => {
    if (isSpreadsheetNativeTextUndoTarget(event.target)) return;
    const selection = currentClipboardSelection();
    if (!selection) return;
    event.clipboardData.setData('text/plain', selection.clipboard);
    event.preventDefault();
    event.stopPropagation();
    if (cut) {
      spreadsheetCommands.clearSelectedCells();
      focusSpreadsheetGrid(spreadsheetCanvasRef.current);
    }
  };
  const handleSpreadsheetPaste = (event: ReactClipboardEvent<HTMLElement>) => {
    if (previewRef.current || isSpreadsheetNativeTextUndoTarget(event.target)) {
      return;
    }
    const values = parseSpreadsheetClipboardText(
      event.clipboardData.getData('text/plain'),
    );
    if (!values.length || !spreadsheetCommands.pasteCells(values)) return;
    event.preventDefault();
    event.stopPropagation();
    focusSpreadsheetGrid(spreadsheetCanvasRef.current);
  };
  const runSheetCommand = (command: () => boolean) => {
    command();
    focusSpreadsheetGrid(spreadsheetCanvasRef.current);
  };
  const openSpreadsheetContextMenu = (
    event: WorkspaceContextMenuEvent,
    requestedKind?: SpreadsheetStructureAxis,
  ): boolean => {
    if (preview) return false;
    const sheetId = selectionState?.sheetId ?? activeSheetIdRef.current;
    const sheet = contentRef.current.sheets.find(
      (candidate) => candidate.id === sheetId,
    );
    const savedSelection =
      selectionState?.selection ?? sheet?.luckysheet_select_save?.at(-1);
    const liveSelection = workbookRef.current?.getSelection()?.at(-1);
    const selection = liveSelection
      ? ({ ...savedSelection, ...liveSelection } as Selection)
      : savedSelection;
    if (!selection) return false;
    const agentSelection = spreadsheetAgentSelection(
      contentRef.current,
      sheetId,
      selection,
    );
    if (!agentSelection) return false;
    event.preventDefault();
    event.stopPropagation();
    const position = workspaceContextMenuPosition(
      event,
      spreadsheetSelectionContextMenuBounds(spreadsheetCanvasRef.current),
    );
    setContextMenu({
      kind:
        requestedKind ?? spreadsheetContextMenuKind(event.target, selection),
      range: spreadsheetSingleRange(selection),
      x: position.x,
      y: position.y,
      selection: agentSelection,
    });
    return true;
  };
  const requestSpreadsheetStructureSize = async (
    axis: SpreadsheetStructureAxis,
  ) => {
    const menu = contextMenu;
    if (!menu) return;
    const row = axis === 'row';
    const label = row ? '行高' : '列宽';
    const maximum = row ? 545 : 2_038;
    const currentSize = spreadsheetStructureSize(
      workbookRef.current,
      axis,
      menu.range,
      menu.selection.sheetId,
    );
    const value = await officeDialog.prompt({
      title: `设置${label}`,
      fieldLabel: `${label}（1–${maximum} 像素）`,
      initialValue: currentSize === null ? '' : String(currentSize),
      inputMode: 'numeric',
      confirmLabel: '应用',
      restoreFocusTarget: () =>
        spreadsheetGridFocusTarget(spreadsheetCanvasRef.current),
      required: `请输入${label}。`,
      validate: (candidate) => {
        const size = Number(candidate);
        return Number.isInteger(size) && size >= 1 && size <= maximum
          ? null
          : `${label}需为 1–${maximum} 之间的整数。`;
      },
    });
    if (value === null) return;
    if (!spreadsheetCommands.setSelectedStructureSize(axis, Number(value))) {
      showToast(`无法设置${label}。`, 'error');
    }
  };
  return (
    <section
      ref={spreadsheetRootRef}
      className={`work-spreadsheet-editor ${preview ? 'preview' : ''}`}
      data-auto-filter={autoFilterActive ? 'active' : undefined}
      data-column-count={targetSheetGridSize?.columnCount}
      data-format-painter={formatPainterMode ?? undefined}
      data-freeze-panes={toolbarSheet?.frozen ? 'active' : undefined}
      data-populated-cell-count={activeSheetProfile?.populatedCellCount}
      data-profile-column-count={activeSheetProfile?.columnCount}
      data-profile-row-count={activeSheetProfile?.rowCount}
      data-row-count={targetSheetGridSize?.rowCount}
      data-spreadsheet-profile={
        activeSheetProfile?.fortuneReady ? 'ready' : undefined
      }
      aria-label="表格工作区"
      onKeyDownCapture={handleSpreadsheetKeyDownCapture}
      onCopyCapture={(event) => handleSpreadsheetCopy(event, false)}
      onCutCapture={(event) => handleSpreadsheetCopy(event, true)}
      onPasteCapture={handleSpreadsheetPaste}
    >
      {preview && (
        <WorkOfficePreviewBar
          ariaLabel="表格预览工具"
          label="只读预览"
          detail={`${content.sheets.length} 个工作表`}
          fileActions={fileActions}
          className="work-spreadsheet-ribbon"
        />
      )}
      {!preview && (
        <SpreadsheetEditorRibbon
          activeTab={ribbonTab}
          autoFilterActive={autoFilterActive}
          can={spreadsheetCan}
          commands={spreadsheetRibbonCommands}
          content={content}
          fileActions={fileActions}
          findOpen={findOpen}
          formatPainterMode={formatPainterMode}
          freezePanesActive={Boolean(toolbarSheet?.frozen)}
          freezePanesSelection={toolbarSelection}
          gridLinesVisible={gridLinesVisible}
          panelId={panelId}
          onTabChange={(tab) => {
            setRibbonTab(tab);
            if (panel) closeWorkbookPanel();
          }}
          onTogglePanel={(nextPanel, trigger) => {
            if (panel === nextPanel) {
              closeWorkbookPanel();
              return;
            }
            panelTriggerRef.current = trigger;
            setPanel(nextPanel);
          }}
          panel={panel}
          toolbarCell={toolbarCell}
          toolbarCellBorders={toolbarCellBorders}
        />
      )}
      <output className="sr-only" aria-live="polite" aria-atomic="true">
        {[
          spreadsheetFormatPainterStatus(formatPainterMode),
          autoFilterStatus,
          spreadsheetFreezePanesStatus(toolbarSheet?.frozen),
        ]
          .filter(Boolean)
          .join(' ')}
      </output>
      <div className="work-spreadsheet-workspace">
        <div
          ref={spreadsheetCanvasRef}
          className="work-spreadsheet-canvas"
          onClickCapture={(event) => {
            if (!spreadsheetHeaderMenuTrigger(event.target)) return;
            openSpreadsheetContextMenu(event, 'column');
          }}
          onContextMenuCapture={openSpreadsheetContextMenu}
          onKeyDownCapture={(event) => {
            if (
              spreadsheetHeaderMenuTrigger(event.target) &&
              (event.key === 'Enter' || event.key === ' ')
            ) {
              openSpreadsheetContextMenu(event, 'column');
              return;
            }
            if (!isWorkspaceContextMenuKeyboardEvent(event)) return;
            openSpreadsheetContextMenu(event);
          }}
        >
          <ControlledWorkbook
            ref={bindWorkbookInstance}
            key={`spreadsheet:${workbookMountRevision}:${preview ? `preview-${previewZoom}` : 'edit'}:${conditionalFormatKey}:${protectionKey}:${chartPreviewKey}`}
            data={displayedWorkbookSheets}
            lang="zh"
            allowEdit={!preview}
            showToolbar={false}
            showFormulaBar
            showSheetTabs={false}
            row={60}
            column={26}
            rowHeaderWidth={44}
            columnHeaderHeight={24}
            defaultRowHeight={24}
            defaultColWidth={96}
            defaultFontSize={11}
            hooks={workbookHooks}
            onChange={handleWorkbookChange}
            onOp={handleWorkbookOperations}
          />
          {findOpen && (
            <SpreadsheetFindBar
              sheet={activeSheet}
              focusRequest={findFocusRequest}
              onClose={closeSpreadsheetFind}
              onSelectMatch={selectSpreadsheetFindMatch}
            />
          )}
        </div>
        {!preview && panel && (
          <SpreadsheetWorkbookPanel
            id={panelId}
            content={materializedContent}
            view={panel}
            activeSheetId={panelSheetId}
            selection={
              selectionState?.sheetId === panelSheetId
                ? selectionState.selection
                : undefined
            }
            can={spreadsheetCan}
            commands={spreadsheetCommands}
            restoreFocusTarget={() => panelTriggerRef.current}
            onClose={closeWorkbookPanel}
          />
        )}
      </div>
      <div className="work-spreadsheet-footer">
        <SpreadsheetSheetBar
          activeSheetId={activeSheetId}
          editable={!preview}
          sheets={materializedContent.sheets}
          onActivate={(sheetId) =>
            runSheetCommand(() => spreadsheetCommands.activateSheet(sheetId))
          }
          onCreate={() => runSheetCommand(() => spreadsheetCommands.addSheet())}
          onDelete={(sheetId) =>
            runSheetCommand(() => spreadsheetCommands.deleteSheet(sheetId))
          }
          onDuplicate={(sheetId) =>
            runSheetCommand(() => spreadsheetCommands.duplicateSheet(sheetId))
          }
          onHide={(sheetId) =>
            runSheetCommand(() => spreadsheetCommands.hideSheet(sheetId))
          }
          onMove={(sheetId, direction) =>
            runSheetCommand(() =>
              spreadsheetCommands.moveSheet(sheetId, direction),
            )
          }
          onRename={(sheetId, name) =>
            runSheetCommand(() =>
              spreadsheetCommands.renameSheet(sheetId, name),
            )
          }
          onSetColor={(sheetId, color) =>
            runSheetCommand(() =>
              spreadsheetCommands.setSheetColor(sheetId, color),
            )
          }
          onShow={(sheetId) =>
            runSheetCommand(() => spreadsheetCommands.activateSheet(sheetId))
          }
        />
        <WorkOfficeStatusBar
          className="work-spreadsheet-status"
          controls={
            <>
              <span
                className="work-spreadsheet-view-mode"
                role="img"
                aria-label="普通表格视图"
                title="普通表格视图"
              >
                <Grid3X3 size={13} />
              </span>
              <span className="work-office-status-divider" />
              <WorkOfficeZoomControls
                zoom={preview ? previewZoom : zoom}
                decreaseLabel="缩小表格"
                increaseLabel="放大表格"
                outputLabel="表格缩放比例"
                sliderLabel="表格缩放"
                onChange={(nextZoom) => {
                  const sheetId = activeSheetIdRef.current;
                  if (sheetId)
                    collaborationView?.setZoom(sheetId, nextZoom / 100);
                  if (preview) setPreviewZoom(nextZoom);
                  else spreadsheetCommands.setZoom(nextZoom);
                }}
              />
            </>
          }
        >
          <output aria-label="表格选区状态">
            {selectionState
              ? spreadsheetSelectionReference(selectionState.selection)
              : '未选择单元格'}
          </output>
          {selectionSummary && selectionSummary.nonEmptyCount > 0 && (
            <output aria-label="表格选区统计">
              {spreadsheetSelectionSummaryText(selectionSummary)}
            </output>
          )}
          {!preview && (
            <output
              aria-label="表格保存状态"
              className="work-office-save-status"
            >
              <Cloud size={12} />
              {saveStatus}
            </output>
          )}
        </WorkOfficeStatusBar>
      </div>
      {contextMenu && (
        <WorkspaceContextMenu
          label={spreadsheetContextMenuLabel(contextMenu)}
          className="work-spreadsheet-context-menu"
          x={contextMenu.x}
          y={contextMenu.y}
          items={[
            ...spreadsheetCoreContextMenuItems({
              can: spreadsheetCan,
              commands: spreadsheetCommands,
              selection: contextMenu.selection,
            }),
            ...(contextMenu.kind === 'cells'
              ? spreadsheetSortContextMenuItems({
                  can: spreadsheetCan,
                  commands: spreadsheetCommands,
                  separatorBefore: true,
                })
              : spreadsheetStructureContextMenuItems({
                  axis: contextMenu.kind,
                  can: spreadsheetCan,
                  commands: spreadsheetCommands,
                  onResize: (axis) => {
                    void requestSpreadsheetStructureSize(axis);
                  },
                }).map((item, index) =>
                  index === 0 ? { ...item, separatorBefore: true } : item,
                )),
            ...(contextMenu.kind === 'cells' && onAgentRequest
              ? spreadsheetAgentMenuItems(
                  contextMenu.selection,
                  onAgentRequest,
                  (changes) => {
                    const outcome = applySpreadsheetAgentProposalChanges(
                      contentRef.current,
                      contextMenu.selection.sheetId,
                      changes,
                    );
                    if (outcome.result.appliedTargetIds.length)
                      spreadsheetCommands.setSpreadsheetContent(
                        outcome.content,
                      );
                    return outcome.result;
                  },
                ).map((item, index) =>
                  index === 0 ? { ...item, separatorBefore: true } : item,
                )
              : []),
          ]}
          onClose={() => setContextMenu(null)}
          onRestoreFocus={() =>
            focusSpreadsheetGrid(spreadsheetCanvasRef.current)
          }
        />
      )}
      {formatCellsSource && (
        <SpreadsheetFormatCellsDialog
          source={formatCellsSource}
          restoreFocusTarget={() =>
            formatCellsInvokerRef.current?.isConnected
              ? formatCellsInvokerRef.current
              : spreadsheetGridFocusTarget(spreadsheetCanvasRef.current)
          }
          onApply={(patch) => {
            formatCellsApplyingRef.current = true;
            let handled = false;
            try {
              handled = spreadsheetCommands.applyCellFormat({
                sheetId: formatCellsSource.sheetId,
                range: formatCellsSource.range,
                patch,
              });
            } finally {
              formatCellsApplyingRef.current = false;
            }
            if (!handled) showToast('无法应用单元格格式。', 'error');
            return handled;
          }}
          onClose={closeSpreadsheetFormatCells}
        />
      )}
      {officeDialog.dialog}
    </section>
  );
}

export function focusSpreadsheetGrid(
  container: HTMLElement | null,
  {
    focusOrigin = document.activeElement,
    forceInitial = true,
  }: {
    focusOrigin?: Element | null;
    forceInitial?: boolean;
  } = {},
): void {
  if (!container) return;
  spreadsheetFocusCleanups.get(container)?.();

  let remainingFrames = spreadsheetFocusRetryFrames;
  let lastFocusedTarget: HTMLElement | null = null;
  const commandTrigger = focusOrigin;
  const initialFocusTarget = spreadsheetGridFocusTarget(container);
  const initialFocusTargetReady =
    spreadsheetGridFocusTargetReady(initialFocusTarget);
  let focusObserver: MutationObserver | null = null;
  let focusObservationTimeout: number | null = null;
  let focusOutHandler: ((event: FocusEvent) => void) | null = null;
  let pointerDownHandler: ((event: PointerEvent) => void) | null = null;
  let tabKeyHandler: ((event: KeyboardEvent) => void) | null = null;
  let focusRestorationStopped = false;

  const stopObservingFocusTarget = () => {
    focusRestorationStopped = true;
    focusObserver?.disconnect();
    focusObserver = null;
    if (focusOutHandler) {
      container.removeEventListener('focusout', focusOutHandler);
      focusOutHandler = null;
    }
    if (pointerDownHandler) {
      document.removeEventListener('pointerdown', pointerDownHandler, true);
      pointerDownHandler = null;
    }
    if (tabKeyHandler) {
      document.removeEventListener('keydown', tabKeyHandler, true);
      tabKeyHandler = null;
    }
    if (focusObservationTimeout !== null) {
      window.clearTimeout(focusObservationTimeout);
      focusObservationTimeout = null;
    }
    if (spreadsheetFocusCleanups.get(container) === stopObservingFocusTarget) {
      spreadsheetFocusCleanups.delete(container);
    }
  };

  const restoreFocus = (force: boolean) => {
    if (focusRestorationStopped) return;
    const activeElement = document.activeElement;
    if (!force && isSpreadsheetCellEditingTarget(activeElement)) {
      stopObservingFocusTarget();
      return;
    }
    if (
      !force &&
      activeElement !== document.body &&
      activeElement !== document.documentElement &&
      activeElement !== commandTrigger &&
      activeElement !== lastFocusedTarget
    ) {
      stopObservingFocusTarget();
      return;
    }

    const focusTarget = spreadsheetGridFocusTarget(container);
    if (focusTarget) {
      focusTarget.focus({ preventScroll: true });
      lastFocusedTarget = focusTarget;
    }
    if (remainingFrames <= 0) return;
    remainingFrames -= 1;
    requestAnimationFrame(() => restoreFocus(false));
  };

  focusOutHandler = (event) => {
    if (event.target !== lastFocusedTarget) return;
    requestAnimationFrame(() => restoreFocus(false));
  };
  container.addEventListener('focusout', focusOutHandler);

  pointerDownHandler = (event) => {
    if (event.target instanceof Node && !container.contains(event.target)) {
      stopObservingFocusTarget();
    }
  };
  tabKeyHandler = (event) => {
    if (event.key === 'Tab') stopObservingFocusTarget();
  };
  document.addEventListener('pointerdown', pointerDownHandler, true);
  document.addEventListener('keydown', tabKeyHandler, true);

  if (typeof MutationObserver !== 'undefined') {
    focusObserver = new MutationObserver(() => {
      const focusTarget = spreadsheetGridFocusTarget(container);
      const focusTargetReady = spreadsheetGridFocusTargetReady(focusTarget);
      const mountedOrReplaced = focusTarget !== initialFocusTarget;
      const initialTargetBecameReady =
        focusTarget === initialFocusTarget &&
        !initialFocusTargetReady &&
        focusTargetReady;
      if (!focusTarget || (!mountedOrReplaced && !initialTargetBecameReady)) {
        return;
      }
      restoreFocus(false);
    });
    focusObserver.observe(container, {
      attributes: true,
      attributeFilter: ['style'],
      childList: true,
      subtree: true,
    });
  }
  focusObservationTimeout = window.setTimeout(
    stopObservingFocusTarget,
    spreadsheetFocusObservationMs,
  );
  spreadsheetFocusCleanups.set(container, stopObservingFocusTarget);

  restoreFocus(forceInitial);
}

export function spreadsheetCommandsWithGridFocus(
  commands: SpreadsheetEditorCommands,
  restoreFocus: () => void,
): SpreadsheetEditorCommands {
  const afterSuccessfulCommand =
    <Arguments extends unknown[]>(command: (...args: Arguments) => boolean) =>
    (...args: Arguments): boolean => {
      const handled = command(...args);
      if (handled) restoreFocus();
      return handled;
    };

  return {
    ...commands,
    adjustDecimalPlaces: afterSuccessfulCommand(commands.adjustDecimalPlaces),
    applyAutoSum: afterSuccessfulCommand(commands.applyAutoSum),
    activateFormatPainter: afterSuccessfulCommand(
      commands.activateFormatPainter,
    ),
    cancelFormatPainter: afterSuccessfulCommand(commands.cancelFormatPainter),
    clearSelectedCells: afterSuccessfulCommand(commands.clearSelectedCells),
    copySelection: afterSuccessfulCommand(commands.copySelection),
    cutSelection: afterSuccessfulCommand(commands.cutSelection),
    deleteSelectedStructure: afterSuccessfulCommand(
      commands.deleteSelectedStructure,
    ),
    fillSelectedCells: afterSuccessfulCommand(commands.fillSelectedCells),
    insertSelectedStructure: afterSuccessfulCommand(
      commands.insertSelectedStructure,
    ),
    mergeSelectedCells: afterSuccessfulCommand(commands.mergeSelectedCells),
    pasteSelection: afterSuccessfulCommand(commands.pasteSelection),
    redo: afterSuccessfulCommand(commands.redo),
    setCellFormat: afterSuccessfulCommand(commands.setCellFormat),
    setFreezePanes: afterSuccessfulCommand(commands.setFreezePanes),
    setGridLines: afterSuccessfulCommand(commands.setGridLines),
    setSelectedCellBorders: afterSuccessfulCommand(
      commands.setSelectedCellBorders,
    ),
    toggleAutoFilter: afterSuccessfulCommand(commands.toggleAutoFilter),
    undo: afterSuccessfulCommand(commands.undo),
  };
}

function ignoreSpreadsheetControlledHistory(): boolean {
  return true;
}

function spreadsheetNavigationShortcutOwnsFocus(event: KeyboardEvent): boolean {
  if (event.altKey || event.shiftKey) return false;
  const key = event.key.toLocaleLowerCase();
  if ((event.metaKey || event.ctrlKey) && key === 'f') return true;
  if (event.ctrlKey && !event.metaKey && key === 'g') return true;
  return !event.ctrlKey && !event.metaKey && event.key === 'F5';
}

function spreadsheetFormatPainterStatus(
  mode: SpreadsheetFormatPainterMode | null,
): string {
  if (mode === 'locked') {
    return '格式刷已锁定，可连续选择目标区域；再次点击格式刷或按 Escape 退出。';
  }
  if (mode === 'once') {
    return '格式刷已开启，请选择一个目标区域；按 Escape 退出。';
  }
  return '';
}

function spreadsheetGridFocusTarget(
  container: HTMLElement | null,
): HTMLElement | null {
  return (
    container?.querySelector<HTMLElement>('.fortune-sheet-overlay') ??
    container?.querySelector<HTMLElement>('.fortune-cell-area') ??
    null
  );
}

function spreadsheetHeaderMenuTrigger(target: EventTarget | null) {
  return target instanceof Element
    ? target.closest<HTMLElement>('.header-arrow')
    : null;
}

function spreadsheetContextMenuKind(
  target: EventTarget | null,
  selection: Selection,
): SpreadsheetContextMenuState['kind'] {
  if (target instanceof Element) {
    if (target.closest('.fortune-col-header')) return 'column';
    if (target.closest('.fortune-row-header')) return 'row';
  }
  if (selection.column_select) return 'column';
  if (selection.row_select) return 'row';
  return 'cells';
}

function spreadsheetContextMenuLabel({
  kind,
  selection,
}: SpreadsheetContextMenuState): string {
  const subject = kind === 'row' ? '行' : kind === 'column' ? '列' : '表格选区';
  return `${subject} ${selection.reference} 操作`;
}

function spreadsheetStructureSize(
  workbook: WorkbookInstance | null,
  axis: SpreadsheetStructureAxis,
  range: SpreadsheetCommandRange,
  sheetId: string,
): number | null {
  if (!workbook) return null;
  const values = axis === 'row' ? range.row : range.column;
  const start = Math.min(values[0] ?? 0, values[1] ?? 0);
  const end = Math.max(values[0] ?? 0, values[1] ?? 0);
  const indices = Array.from(
    { length: end - start + 1 },
    (_, offset) => start + offset,
  );
  try {
    const sizes =
      axis === 'row'
        ? workbook.getRowHeight(indices, { id: sheetId })
        : workbook.getColumnWidth(indices, { id: sheetId });
    const unique = new Set(
      indices.map((index) => sizes[index]).filter(Number.isFinite),
    );
    return unique.size === 1 ? (unique.values().next().value ?? null) : null;
  } catch {
    return null;
  }
}

function spreadsheetSelectionContextMenuBounds(
  container: HTMLElement | null,
): DOMRect | null {
  const selections = [
    ...(container?.querySelectorAll<HTMLElement>('.luckysheet-cell-selected') ??
      []),
  ];
  for (let index = selections.length - 1; index >= 0; index -= 1) {
    const bounds = selections[index]?.getBoundingClientRect();
    if (bounds && bounds.width > 0 && bounds.height > 0) return bounds;
  }
  return null;
}

function spreadsheetGridFocusTargetReady(target: HTMLElement | null): boolean {
  if (!target) return false;
  const bounds = target.getBoundingClientRect();
  return bounds.width > 0 && bounds.height > 0;
}

function spreadsheetPresenceRangeExists(
  sheet: WorkSpreadsheetContent['sheets'][number],
  range: WorkOfficeSpreadsheetPresenceRange,
): boolean {
  const { columns, rows } = spreadsheetPresenceSheetSize(sheet);
  return (
    range.startRow >= 0 &&
    range.startColumn >= 0 &&
    range.endRow < rows &&
    range.endColumn < columns
  );
}

function spreadsheetPresenceCellExists(
  sheet: WorkSpreadsheetContent['sheets'][number],
  cell: WorkOfficeSpreadsheetPresenceCell,
): boolean {
  const { columns, rows } = spreadsheetPresenceSheetSize(sheet);
  return (
    cell.row >= 0 &&
    cell.column >= 0 &&
    cell.row < rows &&
    cell.column < columns
  );
}

function spreadsheetPresenceSheetSize(
  sheet: WorkSpreadsheetContent['sheets'][number],
): { columns: number; rows: number } {
  return {
    rows: Math.max(sheet.row ?? 60, sheet.data?.length ?? 0),
    columns: Math.max(
      sheet.column ?? 26,
      sheet.data?.reduce(
        (maximum, row) => Math.max(maximum, row?.length ?? 0),
        0,
      ) ?? 0,
    ),
  };
}

function spreadsheetSelectionSummaryText(
  summary: ReturnType<typeof spreadsheetSelectionSummary>,
): string {
  const parts = [`计数 ${summary.nonEmptyCount}`];
  if (summary.average !== null && summary.sum !== null) {
    parts.push(`平均 ${formatSpreadsheetStatistic(summary.average)}`);
    parts.push(`求和 ${formatSpreadsheetStatistic(summary.sum)}`);
  }
  return parts.join(' · ');
}

function formatSpreadsheetStatistic(value: number): string {
  return new Intl.NumberFormat('zh-CN', {
    maximumFractionDigits: 6,
  }).format(value);
}
