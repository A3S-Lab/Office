import type { Hooks, Op, Selection } from '@fortune-sheet/core';
import { Workbook, type WorkbookInstance } from '@fortune-sheet/react';
import { Cloud, Grid3X3 } from 'lucide-react';
import {
  type ClipboardEvent as ReactClipboardEvent,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import {
  isWorkspaceContextMenuKeyboardEvent,
  type WorkspaceContextMenuEvent,
  WorkspaceContextMenu,
  workspaceContextMenuPosition,
} from '../../workspace/components/workspace-context-menu';
import { spreadsheetAgentMenuItems } from '../components/work-editor-agent-menus';
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
import {
  reconcileSpreadsheetPivots,
  refreshSpreadsheetPivotTables,
  spreadsheetPivotIntersects,
  spreadsheetPivotOutputContains,
} from '../work-spreadsheet-pivots';
import { spreadsheetProtectionKey } from '../work-spreadsheet-protection';
import type { WorkSpreadsheetContent } from '../work-types';
import {
  createSpreadsheetEditorExtensions,
  type SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import {
  parseSpreadsheetClipboardText,
  spreadsheetCoreContextMenuItems,
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
  spreadsheetCellAt,
  spreadsheetContentWithSelection,
  spreadsheetSelectionReference,
  spreadsheetSelectionSummary,
  spreadsheetSheetsForFortune,
  spreadsheetSheetsWithFiniteSelections,
  spreadsheetSingleRange,
} from './spreadsheet-editor-support';
import { SpreadsheetSheetBar } from './spreadsheet-sheet-bar';
import {
  SpreadsheetWorkbookPanel,
  type SpreadsheetWorkbookPanelView,
} from './spreadsheet-workbook-panel';
import { useOfficeEditorKeyboardShortcuts } from './use-office-editor-keyboard-shortcuts';
import { useOfficeEditorRuntime } from './use-office-editor-runtime';
import { useOfficeHistory } from './use-office-history';
import { useSpreadsheetCalculation } from './use-spreadsheet-calculation';
import { useSpreadsheetWorkbookSync } from './use-spreadsheet-workbook-sync';
import {
  type WorkOfficeFileAction,
  WorkOfficePreviewBar,
  WorkOfficeStatusBar,
  WorkOfficeZoomControls,
} from './work-office-chrome';

export interface SpreadsheetEditorProps {
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

interface SpreadsheetSelectionState {
  sheetId: string;
  selection: Selection;
}

interface SpreadsheetContextMenuState {
  x: number;
  y: number;
  selection: WorkSpreadsheetAgentSelection;
}

export function SpreadsheetEditor({
  content,
  kernelWasmUrl,
  preview,
  saveStatus = '已自动保存',
  fileActions,
  onChange,
  onAgentRequest,
}: SpreadsheetEditorProps) {
  const materializedContent = useMemo(
    () => refreshSpreadsheetPivotTables(content),
    [content],
  );
  const contentRef = useRef(materializedContent);
  const spreadsheetCommandsRef = useRef<SpreadsheetEditorCommands | null>(null);
  const previewRef = useRef(preview);
  const spreadsheetRootRef = useRef<HTMLElement>(null);
  const spreadsheetCanvasRef = useRef<HTMLDivElement>(null);
  const workbookRef = useRef<WorkbookInstance>(null);
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
  const {
    acceptContent: acceptWorkbookContent,
    ignoreChangeDuringExternalSync,
    mountRevision: workbookMountRevision,
    recordOperations: recordWorkbookOperations,
    takeOperations: takeWorkbookOperations,
  } = useSpreadsheetWorkbookSync(materializedContent);
  const calculation = useSpreadsheetCalculation({
    content: materializedContent,
    kernelWasmUrl,
    workbookRef,
  });
  const [ribbonTab, setRibbonTab] = useState<SpreadsheetRibbonTabId>('home');
  const [panel, setPanel] = useState<SpreadsheetWorkbookPanelView | null>(null);
  const [selectionState, setSelectionState] =
    useState<SpreadsheetSelectionState | null>(null);
  const [contextMenu, setContextMenu] =
    useState<SpreadsheetContextMenuState | null>(null);
  const [previewZoom, setPreviewZoom] = useState(100);
  const [previewActiveSheetId, setPreviewActiveSheetId] = useState<
    string | null
  >(null);
  const contentActiveSheetId =
    content.sheets.find((sheet) => sheet.status === 1)?.id ??
    content.sheets.find((sheet) => !sheet.hide)?.id ??
    '';
  const activeSheetId =
    preview &&
    previewActiveSheetId &&
    content.sheets.some(
      (sheet) => sheet.id === previewActiveSheetId && sheet.hide !== 1,
    )
      ? previewActiveSheetId
      : contentActiveSheetId;
  const activeSheetIdRef = useRef(activeSheetId);
  const focusedSheetIdRef = useRef<string | null>(null);
  contentRef.current = materializedContent;
  previewRef.current = preview;
  const history = useOfficeHistory({
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
    sameValue: sameSpreadsheetHistoryContent,
  });
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
  useEffect(() => {
    setPreviewActiveSheetId(preview ? contentActiveSheetId : null);
  }, [contentActiveSheetId, preview]);
  useEffect(() => {
    activeSheetIdRef.current = activeSheetId;
  }, [activeSheetId]);
  useEffect(() => {
    if (!activeSheetId || focusedSheetIdRef.current === activeSheetId) return;
    focusedSheetIdRef.current = activeSheetId;
    const frame = requestAnimationFrame(() =>
      focusSpreadsheetGrid(spreadsheetCanvasRef.current),
    );
    return () => cancelAnimationFrame(frame);
  }, [activeSheetId]);
  useEffect(() => {
    if (!preview) return;
    setPanel(null);
    setContextMenu(null);
  }, [preview]);
  const workbookHooks = useMemo<Hooks>(
    () => ({
      afterActivateSheet: (id) => {
        activeSheetIdRef.current = id;
        if (previewRef.current) setPreviewActiveSheetId(id);
        setSelectionState(null);
      },
      afterSelectionChange: (sheetId, selection) => {
        setSelectionState({ sheetId, selection });
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
        const style = conditionalStylesBySheet
          .get(activeSheetIdRef.current)
          ?.get(`${cellInfo.row}_${cellInfo.column}`);
        if (style?.cellColor) context.fillStyle = style.cellColor;
        return true;
      },
      afterRenderCell: (cell, cellInfo, context) => {
        const style = conditionalStylesBySheet
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
    [conditionalStylesBySheet],
  );
  const conditionalFormatKey = content.sheets
    .map(
      (sheet) =>
        `${sheet.id}:${JSON.stringify(sheet.luckysheet_conditionformat_save ?? [])}`,
    )
    .join('|');
  const protectionKey = spreadsheetProtectionKey(content.sheets);
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
  const workbookSheetsRef = useRef(workbookSheets);
  workbookSheetsRef.current = displayedWorkbookSheets;
  const handleWorkbookChange = useCallback(
    (sheets: WorkSpreadsheetContent['sheets']) => {
      if (
        ignoreChangeDuringExternalSync(
          sameSpreadsheetWorkbookState(sheets, workbookSheetsRef.current),
        )
      ) {
        takeWorkbookOperations();
        return;
      }
      const operations = takeWorkbookOperations();
      if (
        previewRef.current ||
        (!operations.length &&
          !calculation.hasPendingResultPatches() &&
          sameSpreadsheetWorkbookState(sheets, workbookSheetsRef.current))
      )
        return;
      const withCharts = reconcileSpreadsheetChartPreviews(
        contentRef.current,
        sheets,
      );
      const next = reconcileSpreadsheetPivots(
        contentRef.current,
        withCharts.sheets,
      );
      contentRef.current = next;
      acceptWorkbookContent(next);
      calculation.synchronizeWorkbook(next, operations);
      spreadsheetCommandsRef.current?.setSpreadsheetContent(next);
    },
    [
      acceptWorkbookContent,
      calculation,
      ignoreChangeDuringExternalSync,
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
  const chartPreviewKey = workbookSheets
    .flatMap((sheet) =>
      (sheet.images ?? [])
        .filter((image) => image.id.startsWith('work-chart-preview-'))
        .map(
          (image) =>
            `${image.id}:${image.src}:${image.left}:${image.top}:${image.width}:${image.height}`,
        ),
    )
    .join('|');
  const panelSheetId =
    selectionState?.sheetId ?? activeSheetIdRef.current ?? activeSheetId;
  const activeSheet = materializedContent.sheets.find(
    (sheet) => sheet.id === activeSheetId,
  );
  const zoom = Math.round((activeSheet?.zoomRatio ?? 1) * 100);
  useEffect(() => {
    if (preview) setPreviewZoom(zoom);
  }, [preview, zoom]);
  const gridLinesVisible =
    activeSheet?.showGridLines !== false && activeSheet?.showGridLines !== 0;
  const toolbarSheetId = selectionState?.sheetId ?? activeSheetId;
  const toolbarSheet = workbookSheets.find(
    (sheet) => sheet.id === toolbarSheetId,
  );
  const toolbarSelection = finiteSpreadsheetSelection(
    selectionState?.selection ?? toolbarSheet?.luckysheet_select_save?.at(-1),
  );
  const toolbarCell = spreadsheetCellAt(
    toolbarSheet,
    toolbarSelection.row_focus ?? toolbarSelection.row[0],
    toolbarSelection.column_focus ?? toolbarSelection.column[0],
  );
  const selectedRange = spreadsheetSingleRange(toolbarSelection);
  const multipleCellsSelected =
    selectedRange.row[0] !== selectedRange.row[1] ||
    selectedRange.column[0] !== selectedRange.column[1];
  const selectionSummary = multipleCellsSelected
    ? spreadsheetSelectionSummary(toolbarSheet, toolbarSelection)
    : null;
  const activateReadOnlySpreadsheetSheet = useCallback((sheetId: string) => {
    if (!previewRef.current) return false;
    const sheet = contentRef.current.sheets.find(
      (candidate) => candidate.id === sheetId && candidate.hide !== 1,
    );
    if (!sheet) return false;
    activeSheetIdRef.current = sheetId;
    setSelectionState(null);
    setPreviewActiveSheetId(sheetId);
    try {
      workbookRef.current?.activateSheet({ id: sheetId });
    } catch {
      return false;
    }
    return true;
  }, []);
  const spreadsheetExtensions = useMemo(createSpreadsheetEditorExtensions, []);
  const spreadsheetEditor = useOfficeEditorRuntime(
    {
      activeSheetId,
      calculation,
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
      history,
      onChange: (next) => {
        contentRef.current = next;
        onChange(next);
      },
      selection: selectionState,
      targetSheetId: toolbarSheetId,
      toolbarCell,
      view: preview
        ? { activateSheet: activateReadOnlySpreadsheetSheet }
        : null,
      workbook: workbookInstance,
    },
    spreadsheetExtensions,
  );
  const spreadsheetCommands = spreadsheetEditor.commands;
  spreadsheetCommandsRef.current = spreadsheetCommands;
  const spreadsheetCan = spreadsheetEditor.can();
  const restoreSpreadsheetGridFocus = useCallback(
    () => focusSpreadsheetGrid(spreadsheetCanvasRef.current),
    [],
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
  ): boolean => {
    if (preview) return false;
    const sheetId = selectionState?.sheetId ?? activeSheetIdRef.current;
    const sheet = content.sheets.find((candidate) => candidate.id === sheetId);
    const selection =
      selectionState?.selection ?? sheet?.luckysheet_select_save?.at(-1);
    if (!selection) return false;
    const agentSelection = spreadsheetAgentSelection(
      content,
      sheetId,
      selection,
    );
    if (!agentSelection) return false;
    event.preventDefault();
    event.stopPropagation();
    const position = workspaceContextMenuPosition(event);
    setContextMenu({
      x: position.x,
      y: position.y,
      selection: agentSelection,
    });
    return true;
  };
  return (
    <section
      ref={spreadsheetRootRef}
      className={`work-spreadsheet-editor ${preview ? 'preview' : ''}`}
      aria-label="表格工作区"
      onKeyDownCapture={handleSpreadsheetEditingEscape}
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
          can={spreadsheetCan}
          commands={spreadsheetRibbonCommands}
          content={content}
          fileActions={fileActions}
          gridLinesVisible={gridLinesVisible}
          multipleCellsSelected={multipleCellsSelected}
          onTabChange={(tab) => {
            setRibbonTab(tab);
            setPanel(null);
          }}
          onTogglePanel={(nextPanel) =>
            setPanel((current) => (current === nextPanel ? null : nextPanel))
          }
          panel={panel}
          toolbarCell={toolbarCell}
        />
      )}
      {!preview && panel && (
        <SpreadsheetWorkbookPanel
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
          onClose={() => setPanel(null)}
        />
      )}
      <div
        ref={spreadsheetCanvasRef}
        className="work-spreadsheet-canvas"
        onContextMenuCapture={openSpreadsheetContextMenu}
        onKeyDownCapture={(event) => {
          if (!isWorkspaceContextMenuKeyboardEvent(event)) return;
          openSpreadsheetContextMenu(event);
        }}
      >
        <Workbook
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
          label={`表格选区 ${contextMenu.selection.reference} 操作`}
          className="work-spreadsheet-context-menu"
          x={contextMenu.x}
          y={contextMenu.y}
          items={[
            ...spreadsheetCoreContextMenuItems({
              can: spreadsheetCan,
              commands: spreadsheetCommands,
              selection: contextMenu.selection,
            }),
            ...(onAgentRequest
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
    </section>
  );
}

export function focusSpreadsheetGrid(container: HTMLElement | null): void {
  if (!container) return;
  spreadsheetFocusCleanups.get(container)?.();

  let remainingFrames = spreadsheetFocusRetryFrames;
  let lastFocusedTarget: HTMLElement | null = null;
  const commandTrigger = document.activeElement;
  const initialFocusTarget = spreadsheetGridFocusTarget(container);
  const initialFocusTargetReady =
    spreadsheetGridFocusTargetReady(initialFocusTarget);
  let focusObserver: MutationObserver | null = null;
  let focusObservationTimeout: number | null = null;
  let focusOutHandler: ((event: FocusEvent) => void) | null = null;

  const stopObservingFocusTarget = () => {
    focusObserver?.disconnect();
    focusObserver = null;
    if (focusOutHandler) {
      container.removeEventListener('focusout', focusOutHandler);
      focusOutHandler = null;
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
    const activeElement = document.activeElement;
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

  restoreFocus(true);
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
    redo: afterSuccessfulCommand(commands.redo),
    setCellFormat: afterSuccessfulCommand(commands.setCellFormat),
    setGridLines: afterSuccessfulCommand(commands.setGridLines),
    toggleCellMerge: afterSuccessfulCommand(commands.toggleCellMerge),
    undo: afterSuccessfulCommand(commands.undo),
  };
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

function spreadsheetGridFocusTargetReady(target: HTMLElement | null): boolean {
  if (!target) return false;
  const bounds = target.getBoundingClientRect();
  return bounds.width > 0 && bounds.height > 0;
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
