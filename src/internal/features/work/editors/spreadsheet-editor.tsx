import type { Hooks, Op, Selection } from '@fortune-sheet/core';
import { Workbook, type WorkbookInstance } from '@fortune-sheet/react';
import { Cloud, Grid3X3 } from 'lucide-react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { WorkspaceContextMenu } from '../../workspace/components/workspace-context-menu';
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
import { spreadsheetCoreContextMenuItems } from './spreadsheet-context-menu';
import {
  SpreadsheetEditorRibbon,
  type SpreadsheetRibbonTabId,
} from './spreadsheet-editor-ribbon';
import {
  finiteSpreadsheetSelection,
  sameSpreadsheetHistoryContent,
  sameSpreadsheetWorkbookState,
  spreadsheetCellAt,
  spreadsheetSelectionReference,
  spreadsheetSelectionSummary,
  spreadsheetSheetsForFortune,
  spreadsheetSheetsWithFiniteSelections,
  spreadsheetSingleRange,
} from './spreadsheet-editor-support';
import {
  SpreadsheetWorkbookPanel,
  type SpreadsheetWorkbookPanelView,
} from './spreadsheet-workbook-panel';
import { useOfficeHistory } from './use-office-history';
import { useOfficeEditorRuntime } from './use-office-editor-runtime';
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
  const spreadsheetCanvasRef = useRef<HTMLDivElement>(null);
  const workbookRef = useRef<WorkbookInstance>(null);
  const {
    acceptContent: acceptWorkbookContent,
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
  const history = useOfficeHistory({
    content,
    onChange,
    sameValue: sameSpreadsheetHistoryContent,
  });
  const activeSheetId =
    content.sheets.find((sheet) => sheet.status === 1)?.id ??
    content.sheets.find((sheet) => !sheet.hide)?.id ??
    '';
  const activeSheetIdRef = useRef(activeSheetId);
  contentRef.current = materializedContent;
  previewRef.current = preview;
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
    activeSheetIdRef.current = activeSheetId;
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
            zoomRatio: previewZoom / 100,
          }))
        : workbookSheets,
    [preview, previewZoom, workbookSheets],
  );
  const workbookSheetsRef = useRef(workbookSheets);
  workbookSheetsRef.current = displayedWorkbookSheets;
  const handleWorkbookChange = useCallback(
    (sheets: WorkSpreadsheetContent['sheets']) => {
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
    [acceptWorkbookContent, calculation, takeWorkbookOperations],
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
  const activeSheetIndex = Math.max(
    0,
    materializedContent.sheets.findIndex((sheet) => sheet.id === activeSheetId),
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
      workbook: workbookRef.current,
    },
    spreadsheetExtensions,
  );
  const spreadsheetCommands = spreadsheetEditor.commands;
  spreadsheetCommandsRef.current = spreadsheetCommands;
  const spreadsheetCan = spreadsheetEditor.can();
  const handleSpreadsheetShortcut = (
    event: React.KeyboardEvent<HTMLElement>,
  ) => {
    if (spreadsheetEditor.handleKeyDown(event.nativeEvent)) {
      event.stopPropagation();
    }
  };
  return (
    <section
      className={`work-spreadsheet-editor ${preview ? 'preview' : ''}`}
      aria-label="表格工作区"
      onKeyDownCapture={handleSpreadsheetShortcut}
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
          commands={spreadsheetCommands}
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
        onContextMenuCapture={(event) => {
          if (preview) return;
          const sheetId = selectionState?.sheetId ?? activeSheetIdRef.current;
          const sheet = content.sheets.find(
            (candidate) => candidate.id === sheetId,
          );
          const selection =
            selectionState?.selection ?? sheet?.luckysheet_select_save?.at(-1);
          if (!selection) return;
          const agentSelection = spreadsheetAgentSelection(
            content,
            sheetId,
            selection,
          );
          if (!agentSelection) return;
          event.preventDefault();
          event.stopPropagation();
          setContextMenu({
            x: event.clientX,
            y: event.clientY,
            selection: agentSelection,
          });
        }}
      >
        <Workbook
          ref={workbookRef}
          key={`spreadsheet:${workbookMountRevision}:${preview ? `preview-${previewZoom}` : 'edit'}:${conditionalFormatKey}:${protectionKey}:${chartPreviewKey}`}
          data={displayedWorkbookSheets}
          lang="zh"
          allowEdit={!preview}
          showToolbar={false}
          showFormulaBar
          showSheetTabs
          row={60}
          column={26}
          defaultRowHeight={24}
          defaultColWidth={96}
          hooks={workbookHooks}
          onChange={handleWorkbookChange}
          onOp={handleWorkbookOperations}
        />
      </div>
      <WorkOfficeStatusBar
        className="work-spreadsheet-status"
        controls={
          <>
            <button
              type="button"
              aria-label="普通表格视图"
              title="普通表格视图"
              aria-pressed="true"
            >
              <Grid3X3 size={13} />
            </button>
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
        <output aria-label="工作表状态">
          工作表 {activeSheetIndex + 1} / {materializedContent.sheets.length} ·{' '}
          {activeSheet?.name ?? '未命名'}
        </output>
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
          <output aria-label="表格保存状态" className="work-office-save-status">
            <Cloud size={12} />
            {saveStatus}
          </output>
        )}
      </WorkOfficeStatusBar>
      {contextMenu && (
        <WorkspaceContextMenu
          label={`表格选区 ${contextMenu.selection.reference} 操作`}
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
        />
      )}
    </section>
  );
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
