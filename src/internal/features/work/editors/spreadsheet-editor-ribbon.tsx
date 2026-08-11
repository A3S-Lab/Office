import type { Cell, Selection } from '@fortune-sheet/core';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  BarChart3,
  BetweenHorizontalEnd,
  BetweenHorizontalStart,
  BetweenVerticalEnd,
  BetweenVerticalStart,
  Bold,
  Bookmark,
  Calculator,
  ChevronDown,
  ClipboardPaste,
  Copy,
  DecimalsArrowLeft,
  DecimalsArrowRight,
  Columns3,
  Eraser,
  FileX2,
  Grid3X3,
  Italic,
  Link2Off,
  ListFilter,
  MessageSquareX,
  Merge,
  Paintbrush,
  Palette,
  Percent,
  PanelLeft,
  PanelTop,
  PanelsTopLeft,
  Printer,
  Redo2,
  RefreshCw,
  Scissors,
  Search,
  ShieldCheck,
  SortAsc,
  SortDesc,
  Rows3,
  TableProperties,
  Trash2,
  Underline,
  Undo2,
  WrapText,
  X,
} from 'lucide-react';
import { type ReactNode, useMemo } from 'react';
import { Popover } from '../../../design-system/primitives';
import { spreadsheetChartCount } from '../work-spreadsheet-charts';
import { spreadsheetFormulaCount } from '../work-spreadsheet-formula-analysis';
import { spreadsheetPivotCount } from '../work-spreadsheet-pivots';
import { protectedSheetCount } from '../work-spreadsheet-protection';
import type { WorkSpreadsheetContent } from '../work-types';
import { OfficeColorPicker, OfficeSelect } from './office-controls';
import type { SpreadsheetCellMergeCommand } from './spreadsheet-cell-merge';
import type { SpreadsheetCellClearMode } from './spreadsheet-cell-clear';
import {
  type SpreadsheetRibbonTabId,
  spreadsheetCommandCatalog,
  spreadsheetRibbonTabs,
} from './spreadsheet-command-catalog';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { managedConditionalFormatCount } from './spreadsheet-conditional-format-panel';
import {
  spreadsheetFontFamilyOptions,
  spreadsheetFontSizeOptions,
} from './spreadsheet-editor-support';
import type { SpreadsheetFormatPainterMode } from './spreadsheet-format-painter';
import {
  type SpreadsheetFreezePanePreset,
  spreadsheetFreezePanesSelectionLabel,
} from './spreadsheet-freeze-panes';
import {
  adjustSpreadsheetNumberFormat,
  type SpreadsheetNumberFormatPreset,
  spreadsheetNumberFormatCode,
  spreadsheetNumberFormatPreset,
  spreadsheetNumberFormatValue,
} from './spreadsheet-number-format';
import { spreadsheetPrintSettingCount } from './spreadsheet-print-settings-panel';
import type { SpreadsheetWorkbookPanelView } from './spreadsheet-workbook-panel';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import {
  type WorkOfficeFileAction,
  WorkOfficeRibbon,
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

const spreadsheetNumberFormatOptions: readonly {
  value: SpreadsheetNumberFormatPreset;
  label: string;
  disabled?: boolean;
}[] = [
  { value: 'general', label: '常规' },
  { value: 'number', label: '数字' },
  { value: 'percent', label: '百分比' },
  { value: 'custom', label: '自定义', disabled: true },
];

const defaultSpreadsheetFreezePanesSelection: Selection = {
  row: [0, 0],
  column: [0, 0],
  row_focus: 0,
  column_focus: 0,
};

export type { SpreadsheetRibbonTabId } from './spreadsheet-command-catalog';

export function SpreadsheetEditorRibbon({
  activeTab,
  autoFilterActive = false,
  can,
  commands,
  content,
  fileActions,
  findOpen,
  formatPainterMode = null,
  freezePanesActive = false,
  freezePanesSelection = defaultSpreadsheetFreezePanesSelection,
  gridLinesVisible,
  panelId,
  onOpenFind,
  onTabChange,
  onTogglePanel,
  panel,
  toolbarCell,
}: {
  activeTab: SpreadsheetRibbonTabId;
  autoFilterActive?: boolean;
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
  content: WorkSpreadsheetContent;
  fileActions?: readonly WorkOfficeFileAction[];
  findOpen: boolean;
  formatPainterMode?: SpreadsheetFormatPainterMode | null;
  freezePanesActive?: boolean;
  freezePanesSelection?: Selection;
  gridLinesVisible: boolean;
  panelId: string;
  onOpenFind: () => void;
  onTabChange: (tab: SpreadsheetRibbonTabId) => void;
  onTogglePanel: (
    panel: SpreadsheetWorkbookPanelView,
    trigger: HTMLButtonElement,
  ) => void;
  panel: SpreadsheetWorkbookPanelView | null;
  toolbarCell: Cell | null | undefined;
}) {
  const formulaCount = useMemo(
    () => spreadsheetFormulaCount(content),
    [content],
  );
  const pivotCount = useMemo(() => spreadsheetPivotCount(content), [content]);
  const numberFormat = toolbarCell?.ct?.fa?.trim() || 'General';
  const numberFormatPreset = spreadsheetNumberFormatPreset(numberFormat);
  const decreasedNumberFormat = adjustSpreadsheetNumberFormat(numberFormat, -1);
  const increasedNumberFormat = adjustSpreadsheetNumberFormat(numberFormat, 1);
  const currentNumberFormatValue = spreadsheetNumberFormatValue(
    numberFormat,
    toolbarCell,
  );
  const decreasedNumberFormatValue = spreadsheetNumberFormatValue(
    decreasedNumberFormat,
    toolbarCell,
  );
  const increasedNumberFormatValue = spreadsheetNumberFormatValue(
    increasedNumberFormat,
    toolbarCell,
  );

  return (
    <WorkOfficeRibbon
      ariaLabel="表格功能区"
      tabs={spreadsheetRibbonTabs}
      defaultTab="home"
      activeTab={activeTab}
      onTabChange={onTabChange}
      adaptive
      collapsible
      fileActions={fileActions}
      quickAccessActions={[
        {
          id: spreadsheetCommandCatalog.undo.id,
          label: spreadsheetCommandCatalog.undo.label,
          icon: <Undo2 size={15} />,
          shortcut: spreadsheetCommandCatalog.undo.shortcut.label,
          ariaKeyShortcuts: spreadsheetCommandCatalog.undo.shortcut.aria,
          disabled: !can.undo(),
          onSelect: () => {
            commands.undo();
          },
        },
        {
          id: spreadsheetCommandCatalog.redo.id,
          label: spreadsheetCommandCatalog.redo.label,
          icon: <Redo2 size={15} />,
          shortcut: spreadsheetCommandCatalog.redo.shortcut.label,
          ariaKeyShortcuts: spreadsheetCommandCatalog.redo.shortcut.aria,
          disabled: !can.redo(),
          onSelect: () => {
            commands.redo();
          },
        },
      ]}
      className="work-spreadsheet-ribbon"
      toolbarClassName="work-spreadsheet-ribbon-toolbar"
      panels={{
        home: (
          <>
            <WorkOfficeRibbonGroup label="剪贴板" priority="high">
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.paste.label}
                title={`${spreadsheetCommandCatalog.paste.label}（${spreadsheetCommandCatalog.paste.shortcut.label}）`}
                aria-keyshortcuts={
                  spreadsheetCommandCatalog.paste.shortcut.aria
                }
                disabled={!can.pasteSelection()}
                onClick={commands.pasteSelection}
              >
                <ClipboardPaste size={19} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.cut.label}
                title={`${spreadsheetCommandCatalog.cut.label}（${spreadsheetCommandCatalog.cut.shortcut.label}）`}
                aria-keyshortcuts={spreadsheetCommandCatalog.cut.shortcut.aria}
                disabled={!can.cutSelection()}
                onClick={commands.cutSelection}
              >
                <Scissors size={19} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.copy.label}
                title={`${spreadsheetCommandCatalog.copy.label}（${spreadsheetCommandCatalog.copy.shortcut.label}）`}
                aria-keyshortcuts={spreadsheetCommandCatalog.copy.shortcut.aria}
                disabled={!can.copySelection()}
                onClick={commands.copySelection}
              >
                <Copy size={19} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.formatPainter.label}
                title={spreadsheetFormatPainterTitle(formatPainterMode)}
                active={formatPainterMode !== null}
                badge={formatPainterMode === 'locked' ? '连续' : undefined}
                disabled={
                  formatPainterMode === null
                    ? !can.activateFormatPainter('once')
                    : !can.cancelFormatPainter()
                }
                onClick={(event) => {
                  if (event.detail > 1) return;
                  if (formatPainterMode === null) {
                    commands.activateFormatPainter('once');
                  } else {
                    commands.cancelFormatPainter();
                  }
                }}
                onDoubleClick={() => commands.activateFormatPainter('locked')}
              >
                <Paintbrush size={19} />
              </WorkOfficeRibbonButton>
            </WorkOfficeRibbonGroup>
            <WorkOfficeRibbonGroup label="字体" priority="high">
              <OfficeSelect
                className="work-spreadsheet-font-family"
                ariaLabel="字体"
                value={
                  typeof toolbarCell?.ff === 'string' ? toolbarCell.ff : 'Aptos'
                }
                disabled={
                  !can.setCellFormat(
                    'ff',
                    typeof toolbarCell?.ff === 'string'
                      ? toolbarCell.ff
                      : 'Aptos',
                  )
                }
                options={spreadsheetFontFamilyOptions(
                  typeof toolbarCell?.ff === 'string'
                    ? toolbarCell.ff
                    : undefined,
                )}
                onValueChange={(value) => commands.setCellFormat('ff', value)}
              />
              <OfficeSelect
                ariaLabel="字号"
                value={String(toolbarCell?.fs ?? 10)}
                disabled={
                  !can.setCellFormat('fs', Number(toolbarCell?.fs ?? 10))
                }
                options={spreadsheetFontSizeOptions(toolbarCell?.fs)}
                onValueChange={(value) =>
                  commands.setCellFormat('fs', Number(value))
                }
              />
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.bold.label}
                title={`${spreadsheetCommandCatalog.bold.label}（${spreadsheetCommandCatalog.bold.shortcut.label}）`}
                aria-keyshortcuts={spreadsheetCommandCatalog.bold.shortcut.aria}
                displayLabel={false}
                active={Number(toolbarCell?.bl) === 1}
                disabled={
                  !can.setCellFormat(
                    'bl',
                    Number(toolbarCell?.bl) === 1 ? 0 : 1,
                  )
                }
                onClick={() =>
                  commands.setCellFormat(
                    'bl',
                    Number(toolbarCell?.bl) === 1 ? 0 : 1,
                  )
                }
              >
                <Bold size={15} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.italic.label}
                title={`${spreadsheetCommandCatalog.italic.label}（${spreadsheetCommandCatalog.italic.shortcut.label}）`}
                aria-keyshortcuts={
                  spreadsheetCommandCatalog.italic.shortcut.aria
                }
                displayLabel={false}
                active={Number(toolbarCell?.it) === 1}
                disabled={
                  !can.setCellFormat(
                    'it',
                    Number(toolbarCell?.it) === 1 ? 0 : 1,
                  )
                }
                onClick={() =>
                  commands.setCellFormat(
                    'it',
                    Number(toolbarCell?.it) === 1 ? 0 : 1,
                  )
                }
              >
                <Italic size={15} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.underline.label}
                title={`${spreadsheetCommandCatalog.underline.label}（${spreadsheetCommandCatalog.underline.shortcut.label}）`}
                aria-keyshortcuts={
                  spreadsheetCommandCatalog.underline.shortcut.aria
                }
                displayLabel={false}
                active={Number(toolbarCell?.un) === 1}
                disabled={
                  !can.setCellFormat(
                    'un',
                    Number(toolbarCell?.un) === 1 ? 0 : 1,
                  )
                }
                onClick={() =>
                  commands.setCellFormat(
                    'un',
                    Number(toolbarCell?.un) === 1 ? 0 : 1,
                  )
                }
              >
                <Underline size={15} />
              </WorkOfficeRibbonButton>
              <OfficeColorPicker
                compact
                className="work-color-tool"
                ariaLabel="文字颜色"
                value={
                  typeof toolbarCell?.fc === 'string'
                    ? toolbarCell.fc
                    : '#172033'
                }
                disabled={
                  !can.setCellFormat(
                    'fc',
                    typeof toolbarCell?.fc === 'string'
                      ? toolbarCell.fc
                      : '#172033',
                  )
                }
                onValueChange={(value) => commands.setCellFormat('fc', value)}
              />
              <OfficeColorPicker
                compact
                className="work-color-tool work-spreadsheet-fill-color"
                ariaLabel="填充颜色"
                value={
                  typeof toolbarCell?.bg === 'string'
                    ? toolbarCell.bg
                    : '#ffffff'
                }
                disabled={
                  !can.setCellFormat(
                    'bg',
                    typeof toolbarCell?.bg === 'string'
                      ? toolbarCell.bg
                      : '#ffffff',
                  )
                }
                onValueChange={(value) => commands.setCellFormat('bg', value)}
              />
            </WorkOfficeRibbonGroup>
            <WorkOfficeRibbonGroup label="对齐" priority="high">
              <WorkOfficeRibbonButton
                label="左对齐"
                displayLabel={false}
                active={String(toolbarCell?.ht ?? '1') === '1'}
                disabled={!can.setCellFormat('ht', '1')}
                onClick={() => commands.setCellFormat('ht', '1')}
              >
                <AlignLeft size={15} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label="居中"
                displayLabel={false}
                active={String(toolbarCell?.ht) === '0'}
                disabled={!can.setCellFormat('ht', '0')}
                onClick={() => commands.setCellFormat('ht', '0')}
              >
                <AlignCenter size={15} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label="右对齐"
                displayLabel={false}
                active={String(toolbarCell?.ht) === '2'}
                disabled={!can.setCellFormat('ht', '2')}
                onClick={() => commands.setCellFormat('ht', '2')}
              >
                <AlignRight size={15} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label="顶端对齐"
                displayLabel={false}
                active={Number(toolbarCell?.vt ?? 1) === 1}
                disabled={!can.setCellFormat('vt', 1)}
                onClick={() => commands.setCellFormat('vt', 1)}
              >
                <AlignVerticalJustifyStart size={15} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label="垂直居中"
                displayLabel={false}
                active={Number(toolbarCell?.vt) === 0}
                disabled={!can.setCellFormat('vt', 0)}
                onClick={() => commands.setCellFormat('vt', 0)}
              >
                <AlignVerticalJustifyCenter size={15} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label="底端对齐"
                displayLabel={false}
                active={Number(toolbarCell?.vt) === 2}
                disabled={!can.setCellFormat('vt', 2)}
                onClick={() => commands.setCellFormat('vt', 2)}
              >
                <AlignVerticalJustifyEnd size={15} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label="自动换行"
                displayLabel={false}
                active={String(toolbarCell?.tb) === '2'}
                disabled={
                  !can.setCellFormat(
                    'tb',
                    String(toolbarCell?.tb) === '2' ? '1' : '2',
                  )
                }
                onClick={() =>
                  commands.setCellFormat(
                    'tb',
                    String(toolbarCell?.tb) === '2' ? '1' : '2',
                  )
                }
              >
                <WrapText size={15} />
              </WorkOfficeRibbonButton>
              <SpreadsheetMergeMenu can={can} commands={commands} />
            </WorkOfficeRibbonGroup>
            <WorkOfficeRibbonGroup label="数字" priority="high">
              <OfficeSelect
                className="work-spreadsheet-number-format"
                ariaLabel="数字格式"
                value={numberFormatPreset}
                disabled={!can.setCellFormat('ct', currentNumberFormatValue)}
                options={spreadsheetNumberFormatOptions}
                onValueChange={(preset) => {
                  if (preset === 'custom') return;
                  commands.setCellFormat(
                    'ct',
                    spreadsheetNumberFormatValue(
                      spreadsheetNumberFormatCode(preset),
                      toolbarCell,
                    ),
                  );
                }}
              />
              <WorkOfficeRibbonButton
                label="百分比格式"
                title="百分比格式"
                displayLabel={false}
                active={numberFormatPreset === 'percent'}
                disabled={
                  !can.setCellFormat(
                    'ct',
                    spreadsheetNumberFormatValue(
                      spreadsheetNumberFormatCode('percent'),
                      toolbarCell,
                    ),
                  )
                }
                onClick={() =>
                  commands.setCellFormat(
                    'ct',
                    spreadsheetNumberFormatValue(
                      spreadsheetNumberFormatCode('percent'),
                      toolbarCell,
                    ),
                  )
                }
              >
                <Percent size={15} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label="减少小数位"
                title="减少小数位"
                displayLabel={false}
                disabled={
                  decreasedNumberFormat === numberFormat ||
                  !can.setCellFormat('ct', decreasedNumberFormatValue)
                }
                onClick={() =>
                  commands.setCellFormat('ct', decreasedNumberFormatValue)
                }
              >
                <DecimalsArrowLeft size={16} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label="增加小数位"
                title="增加小数位"
                displayLabel={false}
                disabled={
                  increasedNumberFormat === numberFormat ||
                  !can.setCellFormat('ct', increasedNumberFormatValue)
                }
                onClick={() =>
                  commands.setCellFormat('ct', increasedNumberFormatValue)
                }
              >
                <DecimalsArrowRight size={16} />
              </WorkOfficeRibbonButton>
            </WorkOfficeRibbonGroup>
            <WorkOfficeRibbonGroup label="样式">
              <SpreadsheetRibbonTool
                controlsId={panelId}
                panel="conditional-formatting"
                label={spreadsheetCommandCatalog.conditionalFormatting.label}
                count={managedConditionalFormatCount(content)}
                icon={<Palette size={19} />}
                active={panel === 'conditional-formatting'}
                onToggle={onTogglePanel}
              />
            </WorkOfficeRibbonGroup>
            <WorkOfficeRibbonGroup label="单元格">
              <SpreadsheetRowsAndColumnsMenu can={can} commands={commands} />
            </WorkOfficeRibbonGroup>
            <WorkOfficeRibbonGroup label="编辑" priority="low">
              <SpreadsheetClearMenu can={can} commands={commands} />
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.find.label}
                title={`${spreadsheetCommandCatalog.find.label}（${spreadsheetCommandCatalog.find.shortcut.label}）`}
                aria-keyshortcuts={spreadsheetCommandCatalog.find.shortcut.aria}
                active={findOpen}
                onClick={onOpenFind}
              >
                <Search size={19} />
              </WorkOfficeRibbonButton>
            </WorkOfficeRibbonGroup>
          </>
        ),
        insert: (
          <WorkOfficeRibbonGroup label="图表" priority="high">
            <SpreadsheetRibbonTool
              controlsId={panelId}
              panel="charts"
              label={spreadsheetCommandCatalog.insertChart.label}
              count={spreadsheetChartCount(content)}
              icon={<BarChart3 size={19} />}
              active={panel === 'charts'}
              onToggle={onTogglePanel}
            />
          </WorkOfficeRibbonGroup>
        ),
        pageLayout: (
          <WorkOfficeRibbonGroup label="页面设置" priority="high">
            <SpreadsheetRibbonTool
              controlsId={panelId}
              panel="print-area"
              label={spreadsheetCommandCatalog.printSettings.label}
              count={spreadsheetPrintSettingCount(content)}
              icon={<Printer size={19} />}
              active={panel === 'print-area'}
              onToggle={onTogglePanel}
            />
          </WorkOfficeRibbonGroup>
        ),
        formulas: (
          <>
            <WorkOfficeRibbonGroup label="定义的名称" priority="high">
              <SpreadsheetRibbonTool
                controlsId={panelId}
                panel="names"
                label={spreadsheetCommandCatalog.nameManager.label}
                count={content.namedRanges?.length ?? 0}
                icon={<Bookmark size={19} />}
                active={panel === 'names'}
                onToggle={onTogglePanel}
              />
            </WorkOfficeRibbonGroup>
            <WorkOfficeRibbonGroup label="计算" priority="high">
              <SpreadsheetRibbonTool
                controlsId={panelId}
                panel="formulas"
                label={spreadsheetCommandCatalog.formulaManager.label}
                count={formulaCount}
                icon={<Calculator size={19} />}
                active={panel === 'formulas'}
                onToggle={onTogglePanel}
              />
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.recalculateWorkbook.label}
                title={`${spreadsheetCommandCatalog.recalculateWorkbook.label}（${spreadsheetCommandCatalog.recalculateWorkbook.shortcut.label}）`}
                aria-keyshortcuts={
                  spreadsheetCommandCatalog.recalculateWorkbook.shortcut.aria
                }
                disabled={!can.recalculateFormula('workbook')}
                onClick={() => commands.recalculateFormula('workbook')}
              >
                <RefreshCw size={19} />
              </WorkOfficeRibbonButton>
            </WorkOfficeRibbonGroup>
          </>
        ),
        data: (
          <>
            <WorkOfficeRibbonGroup label="排序和筛选" priority="high">
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.sortAscending.label}
                disabled={!can.sortSelectedCells('ascending')}
                onClick={() => commands.sortSelectedCells('ascending')}
              >
                <SortAsc size={19} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.sortDescending.label}
                disabled={!can.sortSelectedCells('descending')}
                onClick={() => commands.sortSelectedCells('descending')}
              >
                <SortDesc size={19} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.autoFilter.label}
                title={`${spreadsheetCommandCatalog.autoFilter.label}（${spreadsheetCommandCatalog.autoFilter.shortcut.label}）；表头菜单（${spreadsheetCommandCatalog.autoFilter.menuShortcut.label}）`}
                aria-keyshortcuts={
                  spreadsheetCommandCatalog.autoFilter.shortcut.aria
                }
                active={autoFilterActive}
                disabled={!can.toggleAutoFilter()}
                onClick={() => commands.toggleAutoFilter()}
              >
                <ListFilter size={19} />
              </WorkOfficeRibbonButton>
            </WorkOfficeRibbonGroup>
            <WorkOfficeRibbonGroup label="分析">
              <SpreadsheetRibbonTool
                controlsId={panelId}
                panel="pivots"
                label={spreadsheetCommandCatalog.pivotTable.label}
                count={pivotCount}
                icon={<TableProperties size={19} />}
                active={panel === 'pivots'}
                onToggle={onTogglePanel}
              />
            </WorkOfficeRibbonGroup>
          </>
        ),
        review: (
          <WorkOfficeRibbonGroup label="保护" priority="high">
            <SpreadsheetRibbonTool
              controlsId={panelId}
              panel="protection"
              label={spreadsheetCommandCatalog.protectSheet.label}
              count={protectedSheetCount(content.sheets)}
              icon={<ShieldCheck size={19} />}
              active={panel === 'protection'}
              onToggle={onTogglePanel}
            />
          </WorkOfficeRibbonGroup>
        ),
        view: (
          <>
            <WorkOfficeRibbonGroup label="工作簿视图" priority="high">
              <WorkOfficeRibbonButton
                label={gridLinesVisible ? '隐藏网格线' : '显示网格线'}
                visibleLabel={spreadsheetCommandCatalog.gridLines.label}
                active={gridLinesVisible}
                disabled={!can.setGridLines(!gridLinesVisible)}
                onClick={() => commands.setGridLines(!gridLinesVisible)}
              >
                <Grid3X3 size={19} />
              </WorkOfficeRibbonButton>
            </WorkOfficeRibbonGroup>
            <WorkOfficeRibbonGroup label="窗口" priority="high">
              <SpreadsheetFreezePanesMenu
                active={freezePanesActive}
                can={can}
                commands={commands}
                selection={freezePanesSelection}
              />
            </WorkOfficeRibbonGroup>
          </>
        ),
      }}
    />
  );
}

function SpreadsheetMergeMenu({
  can,
  commands,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
}) {
  const items: readonly {
    command: SpreadsheetCellMergeCommand;
    id: string;
    label: string;
    icon: ReactNode;
  }[] = [
    {
      command: 'merge-and-center',
      id: spreadsheetCommandCatalog.mergeAndCenter.id,
      label: spreadsheetCommandCatalog.mergeAndCenter.label,
      icon: <Merge size={16} />,
    },
    {
      command: 'merge-cells',
      id: spreadsheetCommandCatalog.mergeCells.id,
      label: spreadsheetCommandCatalog.mergeCells.label,
      icon: <Grid3X3 size={16} />,
    },
    {
      command: 'merge-across',
      id: spreadsheetCommandCatalog.mergeAcross.id,
      label: spreadsheetCommandCatalog.mergeAcross.label,
      icon: <Rows3 size={16} />,
    },
    {
      command: 'unmerge-cells',
      id: spreadsheetCommandCatalog.unmergeCells.id,
      label: spreadsheetCommandCatalog.unmergeCells.label,
      icon: <X size={16} />,
    },
    {
      command: 'unmerge-and-fill',
      id: spreadsheetCommandCatalog.unmergeAndFill.id,
      label: spreadsheetCommandCatalog.unmergeAndFill.label,
      icon: <Grid3X3 size={16} />,
    },
  ];
  const primaryDisabled = !can.mergeSelectedCells('merge-and-center');
  const menuDisabled = items.every(
    ({ command }) => can.mergeSelectedCells(command) === false,
  );

  return (
    <Popover
      label="更多合并方式"
      panelLabel="合并选项"
      panelRole="menu"
      portal
      className="work-spreadsheet-ribbon-split-root"
      panelClassName="work-office-context-menu work-spreadsheet-ribbon-menu"
      disabled={menuDisabled}
      focusFirstOnOpen
      onPanelKeyDown={moveOfficeMenuFocus}
      trigger={(triggerProps, { open }) => (
        <>
          <button
            type="button"
            className="with-label work-spreadsheet-ribbon-split-primary"
            aria-label={spreadsheetCommandCatalog.mergeAndCenter.label}
            aria-keyshortcuts={
              spreadsheetCommandCatalog.mergeAndCenter.shortcut.aria
            }
            title={`${spreadsheetCommandCatalog.mergeAndCenter.label}（${spreadsheetCommandCatalog.mergeAndCenter.shortcut.label}）`}
            disabled={primaryDisabled}
            onClick={() => commands.mergeSelectedCells('merge-and-center')}
          >
            <Merge size={19} />
            <span>{spreadsheetCommandCatalog.mergeAndCenter.label}</span>
          </button>
          <button
            {...triggerProps}
            className={`work-spreadsheet-ribbon-split-disclosure${open ? ' active' : ''}`}
            title="更多合并方式"
          >
            <ChevronDown size={13} aria-hidden="true" />
          </button>
        </>
      )}
    >
      {(close) =>
        items.map(({ command, id, label, icon }) => (
          <button
            key={id}
            type="button"
            role="menuitem"
            tabIndex={-1}
            aria-keyshortcuts={
              command === 'merge-and-center'
                ? spreadsheetCommandCatalog.mergeAndCenter.shortcut.aria
                : undefined
            }
            disabled={!can.mergeSelectedCells(command)}
            onClick={() => {
              close();
              commands.mergeSelectedCells(command);
            }}
          >
            <span className="work-spreadsheet-ribbon-menu-item-icon">
              {icon}
            </span>
            <span>{label}</span>
          </button>
        ))
      }
    </Popover>
  );
}

function SpreadsheetFreezePanesMenu({
  active,
  can,
  commands,
  selection,
}: {
  active: boolean;
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
  selection: Selection;
}) {
  const presets: readonly {
    preset: SpreadsheetFreezePanePreset;
    label: string;
    icon: ReactNode;
  }[] = [
    ...(active
      ? [
          {
            preset: 'none' as const,
            label: '取消冻结窗格',
            icon: <X size={16} />,
          },
        ]
      : []),
    {
      preset: 'selection',
      label: spreadsheetFreezePanesSelectionLabel(selection),
      icon: <PanelsTopLeft size={16} />,
    },
    {
      preset: 'topRow',
      label: '冻结首行',
      icon: <PanelTop size={16} />,
    },
    {
      preset: 'firstColumn',
      label: '冻结首列',
      icon: <PanelLeft size={16} />,
    },
  ];
  const disabled = !presets.some(({ preset }) => can.setFreezePanes(preset));

  return (
    <Popover
      label={spreadsheetCommandCatalog.freezePanes.label}
      panelLabel="冻结窗格选项"
      panelRole="menu"
      portal
      className="work-spreadsheet-ribbon-menu-root"
      panelClassName="work-office-context-menu work-spreadsheet-ribbon-menu"
      disabled={disabled}
      focusFirstOnOpen
      onPanelKeyDown={moveOfficeMenuFocus}
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          className={`with-label work-spreadsheet-ribbon-menu-trigger${active || open ? ' active' : ''}`}
          aria-pressed={active}
          title={active ? '冻结窗格（已启用）' : '冻结窗格'}
        >
          <PanelsTopLeft size={19} />
          <span>{spreadsheetCommandCatalog.freezePanes.label}</span>
        </button>
      )}
    >
      {(close) =>
        presets.map(({ preset, label, icon }) => (
          <button
            key={preset}
            type="button"
            role="menuitem"
            tabIndex={-1}
            disabled={!can.setFreezePanes(preset)}
            onClick={() => {
              close();
              commands.setFreezePanes(preset);
            }}
          >
            <span className="work-spreadsheet-ribbon-menu-item-icon">
              {icon}
            </span>
            <span>{label}</span>
          </button>
        ))
      }
    </Popover>
  );
}

function SpreadsheetRowsAndColumnsMenu({
  can,
  commands,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
}) {
  const items: readonly {
    id: string;
    label: string;
    icon: ReactNode;
    danger?: boolean;
    disabled: boolean;
    execute: () => boolean;
  }[] = [
    {
      id: spreadsheetCommandCatalog.insertRowsAbove.id,
      label: spreadsheetCommandCatalog.insertRowsAbove.label,
      icon: <BetweenHorizontalStart size={16} />,
      disabled: !can.insertSelectedStructure('row', 'before'),
      execute: () => commands.insertSelectedStructure('row', 'before'),
    },
    {
      id: spreadsheetCommandCatalog.insertRowsBelow.id,
      label: spreadsheetCommandCatalog.insertRowsBelow.label,
      icon: <BetweenHorizontalEnd size={16} />,
      disabled: !can.insertSelectedStructure('row', 'after'),
      execute: () => commands.insertSelectedStructure('row', 'after'),
    },
    {
      id: spreadsheetCommandCatalog.insertColumnsLeft.id,
      label: spreadsheetCommandCatalog.insertColumnsLeft.label,
      icon: <BetweenVerticalStart size={16} />,
      disabled: !can.insertSelectedStructure('column', 'before'),
      execute: () => commands.insertSelectedStructure('column', 'before'),
    },
    {
      id: spreadsheetCommandCatalog.insertColumnsRight.id,
      label: spreadsheetCommandCatalog.insertColumnsRight.label,
      icon: <BetweenVerticalEnd size={16} />,
      disabled: !can.insertSelectedStructure('column', 'after'),
      execute: () => commands.insertSelectedStructure('column', 'after'),
    },
    {
      id: spreadsheetCommandCatalog.deleteRows.id,
      label: spreadsheetCommandCatalog.deleteRows.label,
      icon: <Rows3 size={16} />,
      danger: true,
      disabled: !can.deleteSelectedStructure('row'),
      execute: () => commands.deleteSelectedStructure('row'),
    },
    {
      id: spreadsheetCommandCatalog.deleteColumns.id,
      label: spreadsheetCommandCatalog.deleteColumns.label,
      icon: <Columns3 size={16} />,
      danger: true,
      disabled: !can.deleteSelectedStructure('column'),
      execute: () => commands.deleteSelectedStructure('column'),
    },
  ];
  const disabled = items.every((item) => item.disabled);

  return (
    <Popover
      label="行和列"
      panelLabel="行和列选项"
      panelRole="menu"
      portal
      className="work-spreadsheet-ribbon-menu-root"
      panelClassName="work-office-context-menu work-spreadsheet-ribbon-menu"
      disabled={disabled}
      focusFirstOnOpen
      onPanelKeyDown={moveOfficeMenuFocus}
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          className={`with-label work-spreadsheet-ribbon-menu-trigger${open ? ' active' : ''}`}
          title="行和列"
        >
          <Rows3 size={19} />
          <span>行和列</span>
        </button>
      )}
    >
      {(close) =>
        items.map(({ id, label, icon, danger, disabled, execute }) => (
          <button
            key={id}
            type="button"
            role="menuitem"
            tabIndex={-1}
            className={danger ? 'danger' : undefined}
            disabled={disabled}
            onClick={() => {
              close();
              execute();
            }}
          >
            <span className="work-spreadsheet-ribbon-menu-item-icon">
              {icon}
            </span>
            <span>{label}</span>
          </button>
        ))
      }
    </Popover>
  );
}

function SpreadsheetClearMenu({
  can,
  commands,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
}) {
  const items: readonly {
    mode: SpreadsheetCellClearMode;
    id: string;
    label: string;
    icon: ReactNode;
  }[] = [
    {
      mode: 'all',
      id: spreadsheetCommandCatalog.clearAll.id,
      label: spreadsheetCommandCatalog.clearAll.label,
      icon: <Trash2 size={16} />,
    },
    {
      mode: 'formats',
      id: spreadsheetCommandCatalog.clearFormats.id,
      label: spreadsheetCommandCatalog.clearFormats.label,
      icon: <Paintbrush size={16} />,
    },
    {
      mode: 'contents',
      id: spreadsheetCommandCatalog.clearContents.id,
      label: spreadsheetCommandCatalog.clearContents.label,
      icon: <FileX2 size={16} />,
    },
    {
      mode: 'comments',
      id: spreadsheetCommandCatalog.clearComments.id,
      label: spreadsheetCommandCatalog.clearComments.label,
      icon: <MessageSquareX size={16} />,
    },
    {
      mode: 'hyperlinks',
      id: spreadsheetCommandCatalog.clearHyperlinks.id,
      label: spreadsheetCommandCatalog.clearHyperlinks.label,
      icon: <Link2Off size={16} />,
    },
  ];
  const disabled = items.every(({ mode }) => !can.clearSelectedCells(mode));

  return (
    <Popover
      label="清除"
      panelLabel="清除选项"
      panelRole="menu"
      portal
      className="work-spreadsheet-ribbon-menu-root"
      panelClassName="work-office-context-menu work-spreadsheet-ribbon-menu"
      disabled={disabled}
      focusFirstOnOpen
      onPanelKeyDown={moveOfficeMenuFocus}
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          className={`with-label work-spreadsheet-ribbon-menu-trigger${open ? ' active' : ''}`}
          title="清除"
        >
          <Eraser size={19} />
          <span>清除</span>
        </button>
      )}
    >
      {(close) =>
        items.map(({ mode, id, label, icon }) => (
          <button
            key={id}
            type="button"
            role="menuitem"
            tabIndex={-1}
            aria-keyshortcuts={
              mode === 'contents'
                ? spreadsheetCommandCatalog.clearContents.shortcut.aria
                : undefined
            }
            disabled={!can.clearSelectedCells(mode)}
            onClick={() => {
              close();
              commands.clearSelectedCells(mode);
            }}
          >
            <span className="work-spreadsheet-ribbon-menu-item-icon">
              {icon}
            </span>
            <span>{label}</span>
          </button>
        ))
      }
    </Popover>
  );
}

function spreadsheetFormatPainterTitle(
  mode: SpreadsheetFormatPainterMode | null,
): string {
  if (mode === 'locked') {
    return '格式刷已锁定（再次点击或按 Escape 退出）';
  }
  if (mode === 'once') {
    return '格式刷已开启（选择目标区域，按 Escape 退出）';
  }
  return '格式刷（单击应用一次，双击锁定连续应用）';
}

function SpreadsheetRibbonTool({
  active,
  count,
  controlsId,
  icon,
  label,
  panel,
  onToggle,
}: {
  active: boolean;
  count: number;
  controlsId: string;
  icon: ReactNode;
  label: string;
  panel: SpreadsheetWorkbookPanelView;
  onToggle: (
    panel: SpreadsheetWorkbookPanelView,
    trigger: HTMLButtonElement,
  ) => void;
}) {
  const hasCount = count > 0;
  return (
    <WorkOfficeRibbonButton
      label={hasCount ? `${label}（${count}）` : label}
      visibleLabel={label}
      badge={hasCount ? count : undefined}
      active={active}
      aria-controls={controlsId}
      aria-expanded={active}
      onClick={(event) => onToggle(panel, event.currentTarget)}
    >
      {icon}
    </WorkOfficeRibbonButton>
  );
}
