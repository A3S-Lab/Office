import type { Cell } from '@fortune-sheet/core';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  BarChart3,
  Bold,
  Bookmark,
  Calculator,
  DecimalsArrowLeft,
  DecimalsArrowRight,
  Grid3X3,
  Italic,
  Merge,
  Palette,
  Percent,
  Printer,
  Redo2,
  ShieldCheck,
  TableProperties,
  Underline,
  Undo2,
  WrapText,
} from 'lucide-react';
import { useMemo, type ReactNode } from 'react';
import { spreadsheetChartCount } from '../work-spreadsheet-charts';
import { spreadsheetFormulaCount } from '../work-spreadsheet-formula-analysis';
import { spreadsheetPivotCount } from '../work-spreadsheet-pivots';
import { protectedSheetCount } from '../work-spreadsheet-protection';
import type { WorkSpreadsheetContent } from '../work-types';
import { OfficeColorPicker, OfficeSelect } from './office-controls';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { managedConditionalFormatCount } from './spreadsheet-conditional-format-panel';
import {
  spreadsheetFontFamilyOptions,
  spreadsheetFontSizeOptions,
} from './spreadsheet-editor-support';
import {
  adjustSpreadsheetNumberFormat,
  spreadsheetNumberFormatCode,
  spreadsheetNumberFormatPreset,
  spreadsheetNumberFormatValue,
  type SpreadsheetNumberFormatPreset,
} from './spreadsheet-number-format';
import { spreadsheetPrintSettingCount } from './spreadsheet-print-settings-panel';
import type { SpreadsheetWorkbookPanelView } from './spreadsheet-workbook-panel';
import {
  type WorkOfficeFileAction,
  WorkOfficeRibbon,
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

const spreadsheetRibbonTabs = [
  { id: 'home', label: '开始' },
  { id: 'insert', label: '插入' },
  { id: 'pageLayout', label: '页面布局' },
  { id: 'formulas', label: '公式' },
  { id: 'data', label: '数据' },
  { id: 'review', label: '审阅' },
  { id: 'view', label: '视图' },
] as const;

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

export type SpreadsheetRibbonTabId =
  (typeof spreadsheetRibbonTabs)[number]['id'];

export function SpreadsheetEditorRibbon({
  activeTab,
  can,
  commands,
  content,
  fileActions,
  gridLinesVisible,
  multipleCellsSelected,
  onTabChange,
  onTogglePanel,
  panel,
  toolbarCell,
}: {
  activeTab: SpreadsheetRibbonTabId;
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
  content: WorkSpreadsheetContent;
  fileActions?: readonly WorkOfficeFileAction[];
  gridLinesVisible: boolean;
  multipleCellsSelected: boolean;
  onTabChange: (tab: SpreadsheetRibbonTabId) => void;
  onTogglePanel: (panel: SpreadsheetWorkbookPanelView) => void;
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
      fileActions={fileActions}
      className="work-spreadsheet-ribbon"
      toolbarClassName="work-spreadsheet-ribbon-toolbar"
      panels={{
        home: (
          <>
            <WorkOfficeRibbonGroup label="撤销与恢复">
              <WorkOfficeRibbonButton
                label="撤销"
                title="撤销（Cmd/Ctrl+Z）"
                aria-keyshortcuts="Control+Z Meta+Z"
                disabled={!can.undo()}
                onClick={commands.undo}
              >
                <Undo2 size={19} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label="重做"
                title="重做（Cmd/Ctrl+Shift+Z）"
                aria-keyshortcuts="Control+Shift+Z Meta+Shift+Z Control+Y Meta+Y"
                disabled={!can.redo()}
                onClick={commands.redo}
              >
                <Redo2 size={19} />
              </WorkOfficeRibbonButton>
            </WorkOfficeRibbonGroup>
            <WorkOfficeRibbonGroup label="字体">
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
                label="加粗"
                title="加粗（Cmd/Ctrl+B）"
                aria-keyshortcuts="Control+B Meta+B"
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
                label="斜体"
                title="斜体（Cmd/Ctrl+I）"
                aria-keyshortcuts="Control+I Meta+I"
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
                label="下划线"
                title="下划线（Cmd/Ctrl+U）"
                aria-keyshortcuts="Control+U Meta+U"
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
            <WorkOfficeRibbonGroup label="对齐">
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
            </WorkOfficeRibbonGroup>
            <WorkOfficeRibbonGroup label="数字">
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
            <WorkOfficeRibbonGroup label="单元格">
              <WorkOfficeRibbonButton
                label={toolbarCell?.mc ? '取消合并' : '合并单元格'}
                disabled={
                  !can.toggleCellMerge(Boolean(toolbarCell?.mc)) ||
                  (!toolbarCell?.mc && !multipleCellsSelected)
                }
                onClick={() =>
                  commands.toggleCellMerge(Boolean(toolbarCell?.mc))
                }
              >
                <Merge size={19} />
              </WorkOfficeRibbonButton>
            </WorkOfficeRibbonGroup>
          </>
        ),
        insert: (
          <>
            <WorkOfficeRibbonGroup label="图表">
              <SpreadsheetRibbonTool
                label="插入图表"
                count={spreadsheetChartCount(content)}
                icon={<BarChart3 size={19} />}
                active={panel === 'charts'}
                onClick={() => onTogglePanel('charts')}
              />
            </WorkOfficeRibbonGroup>
            <WorkOfficeRibbonGroup label="样式">
              <SpreadsheetRibbonTool
                label="条件格式"
                count={managedConditionalFormatCount(content)}
                icon={<Palette size={19} />}
                active={panel === 'conditional-formatting'}
                onClick={() => onTogglePanel('conditional-formatting')}
              />
            </WorkOfficeRibbonGroup>
          </>
        ),
        pageLayout: (
          <WorkOfficeRibbonGroup label="页面设置">
            <SpreadsheetRibbonTool
              label="打印设置"
              count={spreadsheetPrintSettingCount(content)}
              icon={<Printer size={19} />}
              active={panel === 'print-area'}
              onClick={() => onTogglePanel('print-area')}
            />
          </WorkOfficeRibbonGroup>
        ),
        formulas: (
          <>
            <WorkOfficeRibbonGroup label="定义的名称">
              <SpreadsheetRibbonTool
                label="名称管理器"
                count={content.namedRanges?.length ?? 0}
                icon={<Bookmark size={19} />}
                active={panel === 'names'}
                onClick={() => onTogglePanel('names')}
              />
            </WorkOfficeRibbonGroup>
            <WorkOfficeRibbonGroup label="计算">
              <SpreadsheetRibbonTool
                label="公式与计算"
                count={formulaCount}
                icon={<Calculator size={19} />}
                active={panel === 'formulas'}
                onClick={() => onTogglePanel('formulas')}
              />
            </WorkOfficeRibbonGroup>
          </>
        ),
        data: (
          <WorkOfficeRibbonGroup label="分析">
            <SpreadsheetRibbonTool
              label="数据透视表"
              count={pivotCount}
              icon={<TableProperties size={19} />}
              active={panel === 'pivots'}
              onClick={() => onTogglePanel('pivots')}
            />
          </WorkOfficeRibbonGroup>
        ),
        review: (
          <WorkOfficeRibbonGroup label="保护">
            <SpreadsheetRibbonTool
              label="工作表保护"
              count={protectedSheetCount(content.sheets)}
              icon={<ShieldCheck size={19} />}
              active={panel === 'protection'}
              onClick={() => onTogglePanel('protection')}
            />
          </WorkOfficeRibbonGroup>
        ),
        view: (
          <WorkOfficeRibbonGroup label="工作簿视图">
            <WorkOfficeRibbonButton
              label={gridLinesVisible ? '隐藏网格线' : '显示网格线'}
              visibleLabel="网格线"
              active={gridLinesVisible}
              disabled={!can.setGridLines(!gridLinesVisible)}
              onClick={() => commands.setGridLines(!gridLinesVisible)}
            >
              <Grid3X3 size={19} />
            </WorkOfficeRibbonButton>
          </WorkOfficeRibbonGroup>
        ),
      }}
    />
  );
}

function SpreadsheetRibbonTool({
  active,
  count,
  icon,
  label,
  onClick,
}: {
  active: boolean;
  count: number;
  icon: ReactNode;
  label: string;
  onClick: () => void;
}) {
  const hasCount = count > 0;
  return (
    <WorkOfficeRibbonButton
      label={hasCount ? `${label}（${count}）` : label}
      visibleLabel={label}
      badge={hasCount ? count : undefined}
      active={active}
      onClick={onClick}
    >
      {icon}
    </WorkOfficeRibbonButton>
  );
}
