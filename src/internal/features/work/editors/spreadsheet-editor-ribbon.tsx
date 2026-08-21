import type { Cell, Selection } from '@fortune-sheet/core';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  ArrowDown,
  ArrowDownToLine,
  ArrowLeftToLine,
  ArrowRightToLine,
  ArrowUp,
  ArrowUpToLine,
  BarChart3,
  BetweenHorizontalEnd,
  BetweenHorizontalStart,
  BetweenVerticalEnd,
  BetweenVerticalStart,
  Bookmark,
  Calculator,
  ChevronDown,
  ClipboardPaste,
  Copy,
  Columns3,
  Eraser,
  FileX2,
  Grid3X3,
  Hash,
  Link2Off,
  Link2,
  ListFilter,
  ListChecks,
  LocateFixed,
  MessageSquareX,
  Merge,
  Paintbrush,
  Palette,
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
  Sigma,
  Rows3,
  TableProperties,
  Trash2,
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
import type { SpreadsheetCellMergeCommand } from './spreadsheet-cell-merge';
import type { SpreadsheetAutoSumFunction } from './spreadsheet-auto-sum';
import type { SpreadsheetResolvedCellBorders } from './spreadsheet-cell-border';
import type { SpreadsheetCellClearMode } from './spreadsheet-cell-clear';
import type { SpreadsheetCellFillDirection } from './spreadsheet-cell-fill';
import { SpreadsheetCellStyleRibbon } from './spreadsheet-cell-style-ribbon';
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
import type { SpreadsheetFormatPainterMode } from './spreadsheet-format-painter';
import {
  type SpreadsheetFreezePanePreset,
  spreadsheetFreezePanesSelectionLabel,
} from './spreadsheet-freeze-panes';
import {
  SpreadsheetFontRibbonGroup,
  SpreadsheetNumberRibbonGroup,
} from './spreadsheet-home-format-ribbon';
import { spreadsheetPrintSettingCount } from './spreadsheet-print-settings-panel';
import type { SpreadsheetPasteContent } from './spreadsheet-paste-special';
import type { SpreadsheetWorkbookPanelView } from './spreadsheet-workbook-panel';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import {
  type WorkOfficeFileAction,
  WorkOfficeRibbon,
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

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
  onTabChange,
  onTogglePanel,
  panel,
  toolbarCell,
  toolbarCellBorders,
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
  onTabChange: (tab: SpreadsheetRibbonTabId) => void;
  onTogglePanel: (
    panel: SpreadsheetWorkbookPanelView,
    trigger: HTMLButtonElement,
  ) => void;
  panel: SpreadsheetWorkbookPanelView | null;
  toolbarCell: Cell | null | undefined;
  toolbarCellBorders?: SpreadsheetResolvedCellBorders;
}) {
  const formulaCount = useMemo(
    () => spreadsheetFormulaCount(content),
    [content],
  );
  const pivotCount = useMemo(() => spreadsheetPivotCount(content), [content]);
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
              <SpreadsheetPasteMenu can={can} commands={commands} />
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
            <SpreadsheetFontRibbonGroup
              can={can}
              commands={commands}
              toolbarCell={toolbarCell}
            />
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
            <SpreadsheetNumberRibbonGroup
              can={can}
              commands={commands}
              toolbarCell={toolbarCell}
            />
            <WorkOfficeRibbonGroup label="样式">
              <SpreadsheetCellStyleRibbon
                can={can}
                commands={commands}
                toolbarCell={toolbarCell}
                toolbarCellBorders={toolbarCellBorders}
              />
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
              <SpreadsheetAutoSumMenu can={can} commands={commands} />
              <SpreadsheetFillMenu can={can} commands={commands} />
              <SpreadsheetClearMenu can={can} commands={commands} />
              <SpreadsheetFindAndSelectMenu
                can={can}
                commands={commands}
                findOpen={findOpen}
              />
            </WorkOfficeRibbonGroup>
          </>
        ),
        insert: (
          <>
            <WorkOfficeRibbonGroup label="链接" priority="high">
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.hyperlink.label}
                title={`${spreadsheetCommandCatalog.hyperlink.label}（${spreadsheetCommandCatalog.hyperlink.shortcut.label}）`}
                aria-keyshortcuts={
                  spreadsheetCommandCatalog.hyperlink.shortcut.aria
                }
                disabled={!can.openHyperlink()}
                onClick={commands.openHyperlink}
              >
                <Link2 size={19} />
              </WorkOfficeRibbonButton>
            </WorkOfficeRibbonGroup>
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
          </>
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
            <WorkOfficeRibbonGroup label="数据工具" priority="high">
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.dataValidation.label}
                disabled={!can.openDataValidation()}
                onClick={commands.openDataValidation}
              >
                <ListChecks size={19} />
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

function SpreadsheetFindAndSelectMenu({
  can,
  commands,
  findOpen,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
  findOpen: boolean;
}) {
  const findDisabled = !can.openFind();
  const goToDisabled = !can.openGoTo();
  return (
    <Popover
      label={spreadsheetCommandCatalog.findAndSelect.label}
      panelLabel="查找和选择选项"
      panelRole="menu"
      portal
      className="work-spreadsheet-ribbon-menu-root"
      panelClassName="work-office-context-menu work-spreadsheet-ribbon-menu work-spreadsheet-find-select-menu"
      disabled={findDisabled && goToDisabled}
      focusFirstOnOpen
      onPanelKeyDown={moveOfficeMenuFocus}
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          className={`with-label work-spreadsheet-ribbon-menu-trigger${findOpen || open ? ' active' : ''}`}
          aria-pressed={findOpen}
          title={spreadsheetCommandCatalog.findAndSelect.label}
        >
          <Search size={19} />
          <span>{spreadsheetCommandCatalog.findAndSelect.label}</span>
        </button>
      )}
    >
      {(close) => (
        <>
          <button
            type="button"
            role="menuitem"
            tabIndex={-1}
            aria-label={spreadsheetCommandCatalog.find.label}
            aria-keyshortcuts={spreadsheetCommandCatalog.find.shortcut.aria}
            disabled={findDisabled}
            onClick={() => {
              close();
              commands.openFind();
            }}
          >
            <span className="work-spreadsheet-ribbon-menu-item-icon">
              <Search size={16} />
            </span>
            <span>{spreadsheetCommandCatalog.find.label}</span>
            <kbd>{spreadsheetCommandCatalog.find.shortcut.label}</kbd>
          </button>
          <button
            type="button"
            role="menuitem"
            tabIndex={-1}
            aria-label={spreadsheetCommandCatalog.goTo.label}
            aria-keyshortcuts={spreadsheetCommandCatalog.goTo.shortcut.aria}
            disabled={goToDisabled}
            onClick={() => {
              close();
              commands.openGoTo();
            }}
          >
            <span className="work-spreadsheet-ribbon-menu-item-icon">
              <LocateFixed size={16} />
            </span>
            <span>{spreadsheetCommandCatalog.goTo.label}</span>
            <kbd>{spreadsheetCommandCatalog.goTo.shortcut.label}</kbd>
          </button>
        </>
      )}
    </Popover>
  );
}

function SpreadsheetPasteMenu({
  can,
  commands,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
}) {
  const items: readonly {
    content: SpreadsheetPasteContent;
    id: string;
    label: string;
    icon: ReactNode;
  }[] = [
    {
      content: 'all',
      id: `${spreadsheetCommandCatalog.paste.id}.all`,
      label: '全部',
      icon: <ClipboardPaste size={16} />,
    },
    {
      content: 'values',
      id: `${spreadsheetCommandCatalog.paste.id}.values`,
      label: '值',
      icon: <Hash size={16} />,
    },
    {
      content: 'formulas',
      id: `${spreadsheetCommandCatalog.paste.id}.formulas`,
      label: '公式',
      icon: <Sigma size={16} />,
    },
    {
      content: 'formats',
      id: `${spreadsheetCommandCatalog.paste.id}.formats`,
      label: '格式',
      icon: <Paintbrush size={16} />,
    },
  ];
  const primaryDisabled = !can.pasteSelection();
  const menuDisabled =
    !can.openPasteSpecial() &&
    items.every(({ content }) => !can.pasteSpecial(content));

  return (
    <Popover
      label="更多粘贴方式"
      panelLabel="粘贴选项"
      panelRole="menu"
      portal
      className="work-spreadsheet-ribbon-split-root"
      panelClassName="work-office-context-menu work-spreadsheet-ribbon-menu work-spreadsheet-paste-menu"
      disabled={menuDisabled}
      focusFirstOnOpen
      onPanelKeyDown={moveOfficeMenuFocus}
      trigger={(triggerProps, { open }) => (
        <>
          <button
            type="button"
            className="with-label work-spreadsheet-ribbon-split-primary"
            aria-label={spreadsheetCommandCatalog.paste.label}
            aria-keyshortcuts={spreadsheetCommandCatalog.paste.shortcut.aria}
            title={`${spreadsheetCommandCatalog.paste.label}（${spreadsheetCommandCatalog.paste.shortcut.label}）`}
            disabled={primaryDisabled}
            onClick={commands.pasteSelection}
          >
            <ClipboardPaste size={19} />
            <span>{spreadsheetCommandCatalog.paste.label}</span>
          </button>
          <button
            {...triggerProps}
            className={`work-spreadsheet-ribbon-split-disclosure${open ? ' active' : ''}`}
            title="更多粘贴方式"
          >
            <ChevronDown size={13} aria-hidden="true" />
          </button>
        </>
      )}
    >
      {(close) => (
        <>
          {items.map(({ content, id, label, icon }) => (
            <button
              key={id}
              type="button"
              role="menuitem"
              tabIndex={-1}
              disabled={!can.pasteSpecial(content)}
              onClick={() => {
                close();
                commands.pasteSpecial(content);
              }}
            >
              <span className="work-spreadsheet-ribbon-menu-item-icon">
                {icon}
              </span>
              <span>{label}</span>
            </button>
          ))}
          <hr className="work-spreadsheet-ribbon-menu-separator" />
          <button
            type="button"
            role="menuitem"
            tabIndex={-1}
            aria-keyshortcuts={
              spreadsheetCommandCatalog.pasteSpecial.shortcut.aria
            }
            disabled={!can.openPasteSpecial()}
            onClick={() => {
              close();
              commands.openPasteSpecial();
            }}
          >
            <span className="work-spreadsheet-ribbon-menu-item-icon">
              <TableProperties size={16} />
            </span>
            <span>{spreadsheetCommandCatalog.pasteSpecial.label}…</span>
            <kbd>{spreadsheetCommandCatalog.pasteSpecial.shortcut.label}</kbd>
          </button>
        </>
      )}
    </Popover>
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

function SpreadsheetAutoSumMenu({
  can,
  commands,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
}) {
  const items: readonly {
    functionName: SpreadsheetAutoSumFunction;
    id: string;
    label: string;
    icon: ReactNode;
    shortcut?: string;
  }[] = [
    {
      functionName: 'sum',
      id: spreadsheetCommandCatalog.autoSum.id,
      label: spreadsheetCommandCatalog.autoSum.label,
      icon: <Sigma size={16} />,
      shortcut: spreadsheetCommandCatalog.autoSum.shortcut.aria,
    },
    {
      functionName: 'average',
      id: spreadsheetCommandCatalog.autoAverage.id,
      label: spreadsheetCommandCatalog.autoAverage.label,
      icon: <Calculator size={16} />,
    },
    {
      functionName: 'count',
      id: spreadsheetCommandCatalog.autoCount.id,
      label: spreadsheetCommandCatalog.autoCount.label,
      icon: <Hash size={16} />,
    },
    {
      functionName: 'max',
      id: spreadsheetCommandCatalog.autoMaximum.id,
      label: spreadsheetCommandCatalog.autoMaximum.label,
      icon: <ArrowUp size={16} />,
    },
    {
      functionName: 'min',
      id: spreadsheetCommandCatalog.autoMinimum.id,
      label: spreadsheetCommandCatalog.autoMinimum.label,
      icon: <ArrowDown size={16} />,
    },
  ];
  const primaryDisabled = !can.applyAutoSum('sum');
  const menuDisabled = primaryDisabled;

  return (
    <Popover
      label="更多自动计算方式"
      panelLabel="自动计算选项"
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
            aria-label={spreadsheetCommandCatalog.autoSum.label}
            aria-keyshortcuts={spreadsheetCommandCatalog.autoSum.shortcut.aria}
            title={`${spreadsheetCommandCatalog.autoSum.label}（${spreadsheetCommandCatalog.autoSum.shortcut.label}）`}
            disabled={primaryDisabled}
            onClick={() => commands.applyAutoSum('sum')}
          >
            <Sigma size={19} />
            <span>{spreadsheetCommandCatalog.autoSum.label}</span>
          </button>
          <button
            {...triggerProps}
            className={`work-spreadsheet-ribbon-split-disclosure${open ? ' active' : ''}`}
            title="更多自动计算方式"
          >
            <ChevronDown size={13} aria-hidden="true" />
          </button>
        </>
      )}
    >
      {(close) =>
        items.map(({ functionName, id, label, icon, shortcut }) => (
          <button
            key={id}
            type="button"
            role="menuitem"
            tabIndex={-1}
            aria-keyshortcuts={shortcut}
            disabled={primaryDisabled}
            onClick={() => {
              close();
              commands.applyAutoSum(functionName);
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

function SpreadsheetFillMenu({
  can,
  commands,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
}) {
  const items: readonly {
    direction: SpreadsheetCellFillDirection;
    id: string;
    label: string;
    icon: ReactNode;
    shortcut?: string;
  }[] = [
    {
      direction: 'down',
      id: spreadsheetCommandCatalog.fillDown.id,
      label: spreadsheetCommandCatalog.fillDown.label,
      icon: <ArrowDownToLine size={16} />,
      shortcut: spreadsheetCommandCatalog.fillDown.shortcut.aria,
    },
    {
      direction: 'right',
      id: spreadsheetCommandCatalog.fillRight.id,
      label: spreadsheetCommandCatalog.fillRight.label,
      icon: <ArrowRightToLine size={16} />,
      shortcut: spreadsheetCommandCatalog.fillRight.shortcut.aria,
    },
    {
      direction: 'up',
      id: spreadsheetCommandCatalog.fillUp.id,
      label: spreadsheetCommandCatalog.fillUp.label,
      icon: <ArrowUpToLine size={16} />,
    },
    {
      direction: 'left',
      id: spreadsheetCommandCatalog.fillLeft.id,
      label: spreadsheetCommandCatalog.fillLeft.label,
      icon: <ArrowLeftToLine size={16} />,
    },
  ];
  const disabled = items.every(
    ({ direction }) => !can.fillSelectedCells(direction),
  );

  return (
    <Popover
      label="填充"
      panelLabel="填充选项"
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
          title="填充"
        >
          <ArrowDownToLine size={19} />
          <span>填充</span>
        </button>
      )}
    >
      {(close) =>
        items.map(({ direction, id, label, icon, shortcut }) => (
          <button
            key={id}
            type="button"
            role="menuitem"
            tabIndex={-1}
            aria-keyshortcuts={shortcut}
            disabled={!can.fillSelectedCells(direction)}
            onClick={() => {
              close();
              commands.fillSelectedCells(direction);
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
