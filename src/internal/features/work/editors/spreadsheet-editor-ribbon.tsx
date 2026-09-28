import type { Cell, Selection } from '@fortune-sheet/core';
import {
  BarChart3,
  Bookmark,
  Calculator,
  Code2,
  Grid3X3,
  Heading,
  Link2,
  ListChecks,
  ListFilter,
  Palette,
  Printer,
  Redo2,
  RefreshCw,
  ShieldCheck,
  Sigma,
  SortAsc,
  SortDesc,
  SlidersHorizontal,
  Table2,
  TableProperties,
  Undo2,
} from 'lucide-react';
import { type ReactNode, useMemo } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import { spreadsheetChartCount } from '../work-spreadsheet-charts';
import { spreadsheetFormulaCount } from '../work-spreadsheet-formula-analysis';
import { spreadsheetPivotCount } from '../work-spreadsheet-pivots';
import { protectedSheetCount } from '../work-spreadsheet-protection';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetTable,
} from '../work-types';
import { useOfficeMessages } from './office-messages-context';
import { SpreadsheetAlignmentRibbonGroup } from './spreadsheet-alignment-ribbon';
import type { SpreadsheetResolvedCellBorders } from './spreadsheet-cell-border';
import { SpreadsheetCellStyleRibbon } from './spreadsheet-cell-style-ribbon';
import { SpreadsheetClipboardRibbonGroup } from './spreadsheet-clipboard-ribbon';
import {
  type SpreadsheetRibbonTabId,
  spreadsheetCommandCatalog,
} from './spreadsheet-command-catalog';
import {
  localizedSpreadsheetRibbonTabs,
  spreadsheetCommandLabel,
} from './spreadsheet-command-i18n';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { managedConditionalFormatCount } from './spreadsheet-conditional-format-panel';
import { SpreadsheetEditingRibbonGroup } from './spreadsheet-editing-ribbon';
import type { SpreadsheetFormatPainterMode } from './spreadsheet-format-painter';
import {
  SpreadsheetFontRibbonGroup,
  SpreadsheetNumberRibbonGroup,
} from './spreadsheet-home-format-ribbon';
import { spreadsheetPrintSettingCount } from './spreadsheet-print-settings-panel';
import { SpreadsheetRowsAndColumnsMenu } from './spreadsheet-rows-columns-ribbon';
import { SpreadsheetTableDesignRibbon } from './spreadsheet-table-ribbon';
import { SpreadsheetFreezePanesMenu } from './spreadsheet-view-ribbon';
import type { SpreadsheetWorkbookPanelView } from './spreadsheet-workbook-panel';
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
  activeTable = null,
  activeTableSheetId,
  autoFilterActive = false,
  can,
  commands,
  content,
  fileActions,
  findOpen,
  formatPainterMode = null,
  freezePanesActive = false,
  freezePanesSelection = defaultSpreadsheetFreezePanesSelection,
  formulaBarVisible = true,
  showFormulas = false,
  gridLinesVisible,
  headingsVisible = true,
  panelId,
  onTabChange,
  onToggleFormulaBar,
  onToggleShowFormulas,
  onToggleHeadings,
  onTogglePanel,
  panel,
  toolbarCell,
  toolbarCellBorders,
}: {
  activeTab: SpreadsheetRibbonTabId;
  activeTable?: WorkSpreadsheetTable | null;
  activeTableSheetId?: string;
  autoFilterActive?: boolean;
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
  content: WorkSpreadsheetContent;
  fileActions?: readonly WorkOfficeFileAction[];
  findOpen: boolean;
  formatPainterMode?: SpreadsheetFormatPainterMode | null;
  freezePanesActive?: boolean;
  freezePanesSelection?: Selection;
  formulaBarVisible?: boolean;
  showFormulas?: boolean;
  gridLinesVisible: boolean;
  headingsVisible?: boolean;
  panelId: string;
  onTabChange: (tab: SpreadsheetRibbonTabId) => void;
  onToggleFormulaBar?: () => void;
  onToggleShowFormulas?: () => void;
  onToggleHeadings?: () => void;
  onTogglePanel: (
    panel: SpreadsheetWorkbookPanelView,
    trigger: HTMLButtonElement,
  ) => void;
  panel: SpreadsheetWorkbookPanelView | null;
  toolbarCell: Cell | null | undefined;
  toolbarCellBorders?: SpreadsheetResolvedCellBorders;
}) {
  const messages = useOfficeMessages();
  const ribbonTabs = useMemo(
    () =>
      localizedSpreadsheetRibbonTabs(messages, {
        includeTableDesign: Boolean(activeTable),
      }),
    [activeTable, messages],
  );
  const formulaCount = useMemo(
    () => spreadsheetFormulaCount(content),
    [content],
  );
  const pivotCount = useMemo(() => spreadsheetPivotCount(content), [content]);
  return (
    <WorkOfficeRibbon
      ariaLabel={officeMessage(messages, 'spreadsheet.editorRibbon.aria')}
      tabs={ribbonTabs}
      defaultTab="home"
      activeTab={activeTab}
      onTabChange={onTabChange}
      adaptive
      collapsible
      fileActions={fileActions}
      quickAccessActions={[
        {
          id: spreadsheetCommandCatalog.undo.id,
          label: spreadsheetCommandLabel('undo', messages),
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
          label: spreadsheetCommandLabel('redo', messages),
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
            <SpreadsheetClipboardRibbonGroup
              can={can}
              commands={commands}
              formatPainterMode={formatPainterMode}
            />
            <SpreadsheetFontRibbonGroup
              can={can}
              commands={commands}
              toolbarCell={toolbarCell}
            />
            <SpreadsheetAlignmentRibbonGroup
              can={can}
              commands={commands}
              toolbarCell={toolbarCell}
            />
            <SpreadsheetNumberRibbonGroup
              can={can}
              commands={commands}
              toolbarCell={toolbarCell}
            />
            <WorkOfficeRibbonGroup label={officeMessage(messages, 'spreadsheet.editorRibbon.group.styles')}>
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
            <WorkOfficeRibbonGroup label={officeMessage(messages, 'spreadsheet.editorRibbon.group.cells')}>
              <SpreadsheetRowsAndColumnsMenu can={can} commands={commands} />
            </WorkOfficeRibbonGroup>
            <SpreadsheetEditingRibbonGroup
              can={can}
              commands={commands}
              findOpen={findOpen}
            />
          </>
        ),
        insert: (
          <>
            <WorkOfficeRibbonGroup label={officeMessage(messages, 'spreadsheet.editorRibbon.group.tables')} priority="high">
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.table.label}
                title={`${spreadsheetCommandCatalog.table.label}（${spreadsheetCommandCatalog.table.shortcut.label}）`}
                aria-keyshortcuts={
                  spreadsheetCommandCatalog.table.shortcut.aria
                }
                disabled={!can.openTable()}
                onClick={commands.openTable}
              >
                <Table2 size={19} />
              </WorkOfficeRibbonButton>
            </WorkOfficeRibbonGroup>
            <WorkOfficeRibbonGroup label={officeMessage(messages, 'spreadsheet.editorRibbon.group.links')} priority="high">
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
            <WorkOfficeRibbonGroup label={officeMessage(messages, 'spreadsheet.editorRibbon.group.charts')} priority="high">
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
          <WorkOfficeRibbonGroup label={officeMessage(messages, 'spreadsheet.editorRibbon.group.pageSetup')} priority="high">
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
            <WorkOfficeRibbonGroup label={officeMessage(messages, 'spreadsheet.editorRibbon.group.definedNames')} priority="high">
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
            <WorkOfficeRibbonGroup label={officeMessage(messages, 'spreadsheet.editorRibbon.group.calculation')} priority="high">
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
            <WorkOfficeRibbonGroup label={officeMessage(messages, 'spreadsheet.editorRibbon.group.sortFilter')} priority="high">
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
                label={spreadsheetCommandCatalog.customSort.label}
                disabled={!can.openCustomSort()}
                onClick={() => commands.openCustomSort()}
              >
                <SlidersHorizontal size={19} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.autoFilter.label}
                title={officeMessage(messages, 'spreadsheet.editorRibbon.autoFilterTitle', {
                  label: spreadsheetCommandCatalog.autoFilter.label,
                  shortcut: spreadsheetCommandCatalog.autoFilter.shortcut.label,
                  menuShortcut:
                    spreadsheetCommandCatalog.autoFilter.menuShortcut.label,
                })}
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
            <WorkOfficeRibbonGroup label={officeMessage(messages, 'spreadsheet.editorRibbon.group.dataTools')} priority="high">
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.dataValidation.label}
                disabled={!can.openDataValidation()}
                onClick={commands.openDataValidation}
              >
                <ListChecks size={19} />
              </WorkOfficeRibbonButton>
            </WorkOfficeRibbonGroup>
            <WorkOfficeRibbonGroup label={officeMessage(messages, 'spreadsheet.editorRibbon.group.analyze')}>
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
          <WorkOfficeRibbonGroup label={officeMessage(messages, 'spreadsheet.editorRibbon.group.protect')} priority="high">
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
            <WorkOfficeRibbonGroup label={officeMessage(messages, 'spreadsheet.editorRibbon.group.workbookView')} priority="high">
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.formulaBar.label}
                active={formulaBarVisible}
                disabled={!onToggleFormulaBar}
                title={formulaBarVisible ? officeMessage(messages, 'spreadsheet.editorRibbon.hideFormulaBar') : officeMessage(messages, 'spreadsheet.editorRibbon.showFormulaBar')}
                onClick={() => onToggleFormulaBar?.()}
              >
                <Sigma size={19} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.showFormulas.label}
                active={showFormulas}
                disabled={!onToggleShowFormulas}
                title={
                  showFormulas
                    ? officeMessage(messages, 'spreadsheet.editorRibbon.hideFormulas')
                    : officeMessage(messages, 'spreadsheet.editorRibbon.showFormulas')
                }
                aria-keyshortcuts={
                  spreadsheetCommandCatalog.showFormulas.shortcut.aria
                }
                onClick={() => onToggleShowFormulas?.()}
              >
                <Code2 size={19} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.gridLines.label}
                active={gridLinesVisible}
                disabled={!can.setGridLines(!gridLinesVisible)}
                title={gridLinesVisible ? officeMessage(messages, 'spreadsheet.editorRibbon.hideGridLines') : officeMessage(messages, 'spreadsheet.editorRibbon.showGridLines')}
                onClick={() => commands.setGridLines(!gridLinesVisible)}
              >
                <Grid3X3 size={19} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label={spreadsheetCommandCatalog.headings.label}
                active={headingsVisible}
                disabled={!onToggleHeadings}
                title={headingsVisible ? officeMessage(messages, 'spreadsheet.editorRibbon.hideHeadings') : officeMessage(messages, 'spreadsheet.editorRibbon.showHeadings')}
                onClick={() => onToggleHeadings?.()}
              >
                <Heading size={19} />
              </WorkOfficeRibbonButton>
            </WorkOfficeRibbonGroup>
            <WorkOfficeRibbonGroup label={officeMessage(messages, 'spreadsheet.editorRibbon.group.window')} priority="high">
              <SpreadsheetFreezePanesMenu
                active={freezePanesActive}
                can={can}
                commands={commands}
                selection={freezePanesSelection}
              />
            </WorkOfficeRibbonGroup>
          </>
        ),
        tableDesign:
          activeTable && activeTableSheetId ? (
            <SpreadsheetTableDesignRibbon
              can={can}
              commands={commands}
              sheetId={activeTableSheetId}
              table={activeTable}
            />
          ) : null,
      }}
    />
  );
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
