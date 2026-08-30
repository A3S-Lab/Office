import { Calculator, Columns3, Rows3, TableProperties } from 'lucide-react';
import { type KeyboardEvent, type ReactNode, useEffect, useState } from 'react';
import { Popover } from '../../../design-system/primitives';
import { showToast } from '../../../state/app-state';
import type {
  WorkSpreadsheetTable,
  WorkSpreadsheetTableStyle,
} from '../work-types';
import { OfficeCheckbox } from './office-controls';
import { moveOfficeGridMenuFocus } from './office-menu-keyboard';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { spreadsheetTableStyleChoices } from './spreadsheet-table-style';
import {
  SPREADSHEET_TABLE_TOTALS_FUNCTIONS,
  spreadsheetTableTotalsFunctionLabel,
} from './spreadsheet-table-totals';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

export function SpreadsheetTableDesignRibbon({
  can,
  commands,
  sheetId,
  table,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
  sheetId: string;
  table: WorkSpreadsheetTable;
}) {
  return (
    <>
      <WorkOfficeRibbonGroup label="属性" priority="high">
        <SpreadsheetTableNameControl
          can={can}
          commands={commands}
          sheetId={sheetId}
          table={table}
        />
        <WorkOfficeRibbonButton
          label="转换为区域"
          disabled={!can.convertTableToRange(sheetId, table.id)}
          onClick={() => {
            if (!commands.convertTableToRange(sheetId, table.id)) {
              showToast(
                '无法转换为区域。请先移除引用此表格的结构化引用。',
                'error',
              );
            }
          }}
        >
          <TableProperties size={19} />
        </WorkOfficeRibbonButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="汇总行" priority="high">
        <SpreadsheetTableTotalsMenu
          can={can}
          commands={commands}
          sheetId={sheetId}
          table={table}
        />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="表格样式" priority="high">
        <SpreadsheetTableStyleGallery
          commands={commands}
          sheetId={sheetId}
          table={table}
        />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="表格样式选项">
        <SpreadsheetTableOption
          label="首列"
          active={table.showFirstColumn}
          icon={<Columns3 size={17} />}
          disabled={
            !can.updateTable(sheetId, table.id, {
              showFirstColumn: !table.showFirstColumn,
            })
          }
          onToggle={() =>
            commands.updateTable(sheetId, table.id, {
              showFirstColumn: !table.showFirstColumn,
            })
          }
        />
        <SpreadsheetTableOption
          label="末列"
          active={table.showLastColumn}
          icon={<Columns3 size={17} />}
          disabled={
            !can.updateTable(sheetId, table.id, {
              showLastColumn: !table.showLastColumn,
            })
          }
          onToggle={() =>
            commands.updateTable(sheetId, table.id, {
              showLastColumn: !table.showLastColumn,
            })
          }
        />
        <SpreadsheetTableOption
          label="行条纹"
          active={table.showRowStripes}
          icon={<Rows3 size={17} />}
          disabled={
            !can.updateTable(sheetId, table.id, {
              showRowStripes: !table.showRowStripes,
            })
          }
          onToggle={() =>
            commands.updateTable(sheetId, table.id, {
              showRowStripes: !table.showRowStripes,
            })
          }
        />
        <SpreadsheetTableOption
          label="列条纹"
          active={table.showColumnStripes}
          icon={<Columns3 size={17} />}
          disabled={
            !can.updateTable(sheetId, table.id, {
              showColumnStripes: !table.showColumnStripes,
            })
          }
          onToggle={() =>
            commands.updateTable(sheetId, table.id, {
              showColumnStripes: !table.showColumnStripes,
            })
          }
        />
      </WorkOfficeRibbonGroup>
    </>
  );
}

const totalsFunctionOptions = [
  { value: 'none', label: '不汇总' },
  ...SPREADSHEET_TABLE_TOTALS_FUNCTIONS.map((value) => ({
    value,
    label: spreadsheetTableTotalsFunctionLabel(value),
  })),
] as const;

function SpreadsheetTableTotalsMenu({
  can,
  commands,
  sheetId,
  table,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
  sheetId: string;
  table: WorkSpreadsheetTable;
}) {
  const [drafts, setDrafts] = useState<
    Record<number, { formula: string; label: string }>
  >({});
  useEffect(() => {
    const next: Record<number, { formula: string; label: string }> = {};
    table.columns.forEach((column, offset) => {
      next[offset] = {
        formula: column.totalsFormula ?? '',
        label: column.totalsLabel ?? '',
      };
    });
    setDrafts(next);
  }, [table.id, table.columns]);

  const patchColumn = (
    offset: number,
    patch: {
      totalsFormula?: string | null;
      totalsFunction?:
        | WorkSpreadsheetTable['columns'][number]['totalsFunction']
        | null;
      totalsLabel?: string | null;
    },
  ) => {
    const designPatch = { totalsColumns: { [offset]: patch } };
    if (
      !can.updateTable(sheetId, table.id, designPatch) ||
      !commands.updateTable(sheetId, table.id, designPatch)
    ) {
      showToast('汇总行设置无效，请检查函数、标签或公式。', 'error');
    }
  };

  const toggleTotalsRow = () => {
    const patch = { totalsRow: !table.totalsRow };
    if (
      !can.updateTable(sheetId, table.id, patch) ||
      !commands.updateTable(sheetId, table.id, patch)
    ) {
      showToast(
        table.totalsRow
          ? '无法关闭汇总行。'
          : '汇总行目标区域已有内容，请先清空该行。',
        'error',
      );
    }
  };

  return (
    <Popover
      label="汇总行"
      panelLabel="表格汇总行设置"
      panelRole="dialog"
      portal
      className="work-spreadsheet-table-totals-root"
      panelClassName="work-spreadsheet-table-totals-menu"
      placement="bottom-end"
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          className={`with-label work-spreadsheet-table-totals-trigger${open || table.totalsRow ? ' active' : ''}`}
          aria-pressed={table.totalsRow}
          title="设置表格汇总行"
        >
          <Calculator size={19} />
          <span>汇总行</span>
        </button>
      )}
    >
      {(close) => (
        <div className="work-spreadsheet-table-totals-content">
          <OfficeCheckbox
            ariaLabel="启用汇总行"
            checked={table.totalsRow}
            disabled={
              !table.totalsRow &&
              !can.updateTable(sheetId, table.id, { totalsRow: true })
            }
            onCheckedChange={() => {
              toggleTotalsRow();
            }}
          >
            启用汇总行
          </OfficeCheckbox>
          <p className="work-spreadsheet-table-totals-hint">
            汇总行位于表格末尾，函数会随筛选结果更新。
          </p>
          <div className="work-spreadsheet-table-totals-columns">
            {table.columns.map((column, offset) => {
              const selected = column.totalsFormula
                ? 'custom'
                : (column.totalsFunction ?? 'none');
              const draft = drafts[offset] ?? {
                formula: column.totalsFormula ?? '',
                label: column.totalsLabel ?? '',
              };
              const formulaEnabled = selected === 'custom';
              const labelEnabled = selected === 'none';
              return (
                <div
                  className="work-spreadsheet-table-totals-column"
                  key={`${table.id}-${offset}`}
                >
                  <strong>{column.name}</strong>
                  <label>
                    <span>函数</span>
                    <select
                      aria-label={`${column.name} 汇总函数`}
                      disabled={!table.totalsRow}
                      value={selected}
                      onChange={(event) => {
                        const value = event.currentTarget.value;
                        if (value === 'none') {
                          patchColumn(offset, {
                            totalsFunction: null,
                            totalsFormula: null,
                          });
                        } else if (value === 'custom') {
                          const formula =
                            draft.formula ||
                            `=SUM(${table.name}[${escapeTotalsColumnName(column.name)}])`;
                          setDrafts((current) => ({
                            ...current,
                            [offset]: { ...draft, formula },
                          }));
                          patchColumn(offset, {
                            totalsFunction: 'custom',
                            totalsFormula: formula,
                            totalsLabel: null,
                          });
                        } else {
                          patchColumn(offset, {
                            totalsFunction: value as NonNullable<
                              typeof column.totalsFunction
                            >,
                            totalsFormula: null,
                            totalsLabel: null,
                          });
                        }
                      }}
                    >
                      {totalsFunctionOptions.map((option) => (
                        <option key={option.value} value={option.value}>
                          {option.label}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label>
                    <span>标签</span>
                    <input
                      aria-label={`${column.name} 汇总标签`}
                      disabled={!table.totalsRow || !labelEnabled}
                      value={draft.label}
                      onChange={(event) => {
                        const value = event.currentTarget.value;
                        setDrafts((current) => ({
                          ...current,
                          [offset]: { ...draft, label: value },
                        }));
                      }}
                      onBlur={() => {
                        if (!labelEnabled) return;
                        patchColumn(offset, {
                          totalsLabel: draft.label.trim() || null,
                        });
                      }}
                      onKeyDown={(event) => {
                        if (event.key === 'Enter') event.currentTarget.blur();
                      }}
                    />
                  </label>
                  <label>
                    <span>自定义公式</span>
                    <input
                      aria-label={`${column.name} 汇总公式`}
                      disabled={!table.totalsRow || !formulaEnabled}
                      placeholder="=SUM(Table[Column])"
                      value={draft.formula}
                      onChange={(event) => {
                        const value = event.currentTarget.value;
                        setDrafts((current) => ({
                          ...current,
                          [offset]: { ...draft, formula: value },
                        }));
                      }}
                      onBlur={() => {
                        if (!formulaEnabled) return;
                        patchColumn(offset, {
                          totalsFunction: 'custom',
                          totalsFormula: draft.formula.trim() || null,
                        });
                      }}
                      onKeyDown={(event) => {
                        if (event.key === 'Enter') event.currentTarget.blur();
                      }}
                    />
                  </label>
                </div>
              );
            })}
          </div>
          <button
            type="button"
            className="work-spreadsheet-table-totals-close"
            onClick={close}
          >
            完成
          </button>
        </div>
      )}
    </Popover>
  );
}

function escapeTotalsColumnName(value: string): string {
  return value.replaceAll(']', ']]');
}

function SpreadsheetTableNameControl({
  can,
  commands,
  sheetId,
  table,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
  sheetId: string;
  table: WorkSpreadsheetTable;
}) {
  const [name, setName] = useState(table.name);
  useEffect(() => setName(table.name), [table.id, table.name]);

  const commit = () => {
    const candidate = name.trim();
    if (!candidate || candidate === table.name) {
      setName(table.name);
      return;
    }
    const patch = { name: candidate };
    if (
      !can.updateTable(sheetId, table.id, patch) ||
      !commands.updateTable(sheetId, table.id, patch)
    ) {
      setName(table.name);
      showToast('表格名称无效或已被使用。', 'error');
    }
  };

  return (
    <label className="work-spreadsheet-table-name">
      <span>表格名称</span>
      <input
        type="text"
        aria-label="表格名称"
        autoCapitalize="none"
        spellCheck={false}
        value={name}
        onBlur={commit}
        onChange={(event) => setName(event.currentTarget.value)}
        onKeyDown={(event) => {
          if (event.key === 'Enter') {
            event.preventDefault();
            commit();
          } else if (event.key === 'Escape') {
            event.preventDefault();
            setName(table.name);
          }
        }}
      />
    </label>
  );
}

function SpreadsheetTableStyleGallery({
  commands,
  sheetId,
  table,
}: {
  commands: SpreadsheetEditorCommands;
  sheetId: string;
  table: WorkSpreadsheetTable;
}) {
  const choices = spreadsheetTableStyleChoices();
  const selected = choices.find((choice) =>
    spreadsheetTableUsesStyle(table.style, choice.style),
  );
  const families = [
    { id: 'light' as const, label: '浅色' },
    { id: 'medium' as const, label: '中等' },
    { id: 'dark' as const, label: '深色' },
  ];
  return (
    <Popover
      label="表格样式"
      panelLabel="表格样式库"
      panelRole="menu"
      portal
      className="work-spreadsheet-table-style-root"
      panelClassName="work-spreadsheet-table-style-menu"
      focusFirstOnOpen
      onPanelKeyDown={moveSpreadsheetTableStyleFocus}
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          type="button"
          className={`with-label work-spreadsheet-table-style-trigger${open ? ' active' : ''}`}
          title={`表格样式（当前：${selected?.label ?? '无'}）`}
        >
          <SpreadsheetTableStylePreview choice={selected ?? choices[0]} />
          <span>表格样式</span>
        </button>
      )}
    >
      {(close) => (
        <>
          {families.map((family) => (
            <fieldset
              key={family.id}
              className="work-spreadsheet-table-style-family"
              data-office-menu-grid
              aria-label={family.label}
            >
              <span>{family.label}</span>
              <div>
                {choices
                  .filter((choice) => choice.style.family === family.id)
                  .map((choice) => {
                    const checked = spreadsheetTableUsesStyle(
                      table.style,
                      choice.style,
                    );
                    return (
                      <button
                        key={choice.ooxmlName}
                        type="button"
                        role="menuitemradio"
                        tabIndex={-1}
                        aria-label={`应用表格样式：${choice.label}`}
                        aria-checked={checked}
                        title={choice.label}
                        onClick={() => {
                          close();
                          commands.updateTable(sheetId, table.id, {
                            style: choice.style,
                          });
                        }}
                      >
                        <SpreadsheetTableStylePreview choice={choice} />
                      </button>
                    );
                  })}
              </div>
            </fieldset>
          ))}
        </>
      )}
    </Popover>
  );
}

function moveSpreadsheetTableStyleFocus(
  event: KeyboardEvent<HTMLElement>,
): boolean {
  const active = document.activeElement;
  const group = [
    ...event.currentTarget.querySelectorAll<HTMLElement>(
      '[data-office-menu-grid]',
    ),
  ].find((candidate) => candidate.contains(active));
  const grid = group?.querySelector<HTMLElement>(':scope > div');
  const template = grid ? getComputedStyle(grid).gridTemplateColumns : '';
  const repeated = template.match(/^repeat\(\s*(\d+)/)?.[1];
  const resolved = template
    .split(/\s+/)
    .filter((value) => value && value !== 'none').length;
  const columns = repeated ? Number(repeated) : resolved || 7;
  return moveOfficeGridMenuFocus(event, Math.max(1, columns));
}

function spreadsheetTableUsesStyle(
  tableStyle: WorkSpreadsheetTableStyle,
  choice: Exclude<WorkSpreadsheetTableStyle, { family: 'none' }>,
): boolean {
  return (
    tableStyle.family !== 'none' &&
    tableStyle.family === choice.family &&
    tableStyle.number === choice.number
  );
}

function SpreadsheetTableStylePreview({
  choice,
}: {
  choice: ReturnType<typeof spreadsheetTableStyleChoices>[number];
}) {
  return (
    <span className="work-spreadsheet-table-style-preview" aria-hidden="true">
      <i style={{ backgroundColor: choice.palette.header }} />
      <i style={{ backgroundColor: choice.palette.primaryRow }} />
      <i style={{ backgroundColor: choice.palette.secondaryRow }} />
    </span>
  );
}

function SpreadsheetTableOption({
  active,
  disabled,
  icon,
  label,
  onToggle,
}: {
  active: boolean;
  disabled: boolean;
  icon: ReactNode;
  label: string;
  onToggle: () => boolean;
}) {
  return (
    <WorkOfficeRibbonButton
      label={label}
      active={active}
      disabled={disabled}
      onClick={onToggle}
    >
      {icon}
    </WorkOfficeRibbonButton>
  );
}
