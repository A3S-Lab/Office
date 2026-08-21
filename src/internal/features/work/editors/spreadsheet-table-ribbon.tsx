import { Columns3, Rows3, TableProperties } from 'lucide-react';
import { type KeyboardEvent, type ReactNode, useEffect, useState } from 'react';
import { Popover } from '../../../design-system/primitives';
import { showToast } from '../../../state/app-state';
import type {
  WorkSpreadsheetTable,
  WorkSpreadsheetTableStyle,
} from '../work-types';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { moveOfficeGridMenuFocus } from './office-menu-keyboard';
import { spreadsheetTableStyleChoices } from './spreadsheet-table-style';
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
