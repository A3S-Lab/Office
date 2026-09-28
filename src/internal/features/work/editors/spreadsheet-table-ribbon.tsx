import { Calculator, Columns3, Rows3, TableProperties } from 'lucide-react';
import { type KeyboardEvent, type ReactNode, useEffect, useState } from 'react';
import { Popover } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import { showToast } from '../../../state/app-state';
import type {
  WorkSpreadsheetTable,
  WorkSpreadsheetTableStyle,
} from '../work-types';
import {
  CommittedOfficeTextField,
  OfficeCheckbox,
  OfficeSelect,
} from './office-controls';
import { OfficeMenuGroup } from './office-menu-group';
import { moveOfficeGridMenuFocus } from './office-menu-keyboard';
import { useOfficeMessages } from './office-messages-context';
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
  const messages = useOfficeMessages();
  return (
    <>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'spreadsheet.table.ribbon.properties')}
        priority="high"
      >
        <SpreadsheetTableNameControl
          can={can}
          commands={commands}
          sheetId={sheetId}
          table={table}
        />
        <WorkOfficeRibbonButton
          label={officeMessage(
            messages,
            'spreadsheet.table.ribbon.convertToRange',
          )}
          disabled={!can.convertTableToRange(sheetId, table.id)}
          onClick={() => {
            if (!commands.convertTableToRange(sheetId, table.id)) {
              showToast(
                officeMessage(
                  messages,
                  'spreadsheet.table.ribbon.convertToRangeError',
                ),
                'error',
              );
            }
          }}
        >
          <TableProperties size={19} />
        </WorkOfficeRibbonButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'spreadsheet.table.ribbon.totalsGroup')}
        priority="high"
      >
        <SpreadsheetTableTotalsMenu
          can={can}
          commands={commands}
          sheetId={sheetId}
          table={table}
        />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'spreadsheet.table.ribbon.stylesGroup')}
        priority="high"
      >
        <SpreadsheetTableStyleGallery
          commands={commands}
          sheetId={sheetId}
          table={table}
        />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'spreadsheet.table.ribbon.styleOptions')}
      >
        <SpreadsheetTableOption
          label={officeMessage(
            messages,
            'spreadsheet.table.ribbon.firstColumn',
          )}
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
          label={officeMessage(
            messages,
            'spreadsheet.table.ribbon.lastColumn',
          )}
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
          label={officeMessage(messages, 'spreadsheet.table.ribbon.rowStripes')}
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
          label={officeMessage(
            messages,
            'spreadsheet.table.ribbon.columnStripes',
          )}
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
  const messages = useOfficeMessages();
  const totalsFunctionOptions = [
    {
      value: 'none',
      label: officeMessage(messages, 'spreadsheet.table.totals.none'),
    },
    ...SPREADSHEET_TABLE_TOTALS_FUNCTIONS.map((value) => ({
      value,
      label: spreadsheetTableTotalsFunctionLabel(value),
    })),
  ] as const;

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
      showToast(
        officeMessage(messages, 'spreadsheet.table.totals.invalid'),
        'error',
      );
    }
  };

  const toggleTotalsRow = () => {
    const patch = { totalsRow: !table.totalsRow };
    if (
      !can.updateTable(sheetId, table.id, patch) ||
      !commands.updateTable(sheetId, table.id, patch)
    ) {
      showToast(
        officeMessage(
          messages,
          table.totalsRow
            ? 'spreadsheet.table.totals.cannotDisable'
            : 'spreadsheet.table.totals.targetOccupied',
        ),
        'error',
      );
    }
  };

  return (
    <Popover
      label={officeMessage(messages, 'spreadsheet.table.totals.label')}
      panelLabel={officeMessage(messages, 'spreadsheet.table.totals.panel')}
      panelRole="dialog"
      portal
      focusFirstOnOpen
      className="work-spreadsheet-table-totals-root"
      panelClassName="work-spreadsheet-table-totals-menu"
      placement="bottom-end"
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          className={`with-label work-spreadsheet-table-totals-trigger${open || table.totalsRow ? ' active' : ''}`}
          title={officeMessage(
            messages,
            table.totalsRow
              ? 'spreadsheet.table.totals.titleOn'
              : 'spreadsheet.table.totals.titleOff',
          )}
        >
          <Calculator size={19} />
          <span>
            {officeMessage(messages, 'spreadsheet.table.totals.label')}
          </span>
        </button>
      )}
    >
      {(close) => (
        <div className="work-spreadsheet-table-totals-content">
          <OfficeCheckbox
            ariaLabel={officeMessage(
              messages,
              'spreadsheet.table.totals.enableAria',
            )}
            checked={table.totalsRow}
            disabled={
              !table.totalsRow &&
              !can.updateTable(sheetId, table.id, { totalsRow: true })
            }
            onCheckedChange={() => {
              toggleTotalsRow();
            }}
          >
            {officeMessage(messages, 'spreadsheet.table.totals.enable')}
          </OfficeCheckbox>
          <p className="work-spreadsheet-table-totals-hint">
            {officeMessage(messages, 'spreadsheet.table.totals.hint')}
          </p>
          <div className="work-spreadsheet-table-totals-columns">
            {table.columns.map((column, offset) => {
              const selected = column.totalsFormula
                ? 'custom'
                : (column.totalsFunction ?? 'none');
              const formulaEnabled = selected === 'custom';
              const labelEnabled = selected === 'none';
              const committedLabel = column.totalsLabel ?? '';
              const committedFormula = column.totalsFormula ?? '';
              return (
                <div
                  className="work-spreadsheet-table-totals-column"
                  key={`${table.id}-${offset}`}
                >
                  <strong>{column.name}</strong>
                  <div className="work-office-field">
                    <span>
                      {officeMessage(
                        messages,
                        'spreadsheet.table.totals.function',
                      )}
                    </span>
                    <OfficeSelect
                      ariaLabel={officeMessage(
                        messages,
                        'spreadsheet.table.totals.functionAria',
                        { name: column.name },
                      )}
                      disabled={!table.totalsRow}
                      value={selected}
                      options={totalsFunctionOptions}
                      onValueChange={(value) => {
                        if (value === 'none') {
                          patchColumn(offset, {
                            totalsFunction: null,
                            totalsFormula: null,
                          });
                        } else if (value === 'custom') {
                          const formula =
                            committedFormula ||
                            `=SUM(${table.name}[${escapeTotalsColumnName(column.name)}])`;
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
                    />
                  </div>
                  <div className="work-office-field">
                    <span>
                      {officeMessage(messages, 'spreadsheet.table.totals.tag')}
                    </span>
                    <CommittedOfficeTextField
                      aria-label={officeMessage(
                        messages,
                        'spreadsheet.table.totals.tagAria',
                        { name: column.name },
                      )}
                      disabled={!table.totalsRow || !labelEnabled}
                      value={committedLabel}
                      formatValue={(value) => value}
                      parseValue={(draft) => draft}
                      onValueCommit={(value) => {
                        if (!labelEnabled) return;
                        patchColumn(offset, {
                          totalsLabel: value.trim() || null,
                        });
                      }}
                    />
                  </div>
                  <div className="work-office-field">
                    <span>
                      {officeMessage(
                        messages,
                        'spreadsheet.table.totals.formula',
                      )}
                    </span>
                    <CommittedOfficeTextField
                      aria-label={officeMessage(
                        messages,
                        'spreadsheet.table.totals.formulaAria',
                        { name: column.name },
                      )}
                      disabled={!table.totalsRow || !formulaEnabled}
                      placeholder="=SUM(Table[Column])"
                      value={committedFormula}
                      formatValue={(value) => value}
                      parseValue={(draft) => draft}
                      onValueCommit={(value) => {
                        if (!formulaEnabled) return;
                        patchColumn(offset, {
                          totalsFunction: 'custom',
                          totalsFormula: value.trim() || null,
                        });
                      }}
                    />
                  </div>
                </div>
              );
            })}
          </div>
          <button
            type="button"
            className="work-spreadsheet-table-totals-close"
            onClick={close}
          >
            {officeMessage(messages, 'spreadsheet.table.totals.done')}
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
  const messages = useOfficeMessages();
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
      showToast(
        officeMessage(messages, 'spreadsheet.table.nameInvalid'),
        'error',
      );
    }
  };

  const dirty = name !== table.name;

  return (
    <label className="work-spreadsheet-table-name">
      <span>{officeMessage(messages, 'spreadsheet.table.nameLabel')}</span>
      <input
        type="text"
        aria-label={officeMessage(messages, 'spreadsheet.table.nameLabel')}
        autoCapitalize="none"
        spellCheck={false}
        data-office-escape-consumer={dirty || undefined}
        value={name}
        onBlur={commit}
        onChange={(event) => setName(event.currentTarget.value)}
        onKeyDown={(event) => {
          if (event.key === 'Enter') {
            event.preventDefault();
            commit();
          } else if (event.key === 'Escape' && dirty) {
            event.preventDefault();
            event.stopPropagation();
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
  const messages = useOfficeMessages();
  const choices = spreadsheetTableStyleChoices();
  const selected = choices.find((choice) =>
    spreadsheetTableUsesStyle(table.style, choice.style),
  );
  const families = [
    {
      id: 'light' as const,
      label: officeMessage(messages, 'spreadsheet.table.style.family.light'),
    },
    {
      id: 'medium' as const,
      label: officeMessage(messages, 'spreadsheet.table.style.family.medium'),
    },
    {
      id: 'dark' as const,
      label: officeMessage(messages, 'spreadsheet.table.style.family.dark'),
    },
  ];
  return (
    <Popover
      label={officeMessage(messages, 'spreadsheet.table.style.label')}
      panelLabel={officeMessage(messages, 'spreadsheet.table.style.panel')}
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
          title={officeMessage(messages, 'spreadsheet.table.style.title', {
            current:
              selected?.label ??
              officeMessage(messages, 'spreadsheet.table.style.none'),
          })}
        >
          <SpreadsheetTableStylePreview choice={selected ?? choices[0]} />
          <span>
            {officeMessage(messages, 'spreadsheet.table.style.label')}
          </span>
        </button>
      )}
    >
      {(close) => (
        <>
          {families.map((family) => (
            <OfficeMenuGroup
              key={family.id}
              className="work-spreadsheet-table-style-family"
              data-office-menu-grid
              ariaLabel={family.label}
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
                        aria-label={officeMessage(
                          messages,
                          'spreadsheet.table.style.applyAria',
                          { label: choice.label },
                        )}
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
            </OfficeMenuGroup>
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
