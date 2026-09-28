import {
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  AlignCenter,
  AlignLeft,
  AlignRight,
} from 'lucide-react';
import type { Dispatch, SetStateAction } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import type {
  DocumentTableAlignment,
  DocumentTableCellMarginSide,
  DocumentTablePreferredWidthType,
} from '../work-document-table-geometry';
import type { DocumentTablePropertiesSource } from './document-table-properties-dialog-model';
import {
  draftForColumnWidthType,
  draftForTableWidthType,
  type DocumentTablePropertiesDraft,
  type DocumentTablePropertiesErrors,
  type DocumentTablePropertiesTab,
} from './document-table-properties-dialog-model';
import type { DocumentTableColumnWidthType } from '../work-document-table-column-widths';
import {
  OfficeCheckbox,
  OfficeNumberField,
  OfficeSelect,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';

interface SectionProps {
  draft: DocumentTablePropertiesDraft;
  setDraft: Dispatch<SetStateAction<DocumentTablePropertiesDraft>>;
  errors: DocumentTablePropertiesErrors;
  source: DocumentTablePropertiesSource;
}

export function DocumentTablePropertiesTabs({
  activeTab,
  idBase,
  onTabChange,
}: {
  activeTab: DocumentTablePropertiesTab;
  idBase: string;
  onTabChange: (tab: DocumentTablePropertiesTab) => void;
}) {
  const messages = useOfficeMessages();
  return (
    <div
      className="work-document-table-properties-tabs"
      role="tablist"
      aria-label={officeMessage(messages, 'document.table.properties.tabsAria')}
    >
      {tablePropertiesTabs(messages).map((tab) => (
        <button
          key={tab.value}
          type="button"
          id={`${idBase}-${tab.value}-tab`}
          role="tab"
          aria-controls={`${idBase}-${tab.value}-panel`}
          aria-selected={activeTab === tab.value}
          tabIndex={activeTab === tab.value ? 0 : -1}
          onClick={() => onTabChange(tab.value)}
        >
          {tab.label}
        </button>
      ))}
    </div>
  );
}

export function DocumentTablePropertiesPanel({
  activeTab,
  idBase,
  ...props
}: SectionProps & {
  activeTab: DocumentTablePropertiesTab;
  idBase: string;
}) {
  return (
    <section
      id={`${idBase}-${activeTab}-panel`}
      className="work-document-table-properties-panel"
      role="tabpanel"
      aria-labelledby={`${idBase}-${activeTab}-tab`}
      tabIndex={-1}
    >
      {activeTab === 'table' && <TableSection {...props} />}
      {activeTab === 'row' && <RowSection {...props} />}
      {activeTab === 'column' && <ColumnSection {...props} />}
      {activeTab === 'cell' && <CellSection {...props} />}
    </section>
  );
}

function TableSection({ draft, setDraft, errors, source }: SectionProps) {
  const messages = useOfficeMessages();
  const cm = officeMessage(messages, 'document.table.properties.unit.cm');
  return (
    <>
      <fieldset className="work-document-table-properties-section">
        <legend>
          {officeMessage(messages, 'document.table.properties.preferredWidth')}
        </legend>
        <div className="work-document-table-properties-choice-grid width">
          {widthOptions(messages).map((option) => (
            <label key={option.value}>
              <input
                type="radio"
                name="table-properties-width"
                value={option.value}
                checked={draft.table.widthType === option.value}
                data-autofocus={
                  draft.table.widthType === option.value || undefined
                }
                onChange={() =>
                  setDraft((current) =>
                    draftForTableWidthType(
                      current,
                      option.value,
                      source.renderedTableWidth,
                    ),
                  )
                }
              />
              <span>{option.label}</span>
            </label>
          ))}
        </div>
        {draft.table.widthType !== 'auto' && (
          <NumberRow
            label={officeMessage(
              messages,
              'document.table.properties.widthLabel',
            )}
            ariaLabel={officeMessage(
              messages,
              draft.table.widthType === 'percent'
                ? 'document.table.properties.tableWidthPercentAria'
                : 'document.table.properties.tableWidthCmAria',
            )}
            value={draft.table.width}
            unit={draft.table.widthType === 'percent' ? '%' : cm}
            min={draft.table.widthType === 'percent' ? 1 : 0.5}
            max={draft.table.widthType === 'percent' ? 100 : 30}
            step={draft.table.widthType === 'percent' ? 1 : 0.1}
            invalid={Boolean(errors.tableWidth)}
            onValueChange={(width) =>
              setDraft((current) => ({
                ...current,
                table: { ...current.table, width },
              }))
            }
          />
        )}
        {errors.tableWidth && <p role="alert">{errors.tableWidth}</p>}
      </fieldset>

      <fieldset className="work-document-table-properties-section">
        <legend>
          {officeMessage(messages, 'document.table.properties.position')}
        </legend>
        <div className="work-document-table-properties-choice-grid alignment">
          {alignmentOptions(messages).map((option) => {
            const Icon = option.icon;
            return (
              <label key={option.value}>
                <input
                  type="radio"
                  name="table-properties-alignment"
                  value={option.value}
                  checked={draft.table.alignment === option.value}
                  onChange={() =>
                    setDraft((current) => ({
                      ...current,
                      table: {
                        ...current.table,
                        alignment: option.value,
                      },
                    }))
                  }
                />
                <span>
                  <Icon size={16} aria-hidden="true" />
                  {option.label}
                </span>
              </label>
            );
          })}
        </div>
        <NumberRow
          label={officeMessage(messages, 'document.table.properties.indent')}
          ariaLabel={officeMessage(
            messages,
            'document.table.properties.indentAria',
          )}
          value={draft.table.indent}
          unit={cm}
          min={0}
          max={30}
          step={0.1}
          disabled={draft.table.alignment !== 'left'}
          invalid={Boolean(errors.tableIndent)}
          onValueChange={(indent) =>
            setDraft((current) => ({
              ...current,
              table: { ...current.table, indent },
            }))
          }
        />
        {errors.tableIndent && <p role="alert">{errors.tableIndent}</p>}
      </fieldset>
    </>
  );
}

function RowSection({ draft, setDraft, errors, source }: SectionProps) {
  const messages = useOfficeMessages();
  const cm = officeMessage(messages, 'document.table.properties.unit.cm');
  return (
    <>
      <fieldset className="work-document-table-properties-section">
        <legend>
          {officeMessage(messages, 'document.table.properties.rowSize')}
        </legend>
        <OfficeCheckbox
          ariaLabel={officeMessage(
            messages,
            'document.table.properties.specifyRowHeightAria',
          )}
          checked={draft.row.heightEnabled}
          onCheckedChange={(heightEnabled) =>
            setDraft((current) => ({
              ...current,
              row: { ...current.row, heightEnabled },
            }))
          }
        >
          {officeMessage(
            messages,
            'document.table.properties.specifyRowHeight',
          )}
        </OfficeCheckbox>
        <NumberRow
          label={officeMessage(messages, 'document.table.properties.rowHeight')}
          ariaLabel={officeMessage(
            messages,
            'document.table.properties.rowHeightAria',
          )}
          value={draft.row.height}
          unit={cm}
          min={0.5}
          max={30}
          step={0.1}
          disabled={!draft.row.heightEnabled}
          invalid={Boolean(errors.rowHeight)}
          onValueChange={(height) =>
            setDraft((current) => ({
              ...current,
              row: { ...current.row, height },
            }))
          }
        />
        <div className="work-document-table-properties-select-row">
          <span>
            {officeMessage(
              messages,
              'document.table.properties.rowHeightRule',
            )}
          </span>
          <OfficeSelect
            ariaLabel={officeMessage(
              messages,
              'document.table.properties.rowHeightRuleAria',
            )}
            value={draft.row.heightRule}
            options={rowHeightRuleOptions(messages)}
            disabled={!draft.row.heightEnabled}
            onValueChange={(heightRule) =>
              setDraft((current) => ({
                ...current,
                row: { ...current.row, heightRule },
              }))
            }
          />
        </div>
        {errors.rowHeight && <p role="alert">{errors.rowHeight}</p>}
      </fieldset>

      <fieldset className="work-document-table-properties-section">
        <legend>
          {officeMessage(messages, 'document.table.properties.pagination')}
        </legend>
        <div className="work-document-table-properties-checkboxes">
          <OfficeCheckbox
            ariaLabel={officeMessage(
              messages,
              'document.table.properties.allowSplitAria',
            )}
            checked={!draft.row.cantSplit}
            onCheckedChange={(allowSplit) =>
              setDraft((current) => ({
                ...current,
                row: { ...current.row, cantSplit: !allowSplit },
              }))
            }
          >
            {officeMessage(messages, 'document.table.properties.allowSplit')}
          </OfficeCheckbox>
          <OfficeCheckbox
            ariaLabel={officeMessage(
              messages,
              'document.table.properties.repeatHeaderAria',
            )}
            checked={draft.row.repeatHeader}
            disabled={!source.canRepeatHeader}
            onCheckedChange={(repeatHeader) =>
              setDraft((current) => ({
                ...current,
                row: { ...current.row, repeatHeader },
              }))
            }
          >
            {officeMessage(messages, 'document.table.properties.repeatHeader')}
          </OfficeCheckbox>
        </div>
      </fieldset>
    </>
  );
}

function ColumnSection({ draft, setDraft, errors, source }: SectionProps) {
  const messages = useOfficeMessages();
  const cm = officeMessage(messages, 'document.table.properties.unit.cm');
  return (
    <fieldset className="work-document-table-properties-section">
      <legend>
        {officeMessage(messages, 'document.table.properties.columnSize')}
      </legend>
      <div className="work-document-table-properties-choice-grid width">
        {columnWidthOptions(messages).map((option) => (
          <label key={option.value}>
            <input
              type="radio"
              name="table-properties-column-width-type"
              value={option.value}
              checked={draft.column.widthType === option.value}
              onChange={() =>
                setDraft((current) =>
                  draftForColumnWidthType(current, option.value, source),
                )
              }
            />
            <span>{option.label}</span>
          </label>
        ))}
      </div>
      <NumberRow
        label={officeMessage(
          messages,
          'document.table.properties.columnWidth',
        )}
        ariaLabel={officeMessage(
          messages,
          draft.column.widthType === 'percent'
            ? 'document.table.properties.columnWidthPercentAria'
            : 'document.table.properties.columnWidthCmAria',
        )}
        value={draft.column.width}
        unit={draft.column.widthType === 'percent' ? '%' : cm}
        min={draft.column.widthType === 'percent' ? 1 : 0.5}
        max={draft.column.widthType === 'percent' ? 100 : 30}
        step={draft.column.widthType === 'percent' ? 1 : 0.1}
        invalid={Boolean(errors.columnWidth)}
        onValueChange={(width) =>
          setDraft((current) => ({
            ...current,
            column: { ...current.column, width },
          }))
        }
      />
      {errors.columnWidth && <p role="alert">{errors.columnWidth}</p>}
    </fieldset>
  );
}

function CellSection({ draft, setDraft, errors }: SectionProps) {
  const messages = useOfficeMessages();
  const cm = officeMessage(messages, 'document.table.properties.unit.cm');
  return (
    <>
      <fieldset className="work-document-table-properties-section">
        <legend>
          {officeMessage(messages, 'document.table.properties.verticalAlign')}
        </legend>
        <div className="work-document-table-properties-choice-grid alignment">
          {verticalAlignmentOptions(messages).map((option) => {
            const Icon = option.icon;
            return (
              <label key={option.value}>
                <input
                  type="radio"
                  name="table-properties-vertical-alignment"
                  value={option.value}
                  checked={draft.cell.verticalAlign === option.value}
                  onChange={() =>
                    setDraft((current) => ({
                      ...current,
                      cell: {
                        ...current.cell,
                        verticalAlign: option.value,
                      },
                    }))
                  }
                />
                <span>
                  <Icon size={16} aria-hidden="true" />
                  {option.label}
                </span>
              </label>
            );
          })}
        </div>
      </fieldset>

      <fieldset className="work-document-table-properties-section">
        <legend>
          {officeMessage(messages, 'document.table.properties.cellMargins')}
        </legend>
        <OfficeCheckbox
          ariaLabel={officeMessage(
            messages,
            'document.table.properties.useTableMarginsAria',
          )}
          checked={draft.cell.useTableMargins}
          onCheckedChange={(useTableMargins) =>
            setDraft((current) => ({
              ...current,
              cell: { ...current.cell, useTableMargins },
            }))
          }
        >
          {officeMessage(
            messages,
            'document.table.properties.useTableMargins',
          )}
        </OfficeCheckbox>
        <div className="work-document-table-properties-margin-grid">
          {marginFields(messages).map(({ side, label, ariaLabel }) => (
            <div key={side} className="work-document-table-properties-margin">
              <span>{label}</span>
              <OfficeNumberField
                ariaLabel={ariaLabel}
                value={draft.cell.margins[side]}
                min={0}
                max={5}
                step={0.05}
                disabled={draft.cell.useTableMargins}
                validationInvalid={Boolean(errors.cellMargins[side])}
                onValueChange={(value) =>
                  setDraft((current) => ({
                    ...current,
                    cell: {
                      ...current.cell,
                      margins: { ...current.cell.margins, [side]: value },
                    },
                  }))
                }
              />
              <small>{cm}</small>
              {errors.cellMargins[side] && (
                <p role="alert">{errors.cellMargins[side]}</p>
              )}
            </div>
          ))}
        </div>
      </fieldset>
    </>
  );
}

function tablePropertiesTabs(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'table' as const,
      label: officeMessage(messages, 'document.table.properties.tab.table'),
    },
    {
      value: 'row' as const,
      label: officeMessage(messages, 'document.table.properties.tab.row'),
    },
    {
      value: 'column' as const,
      label: officeMessage(messages, 'document.table.properties.tab.column'),
    },
    {
      value: 'cell' as const,
      label: officeMessage(messages, 'document.table.properties.tab.cell'),
    },
  ];
}

function widthOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'auto' as const satisfies DocumentTablePreferredWidthType,
      label: officeMessage(messages, 'document.table.properties.width.auto'),
    },
    {
      value: 'percent' as const,
      label: officeMessage(messages, 'document.table.properties.width.percent'),
    },
    {
      value: 'pixels' as const,
      label: officeMessage(messages, 'document.table.properties.width.cm'),
    },
  ];
}

function columnWidthOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'pixels' as const satisfies DocumentTableColumnWidthType,
      label: officeMessage(messages, 'document.table.properties.width.cm'),
    },
    {
      value: 'percent' as const,
      label: officeMessage(messages, 'document.table.properties.width.percent'),
    },
  ];
}

function alignmentOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'left' as const satisfies DocumentTableAlignment,
      label: officeMessage(messages, 'document.table.properties.align.left'),
      icon: AlignLeft,
    },
    {
      value: 'center' as const,
      label: officeMessage(messages, 'document.table.properties.align.center'),
      icon: AlignCenter,
    },
    {
      value: 'right' as const,
      label: officeMessage(messages, 'document.table.properties.align.right'),
      icon: AlignRight,
    },
  ];
}

function rowHeightRuleOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'atLeast',
      label: officeMessage(
        messages,
        'document.table.properties.rowHeight.atLeast',
      ),
    },
    {
      value: 'exact',
      label: officeMessage(
        messages,
        'document.table.properties.rowHeight.exact',
      ),
    },
  ] as const;
}

function verticalAlignmentOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'top' as const,
      label: officeMessage(messages, 'document.table.properties.align.top'),
      icon: AlignVerticalJustifyStart,
    },
    {
      value: 'middle' as const,
      label: officeMessage(messages, 'document.table.properties.align.middle'),
      icon: AlignVerticalJustifyCenter,
    },
    {
      value: 'bottom' as const,
      label: officeMessage(messages, 'document.table.properties.align.bottom'),
      icon: AlignVerticalJustifyEnd,
    },
  ];
}

function marginFields(messages: OfficeMessageCatalog): Array<{
  side: DocumentTableCellMarginSide;
  label: string;
  ariaLabel: string;
}> {
  return [
    {
      side: 'top',
      label: officeMessage(messages, 'document.table.properties.margin.top'),
      ariaLabel: officeMessage(
        messages,
        'document.table.properties.margin.topAria',
      ),
    },
    {
      side: 'bottom',
      label: officeMessage(messages, 'document.table.properties.margin.bottom'),
      ariaLabel: officeMessage(
        messages,
        'document.table.properties.margin.bottomAria',
      ),
    },
    {
      side: 'left',
      label: officeMessage(messages, 'document.table.properties.margin.left'),
      ariaLabel: officeMessage(
        messages,
        'document.table.properties.margin.leftAria',
      ),
    },
    {
      side: 'right',
      label: officeMessage(messages, 'document.table.properties.margin.right'),
      ariaLabel: officeMessage(
        messages,
        'document.table.properties.margin.rightAria',
      ),
    },
  ];
}

function NumberRow({
  label,
  ariaLabel,
  value,
  unit,
  min,
  max,
  step,
  disabled = false,
  invalid = false,
  onValueChange,
}: {
  label: string;
  ariaLabel: string;
  value: string;
  unit: string;
  min: number;
  max: number;
  step: number;
  disabled?: boolean;
  invalid?: boolean;
  onValueChange: (value: string) => void;
}) {
  return (
    <div className="work-document-table-properties-number-row">
      <span>{label}</span>
      <OfficeNumberField
        ariaLabel={ariaLabel}
        value={value}
        min={min}
        max={max}
        step={step}
        disabled={disabled}
        validationInvalid={invalid}
        onValueChange={onValueChange}
      />
      <small>{unit}</small>
    </div>
  );
}
