import {
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  AlignCenter,
  AlignLeft,
  AlignRight,
} from 'lucide-react';
import type { Dispatch, SetStateAction } from 'react';
import type {
  DocumentTableAlignment,
  DocumentTableCellMarginSide,
  DocumentTablePreferredWidthType,
} from '../work-document-table-geometry';
import type { DocumentTablePropertiesSource } from './document-table-properties-dialog-model';
import {
  draftForTableWidthType,
  type DocumentTablePropertiesDraft,
  type DocumentTablePropertiesErrors,
  type DocumentTablePropertiesTab,
} from './document-table-properties-dialog-model';
import {
  OfficeCheckbox,
  OfficeNumberField,
  OfficeSelect,
} from './office-controls';

const tabs = [
  { value: 'table', label: '表格' },
  { value: 'row', label: '行' },
  { value: 'column', label: '列' },
  { value: 'cell', label: '单元格' },
] as const satisfies readonly {
  value: DocumentTablePropertiesTab;
  label: string;
}[];

const widthOptions = [
  { value: 'auto', label: '自动' },
  { value: 'percent', label: '百分比' },
  { value: 'pixels', label: '厘米' },
] as const satisfies readonly {
  value: DocumentTablePreferredWidthType;
  label: string;
}[];

const alignmentOptions = [
  { value: 'left', label: '左对齐', icon: AlignLeft },
  { value: 'center', label: '居中', icon: AlignCenter },
  { value: 'right', label: '右对齐', icon: AlignRight },
] as const satisfies readonly {
  value: DocumentTableAlignment;
  label: string;
  icon: typeof AlignLeft;
}[];

const rowHeightRuleOptions = [
  { value: 'atLeast', label: '最小值' },
  { value: 'exact', label: '固定值' },
] as const;

const verticalAlignmentOptions = [
  {
    value: 'top',
    label: '顶端',
    icon: AlignVerticalJustifyStart,
  },
  {
    value: 'middle',
    label: '居中',
    icon: AlignVerticalJustifyCenter,
  },
  {
    value: 'bottom',
    label: '底端',
    icon: AlignVerticalJustifyEnd,
  },
] as const;

const marginFields = [
  { side: 'top', label: '上', ariaLabel: '当前单元格上边距（厘米）' },
  { side: 'bottom', label: '下', ariaLabel: '当前单元格下边距（厘米）' },
  { side: 'left', label: '左', ariaLabel: '当前单元格左边距（厘米）' },
  { side: 'right', label: '右', ariaLabel: '当前单元格右边距（厘米）' },
] as const satisfies readonly {
  side: DocumentTableCellMarginSide;
  label: string;
  ariaLabel: string;
}[];

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
  return (
    <div
      className="work-document-table-properties-tabs"
      role="tablist"
      aria-label="表格属性分类"
    >
      {tabs.map((tab) => (
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
  return (
    <>
      <fieldset className="work-document-table-properties-section">
        <legend>首选宽度</legend>
        <div className="work-document-table-properties-choice-grid width">
          {widthOptions.map((option) => (
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
            label="宽度"
            ariaLabel={
              draft.table.widthType === 'percent'
                ? '表格宽度（百分比）'
                : '表格宽度（厘米）'
            }
            value={draft.table.width}
            unit={draft.table.widthType === 'percent' ? '%' : '厘米'}
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
        <legend>位置</legend>
        <div className="work-document-table-properties-choice-grid alignment">
          {alignmentOptions.map((option) => {
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
          label="左缩进"
          ariaLabel="表格左缩进（厘米）"
          value={draft.table.indent}
          unit="厘米"
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
  return (
    <>
      <fieldset className="work-document-table-properties-section">
        <legend>当前行尺寸</legend>
        <OfficeCheckbox
          ariaLabel="指定行高"
          checked={draft.row.heightEnabled}
          onCheckedChange={(heightEnabled) =>
            setDraft((current) => ({
              ...current,
              row: { ...current.row, heightEnabled },
            }))
          }
        >
          指定行高
        </OfficeCheckbox>
        <NumberRow
          label="行高"
          ariaLabel="当前行高（厘米）"
          value={draft.row.height}
          unit="厘米"
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
          <span>行高规则</span>
          <OfficeSelect
            ariaLabel="行高规则"
            value={draft.row.heightRule}
            options={rowHeightRuleOptions}
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
        <legend>分页</legend>
        <div className="work-document-table-properties-checkboxes">
          <OfficeCheckbox
            ariaLabel="允许跨页断行"
            checked={!draft.row.cantSplit}
            onCheckedChange={(allowSplit) =>
              setDraft((current) => ({
                ...current,
                row: { ...current.row, cantSplit: !allowSplit },
              }))
            }
          >
            允许跨页断行
          </OfficeCheckbox>
          <OfficeCheckbox
            ariaLabel="在各页顶端重复标题行"
            checked={draft.row.repeatHeader}
            disabled={!source.canRepeatHeader}
            onCheckedChange={(repeatHeader) =>
              setDraft((current) => ({
                ...current,
                row: { ...current.row, repeatHeader },
              }))
            }
          >
            在各页顶端重复标题行
          </OfficeCheckbox>
        </div>
      </fieldset>
    </>
  );
}

function ColumnSection({ draft, setDraft, errors }: SectionProps) {
  return (
    <fieldset className="work-document-table-properties-section">
      <legend>当前列尺寸</legend>
      <NumberRow
        label="列宽"
        ariaLabel="当前列宽（厘米）"
        value={draft.column.width}
        unit="厘米"
        min={0.5}
        max={30}
        step={0.1}
        invalid={Boolean(errors.columnWidth)}
        onValueChange={(width) =>
          setDraft((current) => ({
            ...current,
            column: { width },
          }))
        }
      />
      {errors.columnWidth && <p role="alert">{errors.columnWidth}</p>}
    </fieldset>
  );
}

function CellSection({ draft, setDraft, errors }: SectionProps) {
  return (
    <>
      <fieldset className="work-document-table-properties-section">
        <legend>垂直对齐</legend>
        <div className="work-document-table-properties-choice-grid alignment">
          {verticalAlignmentOptions.map((option) => {
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
        <legend>单元格边距</legend>
        <OfficeCheckbox
          ariaLabel="使用表格默认边距"
          checked={draft.cell.useTableMargins}
          onCheckedChange={(useTableMargins) =>
            setDraft((current) => ({
              ...current,
              cell: { ...current.cell, useTableMargins },
            }))
          }
        >
          使用表格默认边距
        </OfficeCheckbox>
        <div className="work-document-table-properties-margin-grid">
          {marginFields.map(({ side, label, ariaLabel }) => (
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
              <small>厘米</small>
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
