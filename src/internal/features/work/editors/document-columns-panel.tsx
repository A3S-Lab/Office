import {
  normalizeDocumentColumns,
  setCustomDocumentColumns,
  updateDocumentColumnWidth,
} from '../work-document-columns';
import type { WorkDocumentColumns } from '../work-types';
import { CommittedOfficeNumberField, OfficeCheckbox } from './office-controls';

export function DocumentColumnsPanel({
  columns,
  onChange,
}: {
  columns: WorkDocumentColumns;
  onChange: (columns: WorkDocumentColumns) => void;
}) {
  const normalized = normalizeDocumentColumns(columns);
  const custom = normalized.custom;
  const updateCustomSpacing = (index: number, spacing: number) => {
    if (!custom) return;
    onChange(
      normalizeDocumentColumns({
        ...normalized,
        custom: custom.map((column, columnIndex) =>
          columnIndex === index ? { ...column, spacing } : column,
        ),
      }),
    );
  };

  return (
    <fieldset className="work-document-columns-panel">
      <legend>分栏</legend>
      <div className="work-office-field">
        <span>栏数</span>
        <CommittedOfficeNumberField
          min={1}
          max={6}
          step={1}
          ariaLabel="分栏数量"
          value={normalized.count}
          normalizeValue={(value) => normalizeRequiredInteger(value, 1, 6)}
          onValueCommit={(count) =>
            onChange(
              normalizeDocumentColumns({
                ...normalized,
                count,
              }),
            )
          }
        />
      </div>
      {!normalized.custom && (
        <div className="work-office-field">
          <span>间距</span>
          <CommittedOfficeNumberField
            min={0}
            max={30}
            step={0.5}
            ariaLabel="分栏间距"
            value={normalized.spacing}
            normalizeValue={(value) => normalizeRequiredDecimal(value, 0, 30)}
            onValueCommit={(spacing) =>
              onChange(
                normalizeDocumentColumns({
                  ...normalized,
                  spacing,
                }),
              )
            }
          />
        </div>
      )}
      <OfficeCheckbox
        className="work-document-column-option"
        ariaLabel="自定义栏宽"
        disabled={normalized.count < 2}
        checked={Boolean(normalized.custom)}
        onCheckedChange={(checked) =>
          onChange(setCustomDocumentColumns(normalized, checked))
        }
      >
        自定义栏宽
      </OfficeCheckbox>
      <OfficeCheckbox
        className="work-document-column-option"
        ariaLabel="分栏分隔线"
        checked={normalized.separator}
        onCheckedChange={(checked) =>
          onChange({ ...normalized, separator: checked })
        }
      >
        分隔线
      </OfficeCheckbox>
      {custom && (
        <div className="work-document-custom-columns">
          {custom.map((column, index) => (
            <div
              key={`document-column-${index + 1}`}
              className="work-document-custom-column"
            >
              <strong>第 {index + 1} 栏</strong>
              <div className="work-office-field">
                <span>宽度 %</span>
                <CommittedOfficeNumberField
                  min={5}
                  max={100 - (normalized.count - 1) * 5}
                  step={0.5}
                  ariaLabel={`第 ${index + 1} 栏宽度百分比`}
                  value={column.widthPercent}
                  normalizeValue={(value) =>
                    normalizeRequiredDecimal(
                      value,
                      5,
                      100 - (normalized.count - 1) * 5,
                    )
                  }
                  onValueCommit={(widthPercent) =>
                    onChange(
                      updateDocumentColumnWidth(
                        normalized,
                        index,
                        widthPercent,
                      ),
                    )
                  }
                />
              </div>
              {index < custom.length - 1 && (
                <div className="work-office-field">
                  <span>栏后间距</span>
                  <CommittedOfficeNumberField
                    min={0}
                    max={30}
                    step={0.5}
                    ariaLabel={`第 ${index + 1} 栏后间距`}
                    value={column.spacing}
                    normalizeValue={(value) =>
                      normalizeRequiredDecimal(value, 0, 30)
                    }
                    onValueCommit={(spacing) =>
                      updateCustomSpacing(index, spacing)
                    }
                  />
                </div>
              )}
            </div>
          ))}
          <p>
            编辑区保持连续多栏输入；预览、PDF 与 DOCX
            使用这里的栏宽比例和独立间距。
          </p>
        </div>
      )}
    </fieldset>
  );
}

function normalizeRequiredInteger(
  value: string,
  minimum: number,
  maximum: number,
): number | null {
  if (!value.trim()) return null;
  const number = Number(value);
  return Number.isFinite(number)
    ? Math.min(maximum, Math.max(minimum, Math.round(number)))
    : null;
}

function normalizeRequiredDecimal(
  value: string,
  minimum: number,
  maximum: number,
): number | null {
  if (!value.trim()) return null;
  const number = Number(value);
  return Number.isFinite(number)
    ? Math.min(maximum, Math.max(minimum, Math.round(number * 10) / 10))
    : null;
}
