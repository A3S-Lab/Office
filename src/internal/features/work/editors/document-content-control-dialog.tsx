import { Button, Dialog } from '../../../design-system/primitives';
import type {
  WorkDocumentContentControlAppearance,
  WorkDocumentContentControlListItem,
  WorkDocumentContentControlLock,
  WorkDocumentContentControlProperties,
  WorkDocumentContentControlType,
} from '../work-document-content-control';
import {
  contentControlDateInputValue,
  contentControlFullDateFromInput,
  DOCUMENT_CONTENT_CONTROL_DATE_FORMATS,
  DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_FORMAT,
  DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_LANGUAGE,
  isListStyleContentControl,
} from '../work-document-content-control';
import {
  OfficeCheckbox,
  OfficeSelect,
  OfficeTextField,
} from './office-controls';

export interface DocumentContentControlDialogProps {
  editing: boolean;
  properties: WorkDocumentContentControlProperties;
  restoreFocusTarget: () => HTMLElement | null;
  onCancel: () => void;
  onChange: (properties: WorkDocumentContentControlProperties) => void;
  onSubmit: () => void;
}

const typeOptions = [
  { value: 'text', label: '纯文本' },
  { value: 'richText', label: '富文本' },
  { value: 'checkbox', label: '复选框' },
  { value: 'dropDownList', label: '下拉列表' },
  { value: 'comboBox', label: '组合框' },
  { value: 'date', label: '日期' },
] as const satisfies readonly {
  value: WorkDocumentContentControlType;
  label: string;
}[];

const lockOptions = [
  { value: 'unlocked', label: '可编辑' },
  { value: 'contentLocked', label: '锁定内容' },
  { value: 'sdtLocked', label: '锁定控件' },
  { value: 'sdtContentLocked', label: '锁定控件和内容' },
] as const satisfies readonly {
  value: WorkDocumentContentControlLock;
  label: string;
}[];

const appearanceOptions = [
  { value: 'boundingBox', label: '边框' },
  { value: 'tags', label: '标签' },
  { value: 'hidden', label: '隐藏边框' },
] as const satisfies readonly {
  value: WorkDocumentContentControlAppearance;
  label: string;
}[];

const dateFormatOptions = DOCUMENT_CONTENT_CONTROL_DATE_FORMATS.map(
  (value) => ({ value, label: value }),
);

export function DocumentContentControlDialog({
  editing,
  properties,
  restoreFocusTarget,
  onCancel,
  onChange,
  onSubmit,
}: DocumentContentControlDialogProps) {
  const update = (patch: Partial<WorkDocumentContentControlProperties>) =>
    onChange({ ...properties, ...patch });
  const aliasId = 'document-content-control-alias';
  const tagId = 'document-content-control-tag';
  const optionsId = 'document-content-control-options';
  const dateId = 'document-content-control-date';
  const isCheckbox = properties.type === 'checkbox';
  const isList = isListStyleContentControl(properties.type);
  const isComboBox = properties.type === 'comboBox';
  const isDate = properties.type === 'date';
  const optionsText = listItemsToLines(properties.options);
  const canSubmit =
    (!isList || properties.options.length > 0) &&
    (!isDate || Boolean(properties.fullDate));
  const dateFormatSelectOptions = dateFormatOptions.some(
    (item) => item.value === properties.dateFormat,
  )
    ? dateFormatOptions
    : [
        { value: properties.dateFormat, label: properties.dateFormat },
        ...dateFormatOptions,
      ];

  return (
    <Dialog
      title={editing ? '编辑内容控件' : '插入内容控件'}
      description="支持内联纯文本、富文本、有界复选框、下拉列表、组合框和日期；绑定与重复区域不会被伪装成普通文本。"
      className="work-document-content-control-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onCancel}
      footer={
        <>
          <Button tone="quiet" onClick={onCancel}>
            取消
          </Button>
          <Button tone="primary" disabled={!canSubmit} onClick={onSubmit}>
            {editing ? '应用' : '插入控件'}
          </Button>
        </>
      }
    >
      <div className="work-document-content-control-dialog-fields">
        <label htmlFor={aliasId}>
          <span>显示名称</span>
          <OfficeTextField
            id={aliasId}
            data-autofocus
            aria-label="内容控件显示名称"
            value={properties.alias}
            maxLength={255}
            placeholder="例如：客户名称"
            onChange={(event) => update({ alias: event.target.value })}
            onKeyDown={(event) => {
              if (event.key !== 'Enter' || event.nativeEvent.isComposing)
                return;
              event.preventDefault();
              if (canSubmit) onSubmit();
            }}
          />
        </label>
        <label htmlFor={tagId}>
          <span>程序标签</span>
          <OfficeTextField
            id={tagId}
            aria-label="内容控件程序标签"
            value={properties.tag}
            maxLength={255}
            placeholder="可选，供文档自动化识别"
            onChange={(event) => update({ tag: event.target.value })}
          />
        </label>
      </div>
      <div className="work-document-content-control-dialog-selects">
        <div className="work-document-content-control-dialog-field">
          <span>控件类型</span>
          <OfficeSelect<WorkDocumentContentControlType>
            ariaLabel="内容控件类型"
            value={properties.type}
            options={typeOptions}
            onValueChange={(type) =>
              update({
                type,
                multiLine:
                  type === 'checkbox' ||
                  type === 'date' ||
                  isListStyleContentControl(type)
                    ? false
                    : properties.multiLine,
                checked: type === 'checkbox' ? properties.checked : false,
                options: isListStyleContentControl(type)
                  ? properties.options.length
                    ? properties.options
                    : [{ displayText: '选项1', value: '选项1' }]
                  : [],
                selectedValue: isListStyleContentControl(type)
                  ? type === 'comboBox'
                    ? properties.selectedValue &&
                      properties.options.some(
                        (item) => item.value === properties.selectedValue,
                      )
                      ? properties.selectedValue
                      : properties.options[0]?.value || '选项1'
                    : properties.selectedValue ||
                      properties.options[0]?.value ||
                      '选项1'
                  : '',
                fullDate:
                  type === 'date'
                    ? properties.fullDate || contentControlFullDateFromInput('')
                    : '',
                dateFormat:
                  type === 'date'
                    ? properties.dateFormat ||
                      DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_FORMAT
                    : DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_FORMAT,
                dateLanguage:
                  type === 'date'
                    ? properties.dateLanguage ||
                      DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_LANGUAGE
                    : DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_LANGUAGE,
                dateMapping:
                  type === 'date' ? properties.dateMapping : 'dateTime',
              })
            }
          />
        </div>
        <div className="work-document-content-control-dialog-field">
          <span>锁定方式</span>
          <OfficeSelect<WorkDocumentContentControlLock>
            ariaLabel="内容控件锁定方式"
            value={properties.lock}
            options={lockOptions}
            onValueChange={(lock) => update({ lock })}
          />
        </div>
        <div className="work-document-content-control-dialog-field">
          <span>外观</span>
          <OfficeSelect<WorkDocumentContentControlAppearance>
            ariaLabel="内容控件外观"
            value={properties.appearance}
            options={appearanceOptions}
            onValueChange={(appearance) => update({ appearance })}
          />
        </div>
      </div>
      {isList ? (
        <div className="work-document-content-control-dialog-fields">
          <label htmlFor={optionsId}>
            <span>列表选项（每行一项）</span>
            <textarea
              id={optionsId}
              aria-label={isComboBox ? '组合框选项' : '下拉列表选项'}
              className="work-document-content-control-options"
              rows={4}
              value={optionsText}
              placeholder={'选项1\n选项2\n显示文字|值'}
              onChange={(event) => {
                const options = linesToListItems(event.target.value);
                update({
                  options,
                  selectedValue: options.some(
                    (item) => item.value === properties.selectedValue,
                  )
                    ? properties.selectedValue
                    : (options[0]?.value ?? ''),
                });
              }}
            />
          </label>
          <div className="work-document-content-control-dialog-field">
            <span>{isComboBox ? '建议选项' : '当前选项'}</span>
            <OfficeSelect<string>
              ariaLabel={isComboBox ? '组合框建议选项' : '下拉列表当前选项'}
              value={
                properties.selectedValue || properties.options[0]?.value || ''
              }
              options={properties.options.map((item) => ({
                value: item.value,
                label: item.displayText,
              }))}
              disabled={!properties.options.length}
              onValueChange={(selectedValue) => update({ selectedValue })}
            />
          </div>
        </div>
      ) : null}
      {isDate ? (
        <div className="work-document-content-control-dialog-fields">
          <label htmlFor={dateId}>
            <span>日期</span>
            <input
              id={dateId}
              type="date"
              aria-label="内容控件日期"
              className="work-document-content-control-date"
              value={contentControlDateInputValue(properties.fullDate)}
              onChange={(event) =>
                update({
                  fullDate: contentControlFullDateFromInput(event.target.value),
                })
              }
            />
          </label>
          <div className="work-document-content-control-dialog-field">
            <span>显示格式</span>
            <OfficeSelect<string>
              ariaLabel="日期显示格式"
              value={properties.dateFormat}
              options={dateFormatSelectOptions}
              onValueChange={(dateFormat) => update({ dateFormat })}
            />
          </div>
        </div>
      ) : null}
      <div className="work-document-content-control-dialog-options">
        {isCheckbox ? (
          <OfficeCheckbox
            ariaLabel="复选框已勾选"
            checked={properties.checked}
            onCheckedChange={(checked) => update({ checked })}
          >
            默认勾选
          </OfficeCheckbox>
        ) : isList || isDate ? null : (
          <OfficeCheckbox
            ariaLabel="允许多行文字"
            checked={properties.multiLine}
            onCheckedChange={(multiLine) => update({ multiLine })}
          >
            允许多行文字
          </OfficeCheckbox>
        )}
        <label className="work-document-content-control-color">
          <span>控件颜色</span>
          <input
            type="color"
            aria-label="内容控件颜色"
            value={properties.color ?? '#2f6fed'}
            onChange={(event) => update({ color: event.target.value })}
          />
          <button
            type="button"
            className="work-document-content-control-color-reset"
            onClick={() => update({ color: null })}
          >
            使用主题色
          </button>
        </label>
      </div>
    </Dialog>
  );
}

function listItemsToLines(
  options: WorkDocumentContentControlListItem[],
): string {
  return options
    .map((item) =>
      item.displayText === item.value
        ? item.displayText
        : `${item.displayText}|${item.value}`,
    )
    .join('\n');
}

function linesToListItems(
  source: string,
): WorkDocumentContentControlListItem[] {
  const items: WorkDocumentContentControlListItem[] = [];
  const seen = new Set<string>();
  for (const line of source.split(/\r?\n/u)) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    const separator = trimmed.indexOf('|');
    const displayText =
      separator >= 0 ? trimmed.slice(0, separator).trim() : trimmed;
    const value =
      separator >= 0
        ? trimmed.slice(separator + 1).trim() || displayText
        : displayText;
    if (!displayText || !value || seen.has(value) || items.length >= 32)
      continue;
    seen.add(value);
    items.push({ displayText, value });
  }
  return items;
}
