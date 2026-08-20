import { Globe2, Grid2X2, Link2, Table2 } from 'lucide-react';
import { type FormEvent, useId, useMemo, useState } from 'react';
import { Button, Dialog, Field } from '../../../design-system/primitives';
import type {
  SpreadsheetHyperlinkDialogSource,
  SpreadsheetHyperlinkDialogValue,
  SpreadsheetHyperlinkType,
} from './spreadsheet-hyperlink';

const hyperlinkTypes: readonly {
  id: SpreadsheetHyperlinkType;
  label: string;
  description: string;
  icon: typeof Globe2;
}[] = [
  {
    id: 'webpage',
    label: '网页',
    description: 'HTTP 或 HTTPS 地址',
    icon: Globe2,
  },
  {
    id: 'cellrange',
    label: '单元格区域',
    description: '当前或其他工作表中的连续区域',
    icon: Grid2X2,
  },
  {
    id: 'sheet',
    label: '工作表',
    description: '直接打开目标工作表',
    icon: Table2,
  },
];

export function SpreadsheetHyperlinkDialog({
  source,
  restoreFocusTarget,
  onApply,
  onClose,
  onRemove,
  onValidate,
}: {
  source: SpreadsheetHyperlinkDialogSource;
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (value: SpreadsheetHyperlinkDialogValue) => boolean;
  onClose: () => void;
  onRemove: () => boolean;
  onValidate: (value: SpreadsheetHyperlinkDialogValue) => string | null;
}) {
  const [linkType, setLinkType] = useState<SpreadsheetHyperlinkType>(
    source.link?.linkType ?? 'webpage',
  );
  const [linkAddress, setLinkAddress] = useState(
    source.link?.linkAddress ?? '',
  );
  const [displayText, setDisplayText] = useState(source.displayText);
  const [addressTouched, setAddressTouched] = useState(false);
  const formId = useId();
  const typeId = useId().replaceAll(':', '');
  const value = useMemo(
    () =>
      spreadsheetHyperlinkDialogValue(
        source,
        linkType,
        linkAddress,
        displayText,
      ),
    [displayText, linkAddress, linkType, source],
  );
  const validationError = onValidate(value);
  const addressMissing = !linkAddress.trim();
  const visibleError =
    addressTouched && !addressMissing ? validationError : undefined;

  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setAddressTouched(true);
    if (addressMissing || validationError) return;
    if (onApply(value)) onClose();
  };

  const changeType = (nextType: SpreadsheetHyperlinkType) => {
    setLinkType(nextType);
    setLinkAddress(defaultSpreadsheetHyperlinkAddress(source, nextType));
    setAddressTouched(false);
  };

  return (
    <Dialog
      title={source.hasHyperlink ? '编辑超链接' : '插入超链接'}
      description={`${source.sheetName}!${source.cellReference}`}
      className="work-spreadsheet-hyperlink-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <>
          {source.hasHyperlink && (
            <Button
              tone="danger"
              className="work-spreadsheet-hyperlink-remove"
              onClick={() => {
                if (onRemove()) onClose();
              }}
            >
              移除超链接
            </Button>
          )}
          <Button tone="quiet" onClick={onClose}>
            取消
          </Button>
          <Button
            tone="primary"
            type="submit"
            form={formId}
            disabled={addressMissing || Boolean(validationError)}
          >
            确定
          </Button>
        </>
      }
    >
      <form id={formId} onSubmit={submit}>
        <fieldset
          className="work-spreadsheet-hyperlink-types"
          aria-label="链接类型"
        >
          <legend>链接类型</legend>
          <div role="radiogroup" aria-label="链接类型">
            {hyperlinkTypes.map((item) => {
              const Icon = item.icon;
              const checked = linkType === item.id;
              const inputId = `${typeId}-${item.id}`;
              return (
                <label key={item.id} className={checked ? 'selected' : ''}>
                  <input
                    id={inputId}
                    type="radio"
                    aria-label={item.label}
                    name={`${typeId}-link-type`}
                    value={item.id}
                    checked={checked}
                    onChange={() => changeType(item.id)}
                  />
                  <Icon size={17} aria-hidden="true" />
                  <span>
                    <strong>{item.label}</strong>
                    <small>{item.description}</small>
                  </span>
                </label>
              );
            })}
          </div>
        </fieldset>

        <Field
          label="显示文本"
          description={
            source.displayTextEditable
              ? '留空时使用链接目标作为显示文本。'
              : '公式单元格的显示文本由公式结果决定。'
          }
        >
          <input
            type="text"
            value={displayText}
            disabled={!source.displayTextEditable}
            maxLength={32_767}
            onChange={(event) => setDisplayText(event.currentTarget.value)}
          />
        </Field>

        {linkType === 'sheet' ? (
          <Field label="工作表" required error={visibleError}>
            <select
              value={linkAddress}
              onBlur={() => setAddressTouched(true)}
              onChange={(event) => {
                setLinkAddress(event.currentTarget.value);
                setAddressTouched(true);
              }}
            >
              {source.sheetOptions.map((sheet) => (
                <option key={sheet.id} value={sheet.name}>
                  {sheet.name}
                </option>
              ))}
            </select>
          </Field>
        ) : (
          <Field
            label={linkType === 'webpage' ? '地址' : '单元格或区域'}
            required
            error={visibleError}
            description={
              linkType === 'webpage'
                ? '例如 https://a3s.dev/office'
                : "例如 A1、B2:D8 或 'Archive 2025'!C9:E12"
            }
          >
            <input
              type="text"
              inputMode={linkType === 'webpage' ? 'url' : 'text'}
              autoCapitalize="none"
              spellCheck={false}
              value={linkAddress}
              onBlur={() => setAddressTouched(true)}
              onChange={(event) => {
                setLinkAddress(event.currentTarget.value);
                setAddressTouched(true);
              }}
            />
          </Field>
        )}

        <div className="work-spreadsheet-hyperlink-preview" aria-live="polite">
          <span aria-hidden="true">
            <Link2 size={17} />
          </span>
          <div>
            <strong>
              {(value.displayText ?? source.displayText) || '超链接'}
            </strong>
            <small>{linkAddress.trim() || '等待输入链接目标'}</small>
          </div>
        </div>
      </form>
    </Dialog>
  );
}

function spreadsheetHyperlinkDialogValue(
  source: SpreadsheetHyperlinkDialogSource,
  linkType: SpreadsheetHyperlinkType,
  linkAddress: string,
  displayText: string,
): SpreadsheetHyperlinkDialogValue {
  const normalizedDisplay = displayText.trim() || linkAddress.trim();
  const shouldChangeDisplay =
    source.displayTextEditable && normalizedDisplay !== source.displayText;
  return {
    linkType,
    linkAddress,
    ...(shouldChangeDisplay ? { displayText: normalizedDisplay } : {}),
  };
}

function defaultSpreadsheetHyperlinkAddress(
  source: SpreadsheetHyperlinkDialogSource,
  linkType: SpreadsheetHyperlinkType,
): string {
  if (source.link?.linkType === linkType) return source.link.linkAddress;
  if (linkType !== 'sheet') return '';
  return (
    source.sheetOptions.find((sheet) => sheet.id !== source.sheetId)?.name ??
    source.sheetOptions[0]?.name ??
    ''
  );
}
