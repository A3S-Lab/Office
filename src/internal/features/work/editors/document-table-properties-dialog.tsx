import type { Editor } from '@tiptap/core';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  TableProperties,
} from 'lucide-react';
import {
  type FormEvent,
  type MouseEvent,
  useId,
  useRef,
  useState,
} from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import type {
  DocumentTableAlignment,
  DocumentTablePreferredWidthType,
  DocumentTableProperties as DocumentTablePropertiesValue,
} from '../work-document-table-geometry';
import {
  documentTableSizing,
  type DocumentTableSizingState,
} from '../work-document-table-sizing';
import { OfficeNumberField } from './office-controls';
import { WorkOfficeRibbonButton } from './work-office-chrome';

const PIXELS_PER_CENTIMETER = 96 / 2.54;
const FALLBACK_TABLE_WIDTH_CENTIMETERS = 15;

interface DocumentTablePropertiesDraft {
  widthType: DocumentTablePreferredWidthType;
  width: string;
  alignment: DocumentTableAlignment;
  indent: string;
}

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

export function DocumentTablePropertiesControl({
  editor,
  renderedTableWidth,
}: {
  editor: Editor;
  renderedTableWidth?: number;
}) {
  const [open, setOpen] = useState(false);
  const [draft, setDraft] = useState<DocumentTablePropertiesDraft>(() =>
    propertiesDraft(documentTableSizing(editor.state), renderedTableWidth),
  );
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  const formId = useId();
  const widthError = tableWidthError(draft);
  const indentError = tableIndentError(draft);
  const invalid = Boolean(widthError || indentError);

  const close = () => setOpen(false);
  const submit = (event?: FormEvent<HTMLFormElement>) => {
    event?.preventDefault();
    const properties = normalizedProperties(draft);
    if (!properties) return;
    if (editor.commands.setDocumentTableProperties(properties)) close();
  };
  const openDialog = (event: MouseEvent<HTMLButtonElement>) => {
    const sizing = documentTableSizing(editor.state);
    if (!sizing) return;
    triggerRef.current = event.currentTarget;
    setDraft(propertiesDraft(sizing, renderedTableWidth));
    setOpen(true);
  };

  return (
    <>
      <WorkOfficeRibbonButton
        label="表格属性"
        visibleLabel="表格属性"
        disabled={!documentTableSizing(editor.state)}
        onClick={openDialog}
      >
        <TableProperties size={18} />
      </WorkOfficeRibbonButton>
      {open && (
        <Dialog
          title="表格属性"
          description="设置整张表格的宽度和位置。"
          className="work-document-table-properties-dialog"
          restoreFocusTarget={() => triggerRef.current}
          onClose={close}
          footer={
            <>
              <Button tone="quiet" onClick={close}>
                取消
              </Button>
              <Button
                tone="primary"
                type="submit"
                form={formId}
                disabled={invalid}
              >
                确定
              </Button>
            </>
          }
        >
          <form id={formId} onSubmit={submit}>
            <fieldset className="work-document-table-properties-section">
              <legend>首选宽度</legend>
              <div className="work-document-table-properties-choice-grid width">
                {widthOptions.map((option) => (
                  <label key={option.value}>
                    <input
                      type="radio"
                      name={`${formId}-width`}
                      value={option.value}
                      checked={draft.widthType === option.value}
                      data-autofocus={
                        draft.widthType === option.value || undefined
                      }
                      onChange={() =>
                        setDraft((current) =>
                          draftForWidthType(
                            current,
                            option.value,
                            renderedTableWidth,
                          ),
                        )
                      }
                    />
                    <span>{option.label}</span>
                  </label>
                ))}
              </div>
              {draft.widthType !== 'auto' && (
                <div className="work-document-table-properties-number-row">
                  <span>宽度</span>
                  <OfficeNumberField
                    ariaLabel={
                      draft.widthType === 'percent'
                        ? '表格宽度（百分比）'
                        : '表格宽度（厘米）'
                    }
                    value={draft.width}
                    min={draft.widthType === 'percent' ? 1 : 0.5}
                    max={draft.widthType === 'percent' ? 100 : 30}
                    step={draft.widthType === 'percent' ? 1 : 0.1}
                    validationInvalid={Boolean(widthError)}
                    onValueChange={(width) =>
                      setDraft((current) => ({ ...current, width }))
                    }
                  />
                  <small>{draft.widthType === 'percent' ? '%' : '厘米'}</small>
                </div>
              )}
              {widthError && <p role="alert">{widthError}</p>}
            </fieldset>

            <fieldset className="work-document-table-properties-section">
              <legend>表格位置</legend>
              <div className="work-document-table-properties-choice-grid alignment">
                {alignmentOptions.map((option) => {
                  const Icon = option.icon;
                  return (
                    <label key={option.value}>
                      <input
                        type="radio"
                        name={`${formId}-alignment`}
                        value={option.value}
                        checked={draft.alignment === option.value}
                        onChange={() =>
                          setDraft((current) => ({
                            ...current,
                            alignment: option.value,
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
              <div className="work-document-table-properties-number-row">
                <span>左缩进</span>
                <OfficeNumberField
                  ariaLabel="表格左缩进（厘米）"
                  value={draft.indent}
                  min={0}
                  max={30}
                  step={0.1}
                  disabled={draft.alignment !== 'left'}
                  validationInvalid={Boolean(indentError)}
                  onValueChange={(indent) =>
                    setDraft((current) => ({ ...current, indent }))
                  }
                />
                <small>厘米</small>
              </div>
              {indentError && <p role="alert">{indentError}</p>}
            </fieldset>
          </form>
        </Dialog>
      )}
    </>
  );
}

function propertiesDraft(
  sizing: DocumentTableSizingState | null,
  renderedTableWidth: number | undefined,
): DocumentTablePropertiesDraft {
  const widthType = sizing?.preferredWidthType ?? 'auto';
  return {
    widthType,
    width:
      widthType === 'percent'
        ? formatNumber(sizing?.preferredWidth ?? 100)
        : widthType === 'pixels'
          ? centimeters(sizing?.preferredWidth ?? renderedTableWidth ?? 0)
          : '',
    alignment: sizing?.alignment ?? 'left',
    indent: centimeters(sizing?.indent ?? 0),
  };
}

function draftForWidthType(
  current: DocumentTablePropertiesDraft,
  widthType: DocumentTablePreferredWidthType,
  renderedTableWidth: number | undefined,
): DocumentTablePropertiesDraft {
  if (widthType === current.widthType) return current;
  if (widthType === 'auto') return { ...current, widthType, width: '' };
  if (widthType === 'percent') {
    return { ...current, widthType, width: '100' };
  }
  return {
    ...current,
    widthType,
    width: centimeters(
      renderedTableWidth ??
        FALLBACK_TABLE_WIDTH_CENTIMETERS * PIXELS_PER_CENTIMETER,
    ),
  };
}

function normalizedProperties(
  draft: DocumentTablePropertiesDraft,
): DocumentTablePropertiesValue | null {
  if (tableWidthError(draft) || tableIndentError(draft)) return null;
  const rawIndent = Number(draft.indent);
  const indent =
    (Number.isFinite(rawIndent) ? rawIndent : 0) * PIXELS_PER_CENTIMETER;
  const width =
    draft.widthType === 'auto'
      ? ({ type: 'auto', value: null } as const)
      : draft.widthType === 'percent'
        ? ({ type: 'percent', value: Number(draft.width) } as const)
        : ({
            type: 'pixels',
            value: Number(draft.width) * PIXELS_PER_CENTIMETER,
          } as const);
  return {
    width: {
      ...width,
      value: width.value === null ? null : Math.round(width.value * 100) / 100,
    },
    alignment: draft.alignment,
    indent: Math.round(indent * 100) / 100,
  };
}

function tableWidthError(draft: DocumentTablePropertiesDraft): string | null {
  if (draft.widthType === 'auto') return null;
  if (!draft.width.trim()) {
    return draft.widthType === 'percent'
      ? '请输入 1 到 100 之间的百分比。'
      : '请输入 0.5 到 30 之间的厘米数。';
  }
  const value = Number(draft.width);
  if (draft.widthType === 'percent') {
    return Number.isFinite(value) && value >= 1 && value <= 100
      ? null
      : '请输入 1 到 100 之间的百分比。';
  }
  return Number.isFinite(value) && value >= 0.5 && value <= 30
    ? null
    : '请输入 0.5 到 30 之间的厘米数。';
}

function tableIndentError(draft: DocumentTablePropertiesDraft): string | null {
  if (draft.alignment !== 'left') return null;
  const value = draft.indent;
  if (!value.trim()) return '请输入 0 到 30 之间的厘米数。';
  const numeric = Number(value);
  return Number.isFinite(numeric) && numeric >= 0 && numeric <= 30
    ? null
    : '请输入 0 到 30 之间的厘米数。';
}

function centimeters(pixels: number): string {
  return formatNumber(pixels / PIXELS_PER_CENTIMETER);
}

function formatNumber(value: number): string {
  return Number.isFinite(value) ? value.toFixed(2).replace(/\.?0+$/, '') : '';
}
