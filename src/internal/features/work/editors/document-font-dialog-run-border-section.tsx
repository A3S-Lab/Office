import type { CSSProperties } from 'react';
import { documentBorderPresentation } from '../work-document-paragraph-borders';
import type { DocumentRunBorderStyle } from '../work-document-run-border';
import {
  type DocumentFontDialogRunBorderDraft,
  type DocumentFontDialogRunBorderMode,
  type DocumentFontDialogRunBorderSource,
  documentFontDialogRunBorderFromDraft,
} from './document-font-dialog-run-border-model';
import {
  OfficeCheckbox,
  OfficeColorPicker,
  OfficeNumberField,
  OfficeSelect,
} from './office-controls';

const runBorderModes: ReadonlyArray<{
  value: DocumentFontDialogRunBorderMode;
  label: string;
  disabled?: boolean;
}> = [
  { value: 'mixed', label: '混合（保持不变）', disabled: true },
  { value: 'inherit', label: '跟随样式' },
  { value: 'none', label: '无（显式重置）' },
  { value: 'value', label: '边框' },
];

const runBorderStyles: ReadonlyArray<{
  value: DocumentRunBorderStyle;
  label: string;
}> = [
  { value: 'single', label: '单实线' },
  { value: 'thick', label: '粗实线' },
  { value: 'double', label: '双线' },
  { value: 'dotted', label: '点线' },
  { value: 'dashed', label: '虚线' },
  { value: 'dotDash', label: '点划线' },
  { value: 'dotDotDash', label: '双点划线' },
  { value: 'triple', label: '三线' },
  { value: 'thinThickSmallGap', label: '细粗线（小间距）' },
  { value: 'thickThinSmallGap', label: '粗细线（小间距）' },
  { value: 'thinThickThinSmallGap', label: '细粗细线（小间距）' },
  { value: 'thinThickMediumGap', label: '细粗线（中间距）' },
  { value: 'thickThinMediumGap', label: '粗细线（中间距）' },
  { value: 'thinThickThinMediumGap', label: '细粗细线（中间距）' },
  { value: 'thinThickLargeGap', label: '细粗线（大间距）' },
  { value: 'thickThinLargeGap', label: '粗细线（大间距）' },
  { value: 'thinThickThinLargeGap', label: '细粗细线（大间距）' },
  { value: 'wave', label: '波浪线' },
  { value: 'doubleWave', label: '双波浪线' },
  { value: 'dashSmallGap', label: '短划线' },
  { value: 'dashDotStroked', label: '粗点划线' },
  { value: 'threeDEmboss', label: '三维浮雕' },
  { value: 'threeDEngrave', label: '三维凹刻' },
  { value: 'outset', label: '外凸' },
  { value: 'inset', label: '内凹' },
];

export function DocumentFontDialogRunBorderSection({
  source,
  draft,
  touched,
  onDraftChange,
  onTouched,
}: {
  source: DocumentFontDialogRunBorderSource;
  draft: DocumentFontDialogRunBorderDraft;
  touched: boolean;
  onDraftChange: (patch: Partial<DocumentFontDialogRunBorderDraft>) => void;
  onTouched: () => void;
}) {
  const enabled = draft.runBorderMode === 'value';
  const update = (patch: Partial<DocumentFontDialogRunBorderDraft>) => {
    onDraftChange(patch);
    onTouched();
  };
  return (
    <fieldset
      className="work-document-font-dialog-run-border"
      aria-label="字符边框设置"
    >
      <legend>字符边框</legend>
      <div className="work-document-font-dialog-field">
        <span>应用方式</span>
        <OfficeSelect
          ariaLabel="字符边框"
          value={draft.runBorderMode}
          options={runBorderModes}
          onValueChange={(runBorderMode) => update({ runBorderMode })}
        />
      </div>
      <div className="work-document-font-dialog-field">
        <span>线型</span>
        <OfficeSelect
          ariaLabel="字符边框线型"
          value={draft.runBorderStyle}
          options={runBorderStyles}
          disabled={!enabled}
          onValueChange={(runBorderStyle) =>
            update({ runBorderMode: 'value', runBorderStyle })
          }
        />
      </div>
      <div className="work-document-font-dialog-field">
        <span>颜色</span>
        <OfficeColorPicker
          ariaLabel="字符边框颜色"
          value={
            draft.runBorderColor === 'auto' ? '#000000' : draft.runBorderColor
          }
          disabled={!enabled}
          resetAction={{
            kind: 'automatic',
            label: '自动颜色',
            onSelect: () =>
              update({ runBorderMode: 'value', runBorderColor: 'auto' }),
          }}
          onValueChange={(runBorderColor) =>
            update({
              runBorderMode: 'value',
              runBorderColor: runBorderColor as `#${string}`,
            })
          }
        />
      </div>
      <div className="work-document-font-dialog-field">
        <span>宽度</span>
        <span className="work-document-font-dialog-measure">
          <OfficeNumberField
            ariaLabel="字符边框宽度（磅）"
            value={draft.runBorderWidthPoints}
            min={0.25}
            max={12}
            step={0.125}
            disabled={!enabled}
            onValueChange={(runBorderWidthPoints) =>
              update({ runBorderMode: 'value', runBorderWidthPoints })
            }
          />
          <span aria-hidden="true">磅</span>
        </span>
      </div>
      <div className="work-document-font-dialog-field">
        <span>文字间距</span>
        <span className="work-document-font-dialog-measure">
          <OfficeNumberField
            ariaLabel="字符边框间距（磅）"
            value={draft.runBorderSpacingPoints}
            min={0}
            max={31}
            step={1}
            disabled={!enabled}
            onValueChange={(runBorderSpacingPoints) =>
              update({ runBorderMode: 'value', runBorderSpacingPoints })
            }
          />
          <span aria-hidden="true">磅</span>
        </span>
      </div>
      <div className="work-document-font-dialog-run-border-effects">
        <OfficeCheckbox
          ariaLabel="字符边框阴影"
          checked={draft.runBorderShadow}
          disabled={!enabled}
          onCheckedChange={(runBorderShadow) =>
            update({ runBorderMode: 'value', runBorderShadow })
          }
        >
          阴影
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel="字符边框框架"
          checked={draft.runBorderFrame}
          disabled={!enabled}
          onCheckedChange={(runBorderFrame) =>
            update({ runBorderMode: 'value', runBorderFrame })
          }
        >
          框架
        </OfficeCheckbox>
      </div>
      {source.mixed && !touched ? (
        <p className="work-document-font-dialog-mixed" role="status">
          当前选区包含不同的字符边框。选择一种应用方式后才会统一修改。
        </p>
      ) : null}
    </fieldset>
  );
}

export function documentFontDialogRunBorderPreviewStyle(
  source: DocumentFontDialogRunBorderSource,
  draft: DocumentFontDialogRunBorderDraft,
): CSSProperties {
  const border = documentFontDialogRunBorderFromDraft(source, draft);
  if (!border) return {};
  const presentation = documentBorderPresentation(border);
  return {
    border: `${presentation.width}px ${presentation.style} ${presentation.color}`,
    padding: `${(border.space ?? 0) * (96 / 72)}px`,
    boxDecorationBreak: 'clone',
    WebkitBoxDecorationBreak: 'clone',
    ...(border.shadow && presentation.width > 0
      ? { boxShadow: `2px 2px 0 ${presentation.color}` }
      : {}),
  };
}
