import type { DocumentRunShadingPattern } from '../work-document-run-shading';
import type {
  DocumentFontDialogRunShadingDraft,
  DocumentFontDialogRunShadingMode,
  DocumentFontDialogRunShadingSource,
} from './document-font-dialog-run-shading-model';
import { OfficeColorPicker, OfficeSelect } from './office-controls';

const shadingModes: ReadonlyArray<{
  value: DocumentFontDialogRunShadingMode;
  label: string;
  disabled?: boolean;
}> = [
  { value: 'mixed', label: '混合（保持不变）', disabled: true },
  { value: 'inherit', label: '跟随样式' },
  { value: 'none', label: '无（显式重置）' },
  { value: 'value', label: '底纹' },
];

const shadingPatterns: ReadonlyArray<{
  value: Exclude<DocumentRunShadingPattern, 'nil'>;
  label: string;
  group?: string;
}> = [
  { value: 'clear', label: '清除（背景色）', group: '基本' },
  { value: 'solid', label: '实心（前景色）', group: '基本' },
  { value: 'horzStripe', label: '水平条纹', group: '条纹' },
  { value: 'vertStripe', label: '垂直条纹', group: '条纹' },
  { value: 'reverseDiagStripe', label: '反向对角条纹', group: '条纹' },
  { value: 'diagStripe', label: '对角条纹', group: '条纹' },
  { value: 'horzCross', label: '水平交叉', group: '交叉' },
  { value: 'diagCross', label: '对角交叉', group: '交叉' },
  { value: 'thinHorzStripe', label: '细水平条纹', group: '细线' },
  { value: 'thinVertStripe', label: '细垂直条纹', group: '细线' },
  { value: 'thinReverseDiagStripe', label: '细反向对角条纹', group: '细线' },
  { value: 'thinDiagStripe', label: '细对角条纹', group: '细线' },
  { value: 'thinHorzCross', label: '细水平交叉', group: '细线' },
  { value: 'thinDiagCross', label: '细对角交叉', group: '细线' },
  ...[
    5, 10, 12, 15, 20, 25, 30, 35, 37, 40, 45, 50, 55, 60, 62, 65, 70, 75, 80,
    85, 87, 90, 95,
  ].map((percentage) => ({
    value: `pct${percentage}` as Exclude<DocumentRunShadingPattern, 'nil'>,
    label: `${percentage}%`,
    group: '密度',
  })),
];

export function DocumentFontDialogRunShadingSection({
  source,
  draft,
  touched,
  onDraftChange,
  onTouched,
}: {
  source: DocumentFontDialogRunShadingSource;
  draft: DocumentFontDialogRunShadingDraft;
  touched: boolean;
  onDraftChange: (patch: Partial<DocumentFontDialogRunShadingDraft>) => void;
  onTouched: () => void;
}) {
  const enabled = draft.runShadingMode === 'value';
  const update = (patch: Partial<DocumentFontDialogRunShadingDraft>) => {
    onDraftChange(patch);
    onTouched();
  };
  return (
    <fieldset
      className="work-document-font-dialog-run-shading"
      aria-label="字符底纹设置"
    >
      <legend>字符底纹</legend>
      <div className="work-document-font-dialog-field">
        <span>应用方式</span>
        <OfficeSelect
          ariaLabel="字符底纹"
          value={draft.runShadingMode}
          options={shadingModes}
          onValueChange={(runShadingMode) => update({ runShadingMode })}
        />
      </div>
      <div className="work-document-font-dialog-field">
        <span>图案</span>
        <OfficeSelect
          ariaLabel="字符底纹图案"
          value={draft.runShadingPattern}
          options={shadingPatterns}
          disabled={!enabled}
          onValueChange={(runShadingPattern) =>
            update({ runShadingMode: 'value', runShadingPattern })
          }
        />
      </div>
      <div className="work-document-font-dialog-field">
        <span>前景色</span>
        <OfficeColorPicker
          ariaLabel="字符底纹前景色"
          value={
            draft.runShadingColor === 'auto' ? '#000000' : draft.runShadingColor
          }
          disabled={!enabled}
          resetAction={{
            kind: 'automatic',
            label: '自动前景色',
            onSelect: () =>
              update({ runShadingMode: 'value', runShadingColor: 'auto' }),
          }}
          onValueChange={(runShadingColor) =>
            update({
              runShadingMode: 'value',
              runShadingColor: runShadingColor as `#${string}`,
            })
          }
        />
      </div>
      <div className="work-document-font-dialog-field">
        <span>背景色</span>
        <OfficeColorPicker
          ariaLabel="字符底纹背景色"
          value={
            draft.runShadingFill === 'auto' ? '#ffffff' : draft.runShadingFill
          }
          disabled={!enabled}
          resetAction={{
            kind: 'automatic',
            label: '自动背景色',
            onSelect: () =>
              update({ runShadingMode: 'value', runShadingFill: 'auto' }),
          }}
          onValueChange={(runShadingFill) =>
            update({
              runShadingMode: 'value',
              runShadingFill: runShadingFill as `#${string}`,
            })
          }
        />
      </div>
      {source.mixed && !touched ? (
        <p className="work-document-font-dialog-mixed" role="status">
          当前选区包含不同的字符底纹。选择一种应用方式后才会统一修改。
        </p>
      ) : null}
    </fieldset>
  );
}
