import { Button, Dialog } from '../../../design-system/primitives';
import type {
  WorkDocumentClockFieldFormat,
  WorkDocumentFieldDraft,
  WorkDocumentFieldKind,
  WorkDocumentNumericFieldFormat,
} from '../work-document-fields';
import {
  OfficeCheckbox,
  OfficeSelect,
  type OfficeSelectOption,
} from './office-controls';

export interface DocumentFieldTargetOption {
  id: string;
  name: string;
}

export interface DocumentFieldDialogProps {
  editing: boolean;
  draft: WorkDocumentFieldDraft;
  preview: string;
  targets: readonly DocumentFieldTargetOption[];
  restoreFocusTarget: () => HTMLElement | null;
  onCancel: () => void;
  onChange: (draft: WorkDocumentFieldDraft) => void;
  onSubmit: () => void;
}

export function DocumentFieldDialog({
  editing,
  draft,
  preview,
  targets,
  restoreFocusTarget,
  onCancel,
  onChange,
  onSubmit,
}: DocumentFieldDialogProps) {
  const numeric = isNumericField(draft.kind);
  const clockKind =
    draft.kind === 'date' ? 'date' : draft.kind === 'time' ? 'time' : null;
  const clock = clockKind !== null;
  const hasTarget =
    draft.kind !== 'pageReference' ||
    targets.some(
      (target) =>
        Boolean(draft.targetName) &&
        target.id === draft.targetId &&
        target.name === draft.targetName,
    );
  const targetValue = `${draft.targetId}:${draft.targetName}`;

  return (
    <Dialog
      title={editing ? '编辑字段' : '插入字段'}
      description="设置页码、日期、统计和书签引用；插入后可使用 F9 更新结果。"
      className="work-document-field-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onCancel}
      footer={
        <>
          <Button tone="quiet" onClick={onCancel}>
            取消
          </Button>
          <Button
            tone="primary"
            disabled={draft.kind === 'pageReference' && !hasTarget}
            onClick={onSubmit}
          >
            {editing ? '应用字段' : '插入字段'}
          </Button>
        </>
      }
    >
      <div className="work-document-field-dialog-grid">
        <div className="work-document-dialog-field">
          <span>字段类型</span>
          <OfficeSelect
            initialFocus
            ariaLabel="字段类型"
            value={draft.kind}
            options={fieldKindOptions}
            onValueChange={(value) => {
              if (!value) return;
              const kind = value as WorkDocumentFieldKind;
              onChange({
                ...draft,
                kind,
                format: defaultFormat(kind),
                targetId: kind === 'pageReference' ? draft.targetId : '',
                targetName: kind === 'pageReference' ? draft.targetName : '',
                hyperlink: kind === 'pageReference' ? draft.hyperlink : false,
              });
            }}
          />
        </div>

        {numeric && (
          <div className="work-document-dialog-field">
            <span>数字格式</span>
            <OfficeSelect
              ariaLabel="数字格式"
              value={
                draft.format.kind === 'numeric' ? draft.format.value : 'arabic'
              }
              options={numericFormatOptions}
              onValueChange={(value) => {
                if (!value) return;
                onChange({
                  ...draft,
                  format: {
                    kind: 'numeric',
                    value: value as WorkDocumentNumericFieldFormat,
                  },
                });
              }}
            />
          </div>
        )}

        {clock && (
          <div className="work-document-dialog-field">
            <span>{draft.kind === 'date' ? '日期格式' : '时间格式'}</span>
            <OfficeSelect
              ariaLabel={clockKind === 'date' ? '日期格式' : '时间格式'}
              value={
                draft.format.kind === 'clock'
                  ? draft.format.source
                    ? '__preserved__'
                    : draft.format.value
                  : defaultClockFormat(clockKind ?? 'date')
              }
              options={clockFormatOptions(clockKind ?? 'date', draft)}
              onValueChange={(value) => {
                if (!value || value === '__preserved__') return;
                onChange({
                  ...draft,
                  format: {
                    kind: 'clock',
                    value: value as WorkDocumentClockFieldFormat,
                  },
                });
              }}
            />
          </div>
        )}

        {draft.kind === 'pageReference' && (
          <div className="work-document-dialog-field">
            <span>引用目标</span>
            <OfficeSelect
              ariaLabel="引用目标"
              value={targetValue}
              options={targets.map((target) => ({
                value: `${target.id}:${target.name}`,
                label: target.name,
              }))}
              disabled={!targets.length}
              onValueChange={(value) => {
                const target = targets.find(
                  (candidate) => `${candidate.id}:${candidate.name}` === value,
                );
                if (!target) return;
                onChange({
                  ...draft,
                  targetId: target.id,
                  targetName: target.name,
                });
              }}
            />
            {!hasTarget && (
              <small className="work-document-field-dialog-help">
                {targets.length
                  ? '请选择一个书签作为引用目标。'
                  : '请先插入一个书签，再插入目标页码字段。'}
              </small>
            )}
          </div>
        )}
      </div>

      {draft.kind === 'pageReference' && (
        <OfficeCheckbox
          ariaLabel="使用超链接"
          checked={draft.hyperlink}
          onCheckedChange={(hyperlink) => onChange({ ...draft, hyperlink })}
        >
          使用超链接
        </OfficeCheckbox>
      )}

      <div className="work-document-field-dialog-preview">
        <span>结果预览</span>
        <output aria-label="结果预览" aria-live="polite">
          {preview || '—'}
        </output>
      </div>
      <p className="work-document-field-dialog-note">
        {draft.mergeFormat
          ? '此字段保留 WPS 的 MERGEFORMAT 开关。'
          : '应用后仍可使用 F9 更新分页、日期和统计结果。'}
      </p>
    </Dialog>
  );
}

const fieldKindOptions = [
  { value: 'page', label: '页码' },
  { value: 'numPages', label: '总页数' },
  { value: 'section', label: '当前节号' },
  { value: 'sectionPages', label: '本节页数' },
  { value: 'date', label: '当前日期' },
  { value: 'time', label: '当前时间' },
  { value: 'wordCount', label: '字数' },
  { value: 'characterCount', label: '字符数' },
  { value: 'pageReference', label: '目标页码' },
] as const satisfies readonly OfficeSelectOption<WorkDocumentFieldKind>[];

const numericFormatOptions = [
  { value: 'arabic', label: '阿拉伯数字（1）' },
  { value: 'roman', label: '大写罗马数字（I）' },
  { value: 'romanLower', label: '小写罗马数字（i）' },
  { value: 'alphabetic', label: '大写字母（A）' },
  { value: 'alphabeticLower', label: '小写字母（a）' },
  { value: 'ordinal', label: '序数（1st）' },
] as const satisfies readonly OfficeSelectOption<WorkDocumentNumericFieldFormat>[];

const dateFormatOptions = [
  { value: 'yyyy年M月d日', label: '2026年9月6日' },
  { value: 'yyyy-MM-dd', label: '2026-09-06' },
  { value: 'MMMM d, yyyy', label: 'September 6, 2026' },
  { value: 'dddd, MMMM d, yyyy', label: 'Sunday, September 6, 2026' },
] as const satisfies readonly OfficeSelectOption<WorkDocumentClockFieldFormat>[];

const timeFormatOptions = [
  { value: 'HH:mm', label: '14:05' },
  { value: 'HH:mm:ss', label: '14:05:09' },
  { value: 'h:mm AM/PM', label: '2:05 PM' },
] as const satisfies readonly OfficeSelectOption<WorkDocumentClockFieldFormat>[];

function isNumericField(kind: WorkDocumentFieldKind): boolean {
  return (
    kind === 'page' ||
    kind === 'numPages' ||
    kind === 'section' ||
    kind === 'sectionPages' ||
    kind === 'pageReference'
  );
}

function defaultFormat(kind: WorkDocumentFieldKind) {
  if (isNumericField(kind))
    return { kind: 'numeric' as const, value: 'arabic' as const };
  if (kind === 'date') {
    return { kind: 'clock' as const, value: 'yyyy年M月d日' as const };
  }
  if (kind === 'time') {
    return { kind: 'clock' as const, value: 'HH:mm' as const };
  }
  return { kind: 'none' as const };
}

function defaultClockFormat(
  kind: 'date' | 'time',
): WorkDocumentClockFieldFormat {
  return kind === 'date' ? 'yyyy年M月d日' : 'HH:mm';
}

function clockFormatOptions(
  kind: 'date' | 'time',
  draft: WorkDocumentFieldDraft,
): readonly OfficeSelectOption<
  WorkDocumentClockFieldFormat | '__preserved__'
>[] {
  const options = kind === 'date' ? dateFormatOptions : timeFormatOptions;
  return draft.format.kind === 'clock' && draft.format.source
    ? [
        {
          value: '__preserved__' as const,
          label: `保留现有格式（${draft.format.source}）`,
        },
        ...options,
      ]
    : options;
}
