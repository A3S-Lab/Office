import type { Editor } from '@tiptap/core';
import { NodeSelection } from '@tiptap/pm/state';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Rows3,
  Settings2,
  TextWrap,
} from 'lucide-react';
import {
  type FormEvent,
  type MouseEvent,
  useId,
  useRef,
  useState,
} from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import {
  documentImageProperties,
  type WorkDocumentImageAlignment,
  type WorkDocumentImageHorizontalReference,
  type WorkDocumentImageLayout,
  type WorkDocumentImageVerticalReference,
} from '../work-document-image-layout';
import {
  createDocumentPicturePropertiesDraft,
  type DocumentPicturePropertiesDraft,
  type DocumentPicturePropertiesErrors,
  type DocumentPicturePropertiesSource,
  documentPicturePropertiesErrors,
  documentPicturePropertyChanges,
  hasDocumentPicturePropertiesErrors,
  withDocumentPictureAspectRatioLock,
  withDocumentPictureDimension,
} from './document-picture-properties-dialog-model';
import {
  OfficeCheckbox,
  OfficeNumberField,
  OfficeSelect,
  OfficeTextArea,
} from './office-controls';
import { WorkOfficeRibbonButton } from './work-office-chrome';

interface PictureDialogSource extends DocumentPicturePropertiesSource {
  position: number;
}

const layoutOptions = [
  { value: 'inline', label: '嵌入文字', icon: Rows3 },
  { value: 'square', label: '四周环绕', icon: TextWrap },
  { value: 'topBottom', label: '上下环绕', icon: Rows3 },
] as const satisfies readonly {
  value: WorkDocumentImageLayout;
  label: string;
  icon: typeof Rows3;
}[];

const alignmentOptions = [
  { value: 'left', label: '左对齐', icon: AlignLeft },
  { value: 'center', label: '居中', icon: AlignCenter },
  { value: 'right', label: '右对齐', icon: AlignRight },
] as const satisfies readonly {
  value: WorkDocumentImageAlignment;
  label: string;
  icon: typeof AlignLeft;
}[];

const horizontalReferenceOptions = [
  { value: 'column', label: '栏' },
  { value: 'margin', label: '页边距' },
  { value: 'page', label: '页面' },
] as const satisfies readonly {
  value: WorkDocumentImageHorizontalReference;
  label: string;
}[];

const verticalReferenceOptions = [
  { value: 'paragraph', label: '段落' },
  { value: 'margin', label: '页边距' },
  { value: 'page', label: '页面' },
] as const satisfies readonly {
  value: WorkDocumentImageVerticalReference;
  label: string;
}[];

export function DocumentPicturePropertiesControl({
  editor,
}: {
  editor: Editor;
}) {
  const [source, setSource] = useState<PictureDialogSource | null>(null);
  const triggerRef = useRef<HTMLButtonElement | null>(null);

  const openDialog = (event: MouseEvent<HTMLButtonElement>) => {
    const position = selectedDocumentImagePosition(editor);
    if (position === null) return;
    const rendered = selectedDocumentImageDimensions(editor, position);
    triggerRef.current = event.currentTarget;
    setSource({
      position,
      properties: documentImageProperties(editor),
      renderedWidth: rendered?.width,
      renderedHeight: rendered?.height,
    });
  };

  return (
    <>
      <WorkOfficeRibbonButton
        label="图片属性"
        visibleLabel="图片属性"
        disabled={!editor.isActive('image')}
        onClick={openDialog}
      >
        <Settings2 size={18} />
      </WorkOfficeRibbonButton>
      {source && (
        <DocumentPicturePropertiesDialog
          editor={editor}
          source={source}
          restoreFocusTarget={() => triggerRef.current}
          onClose={() => setSource(null)}
        />
      )}
    </>
  );
}

function DocumentPicturePropertiesDialog({
  editor,
  source,
  restoreFocusTarget,
  onClose,
}: {
  editor: Editor;
  source: PictureDialogSource;
  restoreFocusTarget: () => HTMLElement | null;
  onClose: () => void;
}) {
  const [initial] = useState(() =>
    createDocumentPicturePropertiesDraft(source),
  );
  const [draft, setDraft] = useState(initial);
  const formId = useId();
  const alternativeTextId = useId();
  const errors = documentPicturePropertiesErrors(draft);
  const invalid = hasDocumentPicturePropertiesErrors(errors);

  const submit = (event?: FormEvent<HTMLFormElement>) => {
    event?.preventDefault();
    if (invalid) return;
    const changes = documentPicturePropertyChanges(initial, draft);
    if (!changes) {
      onClose();
      return;
    }
    if (editor.state.doc.nodeAt(source.position)?.type.name !== 'image') {
      onClose();
      return;
    }
    const applied = editor
      .chain()
      .setNodeSelection(source.position)
      .setDocumentImageProperties(changes, { restoreFocus: false })
      .run();
    if (applied) onClose();
  };

  return (
    <Dialog
      title="图片属性"
      description="调整当前图片的大小、排列方式和辅助说明。"
      className="work-document-picture-properties-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <>
          <Button tone="quiet" onClick={onClose}>
            取消
          </Button>
          <Button tone="primary" type="submit" form={formId} disabled={invalid}>
            确定
          </Button>
        </>
      }
    >
      <form id={formId} onSubmit={submit}>
        <PictureSizeSection
          draft={draft}
          errors={errors}
          onDraftChange={setDraft}
        />
        <PictureLayoutSection draft={draft} onDraftChange={setDraft} />
        <PictureAlignmentSection
          draft={draft}
          errors={errors}
          onDraftChange={setDraft}
        />
        <PictureCropSection
          draft={draft}
          errors={errors}
          onDraftChange={setDraft}
        />
        <label
          className="work-document-picture-properties-alt-text"
          htmlFor={alternativeTextId}
        >
          <span>替代文字</span>
          <OfficeTextArea
            id={alternativeTextId}
            aria-label="图片替代文字"
            value={draft.alternativeText}
            maxLength={512}
            placeholder="描述图片中的关键信息；装饰性图片可留空"
            onChange={(event) => {
              const alternativeText = event.currentTarget.value;
              setDraft((current) => ({ ...current, alternativeText }));
            }}
          />
          <small>供屏幕阅读器使用，不会显示在正文中。</small>
        </label>
      </form>
    </Dialog>
  );
}

function PictureSizeSection({
  draft,
  errors,
  onDraftChange,
}: {
  draft: DocumentPicturePropertiesDraft;
  errors: DocumentPicturePropertiesErrors;
  onDraftChange: React.Dispatch<
    React.SetStateAction<DocumentPicturePropertiesDraft>
  >;
}) {
  return (
    <fieldset className="work-document-picture-properties-section size">
      <legend>大小</legend>
      <PictureNumberRow
        label="宽度"
        ariaLabel="图片宽度（厘米）"
        value={draft.width}
        unit="厘米"
        min={0.01}
        max={55.87}
        step={0.1}
        invalid={Boolean(errors.width)}
        onValueChange={(width) =>
          onDraftChange((current) =>
            withDocumentPictureDimension(current, 'width', width),
          )
        }
      />
      {errors.width && <p role="alert">{errors.width}</p>}
      <PictureNumberRow
        label="高度"
        ariaLabel="图片高度（厘米）"
        value={draft.height}
        unit="厘米"
        min={0.01}
        max={55.87}
        step={0.1}
        invalid={Boolean(errors.height)}
        onValueChange={(height) =>
          onDraftChange((current) =>
            withDocumentPictureDimension(current, 'height', height),
          )
        }
      />
      {errors.height && <p role="alert">{errors.height}</p>}
      <OfficeCheckbox
        ariaLabel="锁定纵横比"
        checked={draft.lockAspectRatio}
        onCheckedChange={(locked) =>
          onDraftChange((current) =>
            withDocumentPictureAspectRatioLock(current, locked),
          )
        }
      >
        锁定纵横比
      </OfficeCheckbox>
    </fieldset>
  );
}

function PictureLayoutSection({
  draft,
  onDraftChange,
}: {
  draft: DocumentPicturePropertiesDraft;
  onDraftChange: React.Dispatch<
    React.SetStateAction<DocumentPicturePropertiesDraft>
  >;
}) {
  const errors = documentPicturePropertiesErrors(draft);
  return (
    <fieldset className="work-document-picture-properties-section layout">
      <legend>文字环绕</legend>
      <div className="work-document-picture-properties-choice-grid">
        {layoutOptions.map((option) => {
          const Icon = option.icon;
          return (
            <label key={option.value}>
              <input
                type="radio"
                name="picture-properties-layout"
                value={option.value}
                checked={draft.layout === option.value}
                onChange={() =>
                  onDraftChange((current) => ({
                    ...current,
                    layout: option.value,
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
      <PictureNumberRow
        label="文字距离"
        ariaLabel="图片与文字距离（毫米）"
        value={draft.wrapDistance}
        unit="毫米"
        min={0}
        max={25}
        step={0.5}
        disabled={draft.layout === 'inline'}
        invalid={Boolean(errors.wrapDistance)}
        onValueChange={(wrapDistance) =>
          onDraftChange((current) => ({ ...current, wrapDistance }))
        }
      />
      {errors.wrapDistance && <p role="alert">{errors.wrapDistance}</p>}
    </fieldset>
  );
}

function PictureAlignmentSection({
  draft,
  errors,
  onDraftChange,
}: {
  draft: DocumentPicturePropertiesDraft;
  errors: DocumentPicturePropertiesErrors;
  onDraftChange: React.Dispatch<
    React.SetStateAction<DocumentPicturePropertiesDraft>
  >;
}) {
  return (
    <fieldset className="work-document-picture-properties-section position">
      <legend>位置</legend>
      <div className="work-document-picture-properties-choice-grid">
        {alignmentOptions.map((option) => {
          const Icon = option.icon;
          return (
            <label key={option.value}>
              <input
                type="radio"
                name="picture-properties-alignment"
                value={option.value}
                checked={draft.alignment === option.value}
                onChange={() =>
                  onDraftChange((current) => ({
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
      <OfficeCheckbox
        ariaLabel="使用精确图片位置"
        checked={draft.precisePosition}
        disabled={draft.layout === 'inline'}
        onCheckedChange={(precisePosition) =>
          onDraftChange((current) => ({ ...current, precisePosition }))
        }
      >
        使用精确位置
      </OfficeCheckbox>
      <div className="work-document-picture-properties-position-grid">
        <PictureNumberRow
          label="水平偏移"
          ariaLabel="图片水平偏移（毫米）"
          value={draft.horizontalOffset}
          unit="毫米"
          min={-558.7}
          max={558.7}
          step={0.5}
          disabled={draft.layout === 'inline' || !draft.precisePosition}
          invalid={Boolean(errors.horizontalOffset)}
          onValueChange={(horizontalOffset) =>
            onDraftChange((current) => ({ ...current, horizontalOffset }))
          }
        />
        <OfficeSelect<WorkDocumentImageHorizontalReference>
          ariaLabel="水平相对于"
          value={draft.horizontalReference}
          options={horizontalReferenceOptions}
          disabled={draft.layout === 'inline' || !draft.precisePosition}
          onValueChange={(horizontalReference) =>
            onDraftChange((current) => ({ ...current, horizontalReference }))
          }
        />
        <PictureNumberRow
          label="垂直偏移"
          ariaLabel="图片垂直偏移（毫米）"
          value={draft.verticalOffset}
          unit="毫米"
          min={-558.7}
          max={558.7}
          step={0.5}
          disabled={draft.layout === 'inline' || !draft.precisePosition}
          invalid={Boolean(errors.verticalOffset)}
          onValueChange={(verticalOffset) =>
            onDraftChange((current) => ({ ...current, verticalOffset }))
          }
        />
        <OfficeSelect<WorkDocumentImageVerticalReference>
          ariaLabel="垂直相对于"
          value={draft.verticalReference}
          options={verticalReferenceOptions}
          disabled={draft.layout === 'inline' || !draft.precisePosition}
          onValueChange={(verticalReference) =>
            onDraftChange((current) => ({ ...current, verticalReference }))
          }
        />
      </div>
      {errors.horizontalOffset && <p role="alert">{errors.horizontalOffset}</p>}
      {errors.verticalOffset && <p role="alert">{errors.verticalOffset}</p>}
    </fieldset>
  );
}

function PictureCropSection({
  draft,
  errors,
  onDraftChange,
}: {
  draft: DocumentPicturePropertiesDraft;
  errors: DocumentPicturePropertiesErrors;
  onDraftChange: React.Dispatch<
    React.SetStateAction<DocumentPicturePropertiesDraft>
  >;
}) {
  const fields = [
    { key: 'cropTop', label: '上方裁剪' },
    { key: 'cropRight', label: '右侧裁剪' },
    { key: 'cropBottom', label: '下方裁剪' },
    { key: 'cropLeft', label: '左侧裁剪' },
  ] as const;
  return (
    <fieldset className="work-document-picture-properties-section crop">
      <legend>裁剪</legend>
      <div className="work-document-picture-properties-crop-grid">
        {fields.map((field) => (
          <PictureNumberRow
            key={field.key}
            label={field.label}
            ariaLabel={`图片${field.label}（百分比）`}
            value={draft[field.key]}
            unit="%"
            min={0}
            max={99.99}
            step={1}
            invalid={Boolean(errors.crop)}
            onValueChange={(value) =>
              onDraftChange((current) => ({
                ...current,
                [field.key]: value,
              }))
            }
          />
        ))}
      </div>
      {errors.crop && <p role="alert">{errors.crop}</p>}
    </fieldset>
  );
}

function PictureNumberRow({
  label,
  ariaLabel,
  value,
  unit,
  min,
  max,
  step,
  disabled = false,
  invalid,
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
  invalid: boolean;
  onValueChange: (value: string) => void;
}) {
  return (
    <div className="work-document-picture-properties-number-row">
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

function selectedDocumentImagePosition(editor: Editor): number | null {
  const selection = editor.state.selection;
  return selection instanceof NodeSelection &&
    selection.node.type.name === 'image'
    ? selection.from
    : null;
}

function selectedDocumentImageDimensions(
  editor: Editor,
  position: number,
): { width: number; height: number } | null {
  const node = editor.view.nodeDOM(position);
  const image =
    node instanceof HTMLImageElement
      ? node
      : node instanceof HTMLElement
        ? node.querySelector('img')
        : null;
  if (!(image instanceof HTMLImageElement)) return null;
  const bounds = image.getBoundingClientRect();
  const width = bounds.width || image.width || image.naturalWidth;
  const height = bounds.height || image.height || image.naturalHeight;
  return width > 0 && height > 0 ? { width, height } : null;
}
