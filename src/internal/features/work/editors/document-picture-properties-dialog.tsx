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
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import {
  documentImageProperties,
  MAX_DOCUMENT_IMAGE_RELATIVE_HEIGHT,
  type WorkDocumentImageAlignment,
  type WorkDocumentImageHorizontalReference,
  type WorkDocumentImageLayout,
  type WorkDocumentImageRotation,
  type WorkDocumentImageVerticalReference,
  type WorkDocumentImageWrapSide,
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
import { useOfficeMessages } from './office-messages-context';
import { WorkOfficeRibbonButton } from './work-office-chrome';

interface PictureDialogSource extends DocumentPicturePropertiesSource {
  position: number;
}

function layoutOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'inline' as const,
      label: officeMessage(messages, 'document.picture.wrap.inline'),
      icon: Rows3,
    },
    {
      value: 'square' as const,
      label: officeMessage(messages, 'document.picture.wrap.square'),
      icon: TextWrap,
    },
    {
      value: 'tight' as const,
      label: officeMessage(messages, 'document.picture.wrap.tight'),
      icon: TextWrap,
    },
    {
      value: 'through' as const,
      label: officeMessage(messages, 'document.picture.wrap.through'),
      icon: TextWrap,
    },
    {
      value: 'topBottom' as const,
      label: officeMessage(messages, 'document.picture.wrap.topBottom'),
      icon: Rows3,
    },
    {
      value: 'none' as const,
      label: officeMessage(messages, 'document.picture.wrap.none'),
      icon: TextWrap,
    },
  ] as const satisfies readonly {
    value: WorkDocumentImageLayout;
    label: string;
    icon: typeof Rows3;
  }[];
}

function alignmentOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'left' as const,
      label: officeMessage(messages, 'document.picture.align.left'),
      icon: AlignLeft,
    },
    {
      value: 'center' as const,
      label: officeMessage(messages, 'document.picture.align.center'),
      icon: AlignCenter,
    },
    {
      value: 'right' as const,
      label: officeMessage(messages, 'document.picture.align.right'),
      icon: AlignRight,
    },
  ] as const satisfies readonly {
    value: WorkDocumentImageAlignment;
    label: string;
    icon: typeof AlignLeft;
  }[];
}

function wrapSideOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'bothSides' as const,
      label: officeMessage(messages, 'document.picture.wrapSide.bothSides'),
    },
    {
      value: 'left' as const,
      label: officeMessage(messages, 'document.picture.wrapSide.left'),
    },
    {
      value: 'right' as const,
      label: officeMessage(messages, 'document.picture.wrapSide.right'),
    },
    {
      value: 'largest' as const,
      label: officeMessage(messages, 'document.picture.wrapSide.largest'),
    },
  ] as const satisfies readonly {
    value: WorkDocumentImageWrapSide;
    label: string;
  }[];
}

function horizontalReferenceOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'column' as const,
      label: officeMessage(messages, 'document.picture.href.column'),
    },
    {
      value: 'margin' as const,
      label: officeMessage(messages, 'document.picture.href.margin'),
    },
    {
      value: 'page' as const,
      label: officeMessage(messages, 'document.picture.href.page'),
    },
  ] as const satisfies readonly {
    value: WorkDocumentImageHorizontalReference;
    label: string;
  }[];
}

function verticalReferenceOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'paragraph' as const,
      label: officeMessage(messages, 'document.picture.vref.paragraph'),
    },
    {
      value: 'margin' as const,
      label: officeMessage(messages, 'document.picture.href.margin'),
    },
    {
      value: 'page' as const,
      label: officeMessage(messages, 'document.picture.href.page'),
    },
  ] as const satisfies readonly {
    value: WorkDocumentImageVerticalReference;
    label: string;
  }[];
}

function rotationOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: '0',
      label: officeMessage(messages, 'document.picture.transform.rotation0'),
    },
    { value: '90', label: '90°' },
    { value: '180', label: '180°' },
    { value: '270', label: '270°' },
  ] as const;
}

export function DocumentPicturePropertiesControl({
  editor,
}: {
  editor: Editor;
}) {
  const messages = useOfficeMessages();
  const [source, setSource] = useState<PictureDialogSource | null>(null);
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  const title = officeMessage(messages, 'document.picture.title');

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
        label={title}
        visibleLabel={title}
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
  const messages = useOfficeMessages();
  const [initial] = useState(() =>
    createDocumentPicturePropertiesDraft(source),
  );
  const [draft, setDraft] = useState(initial);
  const formId = useId();
  const alternativeTextId = useId();
  const errors = documentPicturePropertiesErrors(draft, messages);
  const invalid = hasDocumentPicturePropertiesErrors(errors);

  const submit = (event?: FormEvent<HTMLFormElement>) => {
    event?.preventDefault();
    if (invalid) return;
    const changes = documentPicturePropertyChanges(initial, draft, messages);
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
      title={officeMessage(messages, 'document.picture.title')}
      description={officeMessage(messages, 'document.picture.description')}
      className="work-document-picture-properties-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <>
          <Button tone="quiet" onClick={onClose}>
            {officeMessage(messages, 'document.picture.cancel')}
          </Button>
          <Button tone="primary" type="submit" form={formId} disabled={invalid}>
            {officeMessage(messages, 'document.picture.confirm')}
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
        <PictureLayerSection
          draft={draft}
          errors={errors}
          onDraftChange={setDraft}
        />
        <PictureCropSection
          draft={draft}
          errors={errors}
          onDraftChange={setDraft}
        />
        <PictureTransformSection draft={draft} onDraftChange={setDraft} />
        <label
          className="work-document-picture-properties-alt-text"
          htmlFor={alternativeTextId}
        >
          <span>{officeMessage(messages, 'document.picture.alt.label')}</span>
          <OfficeTextArea
            id={alternativeTextId}
            aria-label={officeMessage(messages, 'document.picture.alt.aria')}
            value={draft.alternativeText}
            maxLength={512}
            placeholder={officeMessage(
              messages,
              'document.picture.alt.placeholder',
            )}
            onChange={(event) => {
              const alternativeText = event.currentTarget.value;
              setDraft((current) => ({ ...current, alternativeText }));
            }}
          />
          <small>{officeMessage(messages, 'document.picture.alt.help')}</small>
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
  const messages = useOfficeMessages();
  const cm = officeMessage(messages, 'document.picture.size.cm');
  return (
    <fieldset className="work-document-picture-properties-section size">
      <legend>{officeMessage(messages, 'document.picture.size.legend')}</legend>
      <PictureNumberRow
        label={officeMessage(messages, 'document.picture.size.width')}
        ariaLabel={officeMessage(messages, 'document.picture.size.widthAria')}
        value={draft.width}
        unit={cm}
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
        label={officeMessage(messages, 'document.picture.size.height')}
        ariaLabel={officeMessage(messages, 'document.picture.size.heightAria')}
        value={draft.height}
        unit={cm}
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
        ariaLabel={officeMessage(messages, 'document.picture.size.lockAspect')}
        checked={draft.lockAspectRatio}
        onCheckedChange={(locked) =>
          onDraftChange((current) =>
            withDocumentPictureAspectRatioLock(current, locked),
          )
        }
      >
        {officeMessage(messages, 'document.picture.size.lockAspect')}
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
  const messages = useOfficeMessages();
  const errors = documentPicturePropertiesErrors(draft, messages);
  const mm = officeMessage(messages, 'document.picture.wrap.mm');
  return (
    <fieldset className="work-document-picture-properties-section layout">
      <legend>{officeMessage(messages, 'document.picture.wrap.legend')}</legend>
      <div className="work-document-picture-properties-choice-grid">
        {layoutOptions(messages).map((option) => {
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
      <div className="work-document-picture-properties-wrap-side">
        <span>{officeMessage(messages, 'document.picture.wrap.side')}</span>
        <OfficeSelect<WorkDocumentImageWrapSide>
          ariaLabel={officeMessage(messages, 'document.picture.wrap.sideAria')}
          value={draft.wrapSide}
          options={[...wrapSideOptions(messages)]}
          disabled={
            draft.layout === 'inline' ||
            draft.layout === 'topBottom' ||
            draft.layout === 'none'
          }
          onValueChange={(wrapSide) =>
            onDraftChange((current) => ({ ...current, wrapSide }))
          }
        />
      </div>
      <PictureNumberRow
        label={officeMessage(messages, 'document.picture.wrap.distance')}
        ariaLabel={officeMessage(
          messages,
          'document.picture.wrap.distanceAria',
        )}
        value={draft.wrapDistance}
        unit={mm}
        min={0}
        max={25}
        step={0.5}
        disabled={draft.layout === 'inline' || draft.layout === 'none'}
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
  const messages = useOfficeMessages();
  const mm = officeMessage(messages, 'document.picture.wrap.mm');
  return (
    <fieldset className="work-document-picture-properties-section position">
      <legend>
        {officeMessage(messages, 'document.picture.position.legend')}
      </legend>
      <div className="work-document-picture-properties-choice-grid">
        {alignmentOptions(messages).map((option) => {
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
        ariaLabel={officeMessage(
          messages,
          'document.picture.position.preciseAria',
        )}
        checked={draft.precisePosition}
        disabled={draft.layout === 'inline'}
        onCheckedChange={(precisePosition) =>
          onDraftChange((current) => ({ ...current, precisePosition }))
        }
      >
        {officeMessage(messages, 'document.picture.position.precise')}
      </OfficeCheckbox>
      <div className="work-document-picture-properties-position-grid">
        <PictureNumberRow
          label={officeMessage(
            messages,
            'document.picture.position.horizontal',
          )}
          ariaLabel={officeMessage(
            messages,
            'document.picture.position.horizontalAria',
          )}
          value={draft.horizontalOffset}
          unit={mm}
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
          ariaLabel={officeMessage(
            messages,
            'document.picture.position.horizontalRelativeAria',
          )}
          value={draft.horizontalReference}
          options={[...horizontalReferenceOptions(messages)]}
          disabled={draft.layout === 'inline' || !draft.precisePosition}
          onValueChange={(horizontalReference) =>
            onDraftChange((current) => ({ ...current, horizontalReference }))
          }
        />
        <PictureNumberRow
          label={officeMessage(messages, 'document.picture.position.vertical')}
          ariaLabel={officeMessage(
            messages,
            'document.picture.position.verticalAria',
          )}
          value={draft.verticalOffset}
          unit={mm}
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
          ariaLabel={officeMessage(
            messages,
            'document.picture.position.verticalRelativeAria',
          )}
          value={draft.verticalReference}
          options={[...verticalReferenceOptions(messages)]}
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

function PictureLayerSection({
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
  const messages = useOfficeMessages();
  const disabled = draft.layout === 'inline';
  return (
    <fieldset className="work-document-picture-properties-section layer">
      <legend>
        {officeMessage(messages, 'document.picture.layer.legend')}
      </legend>
      <PictureNumberRow
        label={officeMessage(messages, 'document.picture.layer.zOrder')}
        ariaLabel={officeMessage(messages, 'document.picture.layer.zOrderAria')}
        value={draft.relativeHeight}
        unit={officeMessage(messages, 'document.picture.layer.zOrderUnit')}
        min={0}
        max={MAX_DOCUMENT_IMAGE_RELATIVE_HEIGHT}
        step={1}
        disabled={disabled}
        invalid={Boolean(errors.relativeHeight)}
        onValueChange={(relativeHeight) =>
          onDraftChange((current) => ({ ...current, relativeHeight }))
        }
      />
      {errors.relativeHeight && <p role="alert">{errors.relativeHeight}</p>}
      <div className="work-document-picture-properties-layer-options">
        <OfficeCheckbox
          ariaLabel={officeMessage(
            messages,
            'document.picture.layer.behindAria',
          )}
          checked={draft.behindDocument}
          disabled={disabled}
          onCheckedChange={(behindDocument) =>
            onDraftChange((current) => ({ ...current, behindDocument }))
          }
        >
          {officeMessage(messages, 'document.picture.layer.behind')}
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel={officeMessage(
            messages,
            'document.picture.layer.overlapAria',
          )}
          checked={draft.allowOverlap}
          disabled={disabled}
          onCheckedChange={(allowOverlap) =>
            onDraftChange((current) => ({ ...current, allowOverlap }))
          }
        >
          {officeMessage(messages, 'document.picture.layer.overlap')}
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel={officeMessage(
            messages,
            'document.picture.layer.layoutInCellAria',
          )}
          checked={draft.layoutInCell}
          disabled={disabled}
          onCheckedChange={(layoutInCell) =>
            onDraftChange((current) => ({ ...current, layoutInCell }))
          }
        >
          {officeMessage(messages, 'document.picture.layer.layoutInCell')}
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel={officeMessage(
            messages,
            'document.picture.layer.lockAnchorAria',
          )}
          checked={draft.lockAnchor}
          disabled={disabled}
          onCheckedChange={(lockAnchor) =>
            onDraftChange((current) => ({ ...current, lockAnchor }))
          }
        >
          {officeMessage(messages, 'document.picture.layer.lockAnchor')}
        </OfficeCheckbox>
      </div>
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
  const messages = useOfficeMessages();
  const fields = [
    {
      key: 'cropTop' as const,
      label: officeMessage(messages, 'document.picture.crop.top'),
    },
    {
      key: 'cropRight' as const,
      label: officeMessage(messages, 'document.picture.crop.right'),
    },
    {
      key: 'cropBottom' as const,
      label: officeMessage(messages, 'document.picture.crop.bottom'),
    },
    {
      key: 'cropLeft' as const,
      label: officeMessage(messages, 'document.picture.crop.left'),
    },
  ];
  return (
    <fieldset className="work-document-picture-properties-section crop">
      <legend>{officeMessage(messages, 'document.picture.crop.legend')}</legend>
      <div className="work-document-picture-properties-crop-grid">
        {fields.map((field) => (
          <PictureNumberRow
            key={field.key}
            label={field.label}
            ariaLabel={officeMessage(messages, 'document.picture.crop.aria', {
              label: field.label,
            })}
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

function PictureTransformSection({
  draft,
  onDraftChange,
}: {
  draft: DocumentPicturePropertiesDraft;
  onDraftChange: React.Dispatch<
    React.SetStateAction<DocumentPicturePropertiesDraft>
  >;
}) {
  const messages = useOfficeMessages();
  return (
    <fieldset className="work-document-picture-properties-section transform">
      <legend>
        {officeMessage(messages, 'document.picture.transform.legend')}
      </legend>
      <div className="work-document-picture-properties-transform-row">
        <span>
          {officeMessage(messages, 'document.picture.transform.rotation')}
        </span>
        <OfficeSelect
          ariaLabel={officeMessage(
            messages,
            'document.picture.transform.rotationAria',
          )}
          value={String(draft.rotation)}
          options={[...rotationOptions(messages)]}
          onValueChange={(value) =>
            onDraftChange((current) => ({
              ...current,
              rotation: Number(value) as WorkDocumentImageRotation,
            }))
          }
        />
      </div>
      <div className="work-document-picture-properties-transform-options">
        <OfficeCheckbox
          ariaLabel={officeMessage(
            messages,
            'document.picture.transform.flipHAria',
          )}
          checked={draft.flipHorizontal}
          onCheckedChange={(flipHorizontal) =>
            onDraftChange((current) => ({ ...current, flipHorizontal }))
          }
        >
          {officeMessage(messages, 'document.picture.transform.flipH')}
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel={officeMessage(
            messages,
            'document.picture.transform.flipVAria',
          )}
          checked={draft.flipVertical}
          onCheckedChange={(flipVertical) =>
            onDraftChange((current) => ({ ...current, flipVertical }))
          }
        >
          {officeMessage(messages, 'document.picture.transform.flipV')}
        </OfficeCheckbox>
      </div>
      <small className="work-document-picture-properties-transform-help">
        {officeMessage(messages, 'document.picture.transform.help')}
      </small>
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
