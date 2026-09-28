import type { Editor } from '@tiptap/core';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  FlipHorizontal2,
  FlipVertical2,
  RotateCcw,
  RotateCw,
  Rows3,
  TextWrap,
  Trash2,
} from 'lucide-react';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import {
  defaultDocumentImageTransform,
  documentImageLayoutOptions,
  documentImageProperties,
  type WorkDocumentImageAlignment,
  type WorkDocumentImageLayout,
  type WorkDocumentImageRotation,
} from '../work-document-image-layout';
import { DocumentPicturePropertiesControl } from './document-picture-properties-dialog';
import { OfficeSelect } from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

export function DocumentPictureRibbon({ editor }: { editor: Editor }) {
  const messages = useOfficeMessages();
  const image = documentImageLayoutOptions(editor);
  const transform =
    documentImageProperties(editor).transform ??
    defaultDocumentImageTransform();
  const imageSelected = editor.isActive('image');
  const wrapDistanceValue = String(image.wrapDistance);
  const updateLayout = (layout: WorkDocumentImageLayout) =>
    editor.commands.setDocumentImageLayoutOptions({ layout });
  const updateAlignment = (alignment: WorkDocumentImageAlignment) =>
    editor.commands.setDocumentImageLayoutOptions({ alignment });
  const updateTransform = (
    next: (
      current: ReturnType<typeof defaultDocumentImageTransform>,
    ) => Partial<ReturnType<typeof defaultDocumentImageTransform>>,
  ) => {
    const current =
      documentImageProperties(editor).transform ??
      defaultDocumentImageTransform();
    return editor.commands.setDocumentImageProperties({
      transform: { ...current, ...next(current) },
    });
  };

  return (
    <>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.picture.ribbon.wrapGroup')}
      >
        <PictureButton
          label={officeMessage(messages, 'document.picture.wrap.inline')}
          active={image.layout === 'inline'}
          disabled={!imageSelected}
          onClick={() => updateLayout('inline')}
        >
          <Rows3 size={18} />
        </PictureButton>
        <PictureButton
          label={officeMessage(messages, 'document.picture.wrap.square')}
          active={image.layout === 'square'}
          disabled={!imageSelected}
          onClick={() => updateLayout('square')}
        >
          <TextWrap size={18} />
        </PictureButton>
        <PictureButton
          label={officeMessage(messages, 'document.picture.wrap.tight')}
          active={image.layout === 'tight'}
          disabled={!imageSelected}
          onClick={() => updateLayout('tight')}
        >
          <TextWrap size={18} />
        </PictureButton>
        <PictureButton
          label={officeMessage(messages, 'document.picture.wrap.through')}
          active={image.layout === 'through'}
          disabled={!imageSelected}
          onClick={() => updateLayout('through')}
        >
          <TextWrap size={18} />
        </PictureButton>
        <PictureButton
          label={officeMessage(messages, 'document.picture.wrap.topBottom')}
          active={image.layout === 'topBottom'}
          disabled={!imageSelected}
          onClick={() => updateLayout('topBottom')}
        >
          <Rows3 size={18} />
        </PictureButton>
        <PictureButton
          label={officeMessage(messages, 'document.picture.wrap.none')}
          active={image.layout === 'none'}
          disabled={!imageSelected}
          onClick={() => updateLayout('none')}
        >
          <TextWrap size={18} />
        </PictureButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(
          messages,
          'document.picture.ribbon.positionGroup',
        )}
      >
        <PictureButton
          label={officeMessage(messages, 'document.picture.align.left')}
          active={image.alignment === 'left'}
          disabled={!imageSelected}
          onClick={() => updateAlignment('left')}
        >
          <AlignLeft size={18} />
        </PictureButton>
        <PictureButton
          label={officeMessage(messages, 'document.picture.align.center')}
          active={image.alignment === 'center'}
          disabled={!imageSelected}
          onClick={() => updateAlignment('center')}
        >
          <AlignCenter size={18} />
        </PictureButton>
        <PictureButton
          label={officeMessage(messages, 'document.picture.align.right')}
          active={image.alignment === 'right'}
          disabled={!imageSelected}
          onClick={() => updateAlignment('right')}
        >
          <AlignRight size={18} />
        </PictureButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(
          messages,
          'document.picture.ribbon.transformGroup',
        )}
      >
        <PictureButton
          label={officeMessage(messages, 'document.picture.ribbon.rotateLeft')}
          disabled={!imageSelected}
          onClick={() =>
            updateTransform((current) => ({
              rotation: rotateImage(current.rotation, -90),
            }))
          }
        >
          <RotateCcw size={18} />
        </PictureButton>
        <PictureButton
          label={officeMessage(messages, 'document.picture.ribbon.rotateRight')}
          disabled={!imageSelected}
          onClick={() =>
            updateTransform((current) => ({
              rotation: rotateImage(current.rotation, 90),
            }))
          }
        >
          <RotateCw size={18} />
        </PictureButton>
        <PictureButton
          label={officeMessage(messages, 'document.picture.transform.flipH')}
          active={transform.flipHorizontal}
          disabled={!imageSelected}
          onClick={() =>
            updateTransform((current) => ({
              flipHorizontal: !current.flipHorizontal,
            }))
          }
        >
          <FlipHorizontal2 size={18} />
        </PictureButton>
        <PictureButton
          label={officeMessage(messages, 'document.picture.transform.flipV')}
          active={transform.flipVertical}
          disabled={!imageSelected}
          onClick={() =>
            updateTransform((current) => ({
              flipVertical: !current.flipVertical,
            }))
          }
        >
          <FlipVertical2 size={18} />
        </PictureButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(
          messages,
          'document.picture.ribbon.distanceGroup',
        )}
      >
        <OfficeSelect
          className="work-document-picture-wrap-distance-select"
          ariaLabel={officeMessage(
            messages,
            'document.picture.ribbon.distanceAria',
          )}
          value={wrapDistanceValue}
          options={imageWrapDistanceOptionsForValue(
            messages,
            wrapDistanceValue,
          )}
          disabled={
            !imageSelected ||
            image.layout === 'inline' ||
            image.layout === 'none'
          }
          onValueChange={(value) =>
            editor.commands.setDocumentImageLayoutOptions({
              wrapDistance: Number(value),
            })
          }
        />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.picture.ribbon.pictureGroup')}
      >
        <DocumentPicturePropertiesControl editor={editor} />
        <PictureButton
          label={officeMessage(messages, 'document.picture.ribbon.delete')}
          disabled={!imageSelected}
          onClick={() => editor.chain().focus().deleteSelection().run()}
        >
          <Trash2 size={18} />
        </PictureButton>
      </WorkOfficeRibbonGroup>
    </>
  );
}

function rotateImage(
  rotation: WorkDocumentImageRotation,
  delta: -90 | 90,
): WorkDocumentImageRotation {
  const next = (rotation + delta + 360) % 360;
  return next as WorkDocumentImageRotation;
}

function imageWrapDistanceOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: '0',
      label: officeMessage(messages, 'document.picture.ribbon.distance.none'),
    },
    {
      value: '2',
      label: officeMessage(messages, 'document.picture.ribbon.distance.mm', {
        value: '2',
      }),
    },
    {
      value: '3',
      label: officeMessage(messages, 'document.picture.ribbon.distance.mm', {
        value: '3',
      }),
    },
    {
      value: '5',
      label: officeMessage(messages, 'document.picture.ribbon.distance.mm', {
        value: '5',
      }),
    },
    {
      value: '10',
      label: officeMessage(messages, 'document.picture.ribbon.distance.mm', {
        value: '10',
      }),
    },
  ] as const;
}

function imageWrapDistanceOptionsForValue(
  messages: OfficeMessageCatalog,
  value: string,
) {
  const options = imageWrapDistanceOptions(messages);
  if (options.some((option) => option.value === value)) {
    return options;
  }
  return [
    ...options,
    {
      value,
      label: officeMessage(messages, 'document.picture.ribbon.distance.mm', {
        value,
      }),
    },
  ];
}

function PictureButton({
  label,
  active = false,
  disabled = false,
  onClick,
  children,
}: {
  label: string;
  active?: boolean;
  disabled?: boolean;
  onClick: () => void;
  children: React.ReactNode;
}) {
  return (
    <WorkOfficeRibbonButton
      label={label}
      displayLabel
      active={active}
      disabled={disabled}
      onClick={onClick}
    >
      {children}
    </WorkOfficeRibbonButton>
  );
}
