import type { Editor } from '@tiptap/core';
import {
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  Move,
  Rows3,
  Shapes,
  SquareDashedMousePointer,
  Trash2,
} from 'lucide-react';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import {
  DOCUMENT_TEXT_BOX_DEFAULTS,
  DOCUMENT_TEXT_BOX_LIMITS,
  documentTextBoxProperties,
  type WorkDocumentShapeType,
  type WorkDocumentTextBoxHorizontalReference,
  type WorkDocumentTextBoxLayout,
  type WorkDocumentTextBoxProperties,
  type WorkDocumentTextBoxVerticalAlign,
  type WorkDocumentTextBoxVerticalReference,
} from '../work-document-text-box';
import {
  CommittedOfficeNumberField,
  OfficeColorPicker,
  OfficeSelect,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

export function DocumentTextBoxRibbon({ editor }: { editor: Editor }) {
  const messages = useOfficeMessages();
  const properties = documentTextBoxProperties(editor);
  const selected = editor.isActive('documentTextBox');
  const update = (value: Partial<WorkDocumentTextBoxProperties>) => {
    if (!selected) return;
    // Keep focus in the active ribbon control while a committed field is
    // settling.  Restoring the editor focus here races with the next control
    // interaction and can send its keystrokes into the text box content.
    editor.commands.setDocumentTextBoxProperties(value, {
      restoreFocus: false,
    });
  };
  const updateLayout = (layout: WorkDocumentTextBoxLayout) => {
    update({
      layout,
      ...(layout === 'floating'
        ? {
            horizontalOffset: properties.horizontalOffset ?? 0,
            verticalOffset: properties.verticalOffset ?? 0,
          }
        : {}),
    });
  };

  return (
    <>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.textBox.group.shape')}
        priority="high"
      >
        <OfficeSelect<WorkDocumentShapeType>
          className="work-document-text-box-shape-select"
          ariaLabel={officeMessage(messages, 'document.textBox.shapeAria')}
          value={properties.shapeType}
          options={shapeOptions(messages)}
          disabled={!selected}
          onValueChange={(shapeType) => update({ shapeType })}
        />
        <Shapes size={18} aria-hidden="true" />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.textBox.group.layout')}
        priority="high"
      >
        {layoutOptions(messages).map((option) => {
          const Icon = option.icon;
          return (
            <TextBoxButton
              key={option.value}
              label={option.label}
              active={selected && properties.layout === option.value}
              disabled={!selected}
              onClick={() => updateLayout(option.value)}
            >
              <Icon size={18} />
            </TextBoxButton>
          );
        })}
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.picture.size.legend')}
      >
        <div className="work-office-field work-document-text-box-size-field">
          <span>
            {officeMessage(messages, 'document.textBox.size.width')}
          </span>
          <CommittedOfficeNumberField
            ariaLabel={officeMessage(
              messages,
              'document.textBox.size.widthAria',
            )}
            value={properties.width}
            min={DOCUMENT_TEXT_BOX_LIMITS.width.min}
            max={DOCUMENT_TEXT_BOX_LIMITS.width.max}
            step={0.1}
            disabled={!selected}
            normalizeValue={(value) =>
              normalizeTextBoxNumber(
                value,
                DOCUMENT_TEXT_BOX_LIMITS.width.min,
                DOCUMENT_TEXT_BOX_LIMITS.width.max,
              )
            }
            onValueCommit={(width) => update({ width })}
          />
        </div>
        <div className="work-office-field work-document-text-box-size-field">
          <span>
            {officeMessage(messages, 'document.textBox.size.height')}
          </span>
          <CommittedOfficeNumberField
            ariaLabel={officeMessage(
              messages,
              'document.textBox.size.heightAria',
            )}
            value={properties.height}
            min={DOCUMENT_TEXT_BOX_LIMITS.height.min}
            max={DOCUMENT_TEXT_BOX_LIMITS.height.max}
            step={0.1}
            disabled={!selected}
            normalizeValue={(value) =>
              normalizeTextBoxNumber(
                value,
                DOCUMENT_TEXT_BOX_LIMITS.height.min,
                DOCUMENT_TEXT_BOX_LIMITS.height.max,
              )
            }
            onValueCommit={(height) => update({ height })}
          />
        </div>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.picture.position.legend')}
      >
        <div className="work-office-field work-document-text-box-position-field">
          <span>
            {officeMessage(
              messages,
              'document.textBox.position.horizontalRelative',
            )}
          </span>
          <OfficeSelect<WorkDocumentTextBoxHorizontalReference>
            ariaLabel={officeMessage(
              messages,
              'document.textBox.position.horizontalRelativeAria',
            )}
            value={properties.horizontalReference}
            options={horizontalReferenceOptions(messages)}
            disabled={!selected || properties.layout !== 'floating'}
            onValueChange={(horizontalReference) =>
              update({ horizontalReference })
            }
          />
        </div>
        <div className="work-office-field work-document-text-box-position-field">
          <span>
            {officeMessage(
              messages,
              'document.textBox.position.verticalRelative',
            )}
          </span>
          <OfficeSelect<WorkDocumentTextBoxVerticalReference>
            ariaLabel={officeMessage(
              messages,
              'document.textBox.position.verticalRelativeAria',
            )}
            value={properties.verticalReference}
            options={verticalReferenceOptions(messages)}
            disabled={!selected || properties.layout !== 'floating'}
            onValueChange={(verticalReference) => update({ verticalReference })}
          />
        </div>
        <div className="work-office-field work-document-text-box-offset-field">
          <span>
            {officeMessage(
              messages,
              'document.textBox.position.horizontalOffset',
            )}
          </span>
          <CommittedOfficeNumberField
            ariaLabel={officeMessage(
              messages,
              'document.textBox.position.horizontalOffsetAria',
            )}
            value={properties.horizontalOffset ?? 0}
            min={DOCUMENT_TEXT_BOX_LIMITS.offset.min}
            max={DOCUMENT_TEXT_BOX_LIMITS.offset.max}
            step={0.1}
            disabled={!selected || properties.layout !== 'floating'}
            normalizeValue={(value) =>
              normalizeTextBoxNumber(
                value,
                DOCUMENT_TEXT_BOX_LIMITS.offset.min,
                DOCUMENT_TEXT_BOX_LIMITS.offset.max,
              )
            }
            onValueCommit={(horizontalOffset) => update({ horizontalOffset })}
          />
        </div>
        <div className="work-office-field work-document-text-box-offset-field">
          <span>
            {officeMessage(
              messages,
              'document.textBox.position.verticalOffset',
            )}
          </span>
          <CommittedOfficeNumberField
            ariaLabel={officeMessage(
              messages,
              'document.textBox.position.verticalOffsetAria',
            )}
            value={properties.verticalOffset ?? 0}
            min={DOCUMENT_TEXT_BOX_LIMITS.offset.min}
            max={DOCUMENT_TEXT_BOX_LIMITS.offset.max}
            step={0.1}
            disabled={!selected || properties.layout !== 'floating'}
            normalizeValue={(value) =>
              normalizeTextBoxNumber(
                value,
                DOCUMENT_TEXT_BOX_LIMITS.offset.min,
                DOCUMENT_TEXT_BOX_LIMITS.offset.max,
              )
            }
            onValueCommit={(verticalOffset) => update({ verticalOffset })}
          />
        </div>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.textBox.group.align')}
      >
        {verticalAlignOptions(messages).map((option) => {
          const Icon = option.icon;
          return (
            <TextBoxButton
              key={option.value}
              label={option.label}
              active={selected && properties.verticalAlign === option.value}
              disabled={!selected}
              onClick={() => update({ verticalAlign: option.value })}
            >
              <Icon size={18} />
            </TextBoxButton>
          );
        })}
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.textBox.group.fillStroke')}
      >
        <OfficeColorPicker
          ariaLabel={officeMessage(messages, 'document.textBox.fillAria')}
          triggerLabel={officeMessage(
            messages,
            'document.textBox.fillTrigger',
          )}
          value={properties.fill}
          disabled={!selected}
          resetAction={{
            kind: 'automatic',
            label: officeMessage(messages, 'document.textBox.fillNone'),
            onSelect: () => update({ fill: 'transparent' }),
          }}
          onValueChange={(fill) => update({ fill })}
        />
        <OfficeColorPicker
          ariaLabel={officeMessage(messages, 'document.textBox.strokeAria')}
          triggerLabel={officeMessage(
            messages,
            'document.textBox.strokeTrigger',
          )}
          value={
            properties.borderColor === 'none'
              ? '#ffffff'
              : properties.borderColor
          }
          disabled={!selected}
          resetAction={{
            kind: 'none',
            label: officeMessage(messages, 'document.textBox.strokeNone'),
            onSelect: () => update({ borderColor: 'none', borderWidth: 0 }),
          }}
          onValueChange={(borderColor) => update({ borderColor })}
        />
        <OfficeSelect
          className="work-document-text-box-border-width-select"
          ariaLabel={officeMessage(
            messages,
            'document.textBox.strokeWidthAria',
          )}
          value={String(properties.borderWidth)}
          options={borderWidthOptionsForValue(
            messages,
            String(properties.borderWidth),
          )}
          disabled={!selected}
          onValueChange={(value) =>
            update({
              borderWidth: Number(value),
              ...(value === '0' ? { borderColor: 'none' } : {}),
            })
          }
        />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.textBox.group.object')}
      >
        <TextBoxButton
          label={officeMessage(messages, 'document.textBox.delete')}
          disabled={!selected}
          onClick={() => editor.commands.deleteDocumentTextBox()}
        >
          <Trash2 size={18} />
        </TextBoxButton>
        <TextBoxButton
          label={officeMessage(messages, 'document.textBox.resetStyle')}
          disabled={!selected}
          onClick={() =>
            update({
              width: DOCUMENT_TEXT_BOX_DEFAULTS.width,
              height: DOCUMENT_TEXT_BOX_DEFAULTS.height,
              shapeType: DOCUMENT_TEXT_BOX_DEFAULTS.shapeType,
              layout: DOCUMENT_TEXT_BOX_DEFAULTS.layout,
              horizontalOffset: DOCUMENT_TEXT_BOX_DEFAULTS.horizontalOffset,
              verticalOffset: DOCUMENT_TEXT_BOX_DEFAULTS.verticalOffset,
              horizontalReference:
                DOCUMENT_TEXT_BOX_DEFAULTS.horizontalReference,
              verticalReference: DOCUMENT_TEXT_BOX_DEFAULTS.verticalReference,
              fill: DOCUMENT_TEXT_BOX_DEFAULTS.fill,
              borderColor: DOCUMENT_TEXT_BOX_DEFAULTS.borderColor,
              borderWidth: DOCUMENT_TEXT_BOX_DEFAULTS.borderWidth,
              padding: DOCUMENT_TEXT_BOX_DEFAULTS.padding,
              verticalAlign: DOCUMENT_TEXT_BOX_DEFAULTS.verticalAlign,
            })
          }
        >
          <Move size={18} />
        </TextBoxButton>
      </WorkOfficeRibbonGroup>
    </>
  );
}

function layoutOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'inline' as const,
      label: officeMessage(messages, 'document.picture.wrap.inline'),
      icon: Rows3,
    },
    {
      value: 'floating' as const,
      label: officeMessage(messages, 'document.textBox.layout.floating'),
      icon: SquareDashedMousePointer,
    },
  ];
}

function shapeOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'rectangle' as const,
      label: officeMessage(messages, 'document.textBox.shape.rectangle'),
    },
    {
      value: 'roundedRectangle' as const,
      label: officeMessage(messages, 'document.textBox.shape.roundedRectangle'),
    },
    {
      value: 'ellipse' as const,
      label: officeMessage(messages, 'document.textBox.shape.ellipse'),
    },
    {
      value: 'diamond' as const,
      label: officeMessage(messages, 'document.textBox.shape.diamond'),
    },
    {
      value: 'triangle' as const,
      label: officeMessage(messages, 'document.textBox.shape.triangle'),
    },
    {
      value: 'parallelogram' as const,
      label: officeMessage(messages, 'document.textBox.shape.parallelogram'),
    },
    {
      value: 'hexagon' as const,
      label: officeMessage(messages, 'document.textBox.shape.hexagon'),
    },
  ];
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
  ];
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
  ];
}

function verticalAlignOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'top' as const,
      label: officeMessage(messages, 'document.textBox.align.top'),
      icon: AlignVerticalJustifyStart,
    },
    {
      value: 'center' as const,
      label: officeMessage(messages, 'document.textBox.align.center'),
      icon: AlignVerticalJustifyCenter,
    },
    {
      value: 'bottom' as const,
      label: officeMessage(messages, 'document.textBox.align.bottom'),
      icon: AlignVerticalJustifyEnd,
    },
  ];
}

function borderWidthOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: '0',
      label: officeMessage(messages, 'document.textBox.border.none'),
    },
    {
      value: '0.35',
      label: officeMessage(messages, 'document.textBox.border.thin'),
    },
    {
      value: '0.7',
      label: officeMessage(messages, 'document.textBox.border.medium'),
    },
    {
      value: '1.4',
      label: officeMessage(messages, 'document.textBox.border.thick'),
    },
  ] as const;
}

function TextBoxButton({
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

function normalizeTextBoxNumber(
  value: string,
  min: number,
  max: number,
): number | null {
  const number = Number(value);
  if (!Number.isFinite(number) || number < min || number > max) return null;
  return Number(number.toFixed(2));
}

function borderWidthOptionsForValue(
  messages: OfficeMessageCatalog,
  value: string,
) {
  const options = borderWidthOptions(messages);
  if (options.some((option) => option.value === value)) {
    return options;
  }
  return [
    ...options,
    {
      value,
      label: officeMessage(messages, 'document.textBox.border.customMm', {
        value,
      }),
    },
  ] as const;
}
