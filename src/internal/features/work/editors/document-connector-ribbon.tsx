import type { Editor } from '@tiptap/core';
import {
  ArrowDownRight,
  ArrowUpRight,
  Rows3,
  SquareDashedMousePointer,
  Trash2,
} from 'lucide-react';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import {
  DOCUMENT_CONNECTOR_DEFAULTS,
  DOCUMENT_CONNECTOR_LIMITS,
  documentConnectorProperties,
  type WorkDocumentConnectorArrow,
  type WorkDocumentConnectorKind,
  type WorkDocumentConnectorLayout,
  type WorkDocumentConnectorLineStyle,
  type WorkDocumentConnectorProperties,
  type WorkDocumentConnectorReference,
  type WorkDocumentConnectorVerticalReference,
} from '../work-document-connector';
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

export function DocumentConnectorRibbon({ editor }: { editor: Editor }) {
  const messages = useOfficeMessages();
  const properties = documentConnectorProperties(editor);
  const selected = editor.isActive('documentConnector');
  const update = (value: Partial<WorkDocumentConnectorProperties>) => {
    if (!selected) return;
    editor.commands.setDocumentConnectorProperties(value, {
      restoreFocus: false,
    });
  };
  const updateLayout = (layout: WorkDocumentConnectorLayout) => {
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
        label={officeMessage(messages, 'document.connector.group.line')}
        priority="high"
      >
        <OfficeSelect<WorkDocumentConnectorKind>
          className="work-document-connector-kind-select"
          ariaLabel={officeMessage(messages, 'document.connector.kindAria')}
          value={properties.connectorKind}
          options={connectorKindOptions(messages)}
          disabled={!selected}
          onValueChange={(connectorKind) => update({ connectorKind })}
        />
        <OfficeColorPicker
          ariaLabel={officeMessage(messages, 'document.connector.lineColorAria')}
          triggerLabel={officeMessage(
            messages,
            'document.connector.lineColorTrigger',
          )}
          value={properties.lineColor}
          disabled={!selected}
          onValueChange={(lineColor) => update({ lineColor })}
        />
        <OfficeSelect
          className="work-document-connector-width-select"
          ariaLabel={officeMessage(messages, 'document.connector.lineWidthAria')}
          value={String(properties.lineWidth)}
          options={widthOptionsForValue(messages, String(properties.lineWidth))}
          disabled={!selected}
          onValueChange={(value) => update({ lineWidth: Number(value) })}
        />
        <OfficeSelect<WorkDocumentConnectorLineStyle>
          className="work-document-connector-style-select"
          ariaLabel={officeMessage(messages, 'document.connector.lineStyleAria')}
          value={properties.lineStyle}
          options={lineStyleOptions(messages)}
          disabled={!selected}
          onValueChange={(lineStyle) => update({ lineStyle })}
        />
        <OfficeSelect<WorkDocumentConnectorArrow>
          className="work-document-connector-arrow-select"
          ariaLabel={officeMessage(
            messages,
            'document.connector.startArrowAria',
          )}
          value={properties.startArrow}
          options={arrowOptions(messages)}
          disabled={!selected}
          onValueChange={(startArrow) => update({ startArrow })}
        />
        <ArrowUpRight size={18} aria-hidden="true" />
        <OfficeSelect<WorkDocumentConnectorArrow>
          className="work-document-connector-arrow-select"
          ariaLabel={officeMessage(messages, 'document.connector.endArrowAria')}
          value={properties.endArrow}
          options={arrowOptions(messages)}
          disabled={!selected}
          onValueChange={(endArrow) => update({ endArrow })}
        />
        <ArrowDownRight size={18} aria-hidden="true" />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.connector.group.layout')}
        priority="high"
      >
        {layoutOptions(messages).map((option) => {
          const Icon = option.icon;
          return (
            <ConnectorButton
              key={option.value}
              label={option.label}
              active={selected && properties.layout === option.value}
              disabled={!selected}
              onClick={() => updateLayout(option.value)}
            >
              <Icon size={18} />
            </ConnectorButton>
          );
        })}
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.picture.size.legend')}
      >
        <ConnectorNumberField
          label={officeMessage(messages, 'document.connector.size.width')}
          ariaLabel={officeMessage(
            messages,
            'document.connector.size.widthAria',
          )}
          value={properties.width}
          min={DOCUMENT_CONNECTOR_LIMITS.width.min}
          max={DOCUMENT_CONNECTOR_LIMITS.width.max}
          disabled={!selected}
          onValueCommit={(width) => update({ width })}
        />
        <ConnectorNumberField
          label={officeMessage(messages, 'document.connector.size.height')}
          ariaLabel={officeMessage(
            messages,
            'document.connector.size.heightAria',
          )}
          value={properties.height}
          min={DOCUMENT_CONNECTOR_LIMITS.height.min}
          max={DOCUMENT_CONNECTOR_LIMITS.height.max}
          disabled={!selected}
          onValueCommit={(height) => update({ height })}
        />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.connector.group.endpoints')}
      >
        <ConnectorNumberField
          label={officeMessage(
            messages,
            'document.connector.endpoint.startX',
          )}
          ariaLabel={officeMessage(
            messages,
            'document.connector.endpoint.startXAria',
          )}
          value={properties.startX}
          min={DOCUMENT_CONNECTOR_LIMITS.endpoint.min}
          max={DOCUMENT_CONNECTOR_LIMITS.endpoint.max}
          disabled={!selected}
          onValueCommit={(startX) => update({ startX })}
        />
        <ConnectorNumberField
          label={officeMessage(
            messages,
            'document.connector.endpoint.startY',
          )}
          ariaLabel={officeMessage(
            messages,
            'document.connector.endpoint.startYAria',
          )}
          value={properties.startY}
          min={DOCUMENT_CONNECTOR_LIMITS.endpoint.min}
          max={DOCUMENT_CONNECTOR_LIMITS.endpoint.max}
          disabled={!selected}
          onValueCommit={(startY) => update({ startY })}
        />
        <ConnectorNumberField
          label={officeMessage(messages, 'document.connector.endpoint.endX')}
          ariaLabel={officeMessage(
            messages,
            'document.connector.endpoint.endXAria',
          )}
          value={properties.endX}
          min={DOCUMENT_CONNECTOR_LIMITS.endpoint.min}
          max={DOCUMENT_CONNECTOR_LIMITS.endpoint.max}
          disabled={!selected}
          onValueCommit={(endX) => update({ endX })}
        />
        <ConnectorNumberField
          label={officeMessage(messages, 'document.connector.endpoint.endY')}
          ariaLabel={officeMessage(
            messages,
            'document.connector.endpoint.endYAria',
          )}
          value={properties.endY}
          min={DOCUMENT_CONNECTOR_LIMITS.endpoint.min}
          max={DOCUMENT_CONNECTOR_LIMITS.endpoint.max}
          disabled={!selected}
          onValueCommit={(endY) => update({ endY })}
        />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.picture.position.legend')}
      >
        <div className="work-office-field work-document-connector-offset-field">
          <span>
            {officeMessage(
              messages,
              'document.connector.position.horizontalRelative',
            )}
          </span>
          <OfficeSelect<WorkDocumentConnectorReference>
            ariaLabel={officeMessage(
              messages,
              'document.connector.position.horizontalRelativeAria',
            )}
            value={properties.horizontalReference}
            options={horizontalReferenceOptions(messages)}
            disabled={!selected || properties.layout !== 'floating'}
            onValueChange={(horizontalReference) =>
              update({ horizontalReference })
            }
          />
        </div>
        <div className="work-office-field work-document-connector-offset-field">
          <span>
            {officeMessage(
              messages,
              'document.connector.position.verticalRelative',
            )}
          </span>
          <OfficeSelect<WorkDocumentConnectorVerticalReference>
            ariaLabel={officeMessage(
              messages,
              'document.connector.position.verticalRelativeAria',
            )}
            value={properties.verticalReference}
            options={verticalReferenceOptions(messages)}
            disabled={!selected || properties.layout !== 'floating'}
            onValueChange={(verticalReference) => update({ verticalReference })}
          />
        </div>
        <ConnectorNumberField
          label={officeMessage(
            messages,
            'document.connector.position.horizontalOffset',
          )}
          ariaLabel={officeMessage(
            messages,
            'document.connector.position.horizontalOffsetAria',
          )}
          value={properties.horizontalOffset ?? 0}
          min={DOCUMENT_CONNECTOR_LIMITS.offset.min}
          max={DOCUMENT_CONNECTOR_LIMITS.offset.max}
          disabled={!selected || properties.layout !== 'floating'}
          onValueCommit={(horizontalOffset) => update({ horizontalOffset })}
        />
        <ConnectorNumberField
          label={officeMessage(
            messages,
            'document.connector.position.verticalOffset',
          )}
          ariaLabel={officeMessage(
            messages,
            'document.connector.position.verticalOffsetAria',
          )}
          value={properties.verticalOffset ?? 0}
          min={DOCUMENT_CONNECTOR_LIMITS.offset.min}
          max={DOCUMENT_CONNECTOR_LIMITS.offset.max}
          disabled={!selected || properties.layout !== 'floating'}
          onValueCommit={(verticalOffset) => update({ verticalOffset })}
        />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.connector.group.object')}
      >
        <ConnectorButton
          label={officeMessage(messages, 'document.connector.delete')}
          disabled={!selected}
          onClick={() => editor.commands.deleteDocumentConnector()}
        >
          <Trash2 size={18} />
        </ConnectorButton>
        <ConnectorButton
          label={officeMessage(messages, 'document.connector.resetStyle')}
          disabled={!selected}
          onClick={() =>
            update({
              width: DOCUMENT_CONNECTOR_DEFAULTS.width,
              height: DOCUMENT_CONNECTOR_DEFAULTS.height,
              connectorKind: DOCUMENT_CONNECTOR_DEFAULTS.connectorKind,
              layout: DOCUMENT_CONNECTOR_DEFAULTS.layout,
              horizontalOffset: DOCUMENT_CONNECTOR_DEFAULTS.horizontalOffset,
              verticalOffset: DOCUMENT_CONNECTOR_DEFAULTS.verticalOffset,
              horizontalReference:
                DOCUMENT_CONNECTOR_DEFAULTS.horizontalReference,
              verticalReference: DOCUMENT_CONNECTOR_DEFAULTS.verticalReference,
              startX: DOCUMENT_CONNECTOR_DEFAULTS.startX,
              startY: DOCUMENT_CONNECTOR_DEFAULTS.startY,
              endX: DOCUMENT_CONNECTOR_DEFAULTS.endX,
              endY: DOCUMENT_CONNECTOR_DEFAULTS.endY,
              lineColor: DOCUMENT_CONNECTOR_DEFAULTS.lineColor,
              lineWidth: DOCUMENT_CONNECTOR_DEFAULTS.lineWidth,
              lineStyle: DOCUMENT_CONNECTOR_DEFAULTS.lineStyle,
              startArrow: DOCUMENT_CONNECTOR_DEFAULTS.startArrow,
              endArrow: DOCUMENT_CONNECTOR_DEFAULTS.endArrow,
            })
          }
        >
          <Rows3 size={18} />
        </ConnectorButton>
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
      label: officeMessage(messages, 'document.connector.layout.floating'),
      icon: SquareDashedMousePointer,
    },
  ];
}

function arrowOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'none' as const,
      label: officeMessage(messages, 'document.connector.arrow.none'),
    },
    {
      value: 'triangle' as const,
      label: officeMessage(messages, 'document.connector.arrow.triangle'),
    },
    {
      value: 'stealth' as const,
      label: officeMessage(messages, 'document.connector.arrow.stealth'),
    },
    {
      value: 'diamond' as const,
      label: officeMessage(messages, 'document.connector.arrow.diamond'),
    },
    {
      value: 'oval' as const,
      label: officeMessage(messages, 'document.connector.arrow.oval'),
    },
    {
      value: 'open' as const,
      label: officeMessage(messages, 'document.connector.arrow.open'),
    },
  ];
}

function connectorKindOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'straight' as const,
      label: officeMessage(messages, 'document.connector.kind.straight'),
    },
    {
      value: 'elbow' as const,
      label: officeMessage(messages, 'document.connector.kind.elbow'),
    },
    {
      value: 'curved' as const,
      label: officeMessage(messages, 'document.connector.kind.curved'),
    },
  ];
}

function lineStyleOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'solid' as const,
      label: officeMessage(messages, 'document.connector.line.solid'),
    },
    {
      value: 'dash' as const,
      label: officeMessage(messages, 'document.connector.line.dash'),
    },
    {
      value: 'dot' as const,
      label: officeMessage(messages, 'document.connector.line.dot'),
    },
    {
      value: 'dashDot' as const,
      label: officeMessage(messages, 'document.connector.line.dashDot'),
    },
  ];
}

function widthOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: '0.35',
      label: officeMessage(messages, 'document.connector.width.thin'),
    },
    {
      value: '0.7',
      label: officeMessage(messages, 'document.connector.width.medium'),
    },
    {
      value: '1.4',
      label: officeMessage(messages, 'document.connector.width.thick'),
    },
  ] as const;
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

function ConnectorNumberField({
  label,
  ariaLabel,
  value,
  min,
  max,
  disabled,
  onValueCommit,
}: {
  label: string;
  ariaLabel: string;
  value: number;
  min: number;
  max: number;
  disabled: boolean;
  onValueCommit: (value: number) => void;
}) {
  return (
    <div className="work-office-field work-document-connector-endpoint-field">
      <span>{label}</span>
      <CommittedOfficeNumberField
        ariaLabel={ariaLabel}
        value={value}
        min={min}
        max={max}
        step={0.1}
        disabled={disabled}
        normalizeValue={(next) => {
          const number = Number(next);
          if (!Number.isFinite(number) || number < min || number > max) {
            return null;
          }
          return Number(number.toFixed(2));
        }}
        onValueCommit={onValueCommit}
      />
    </div>
  );
}

function widthOptionsForValue(messages: OfficeMessageCatalog, value: string) {
  const options = widthOptions(messages);
  if (options.some((option) => option.value === value)) return options;
  return [
    ...options,
    {
      value,
      label: officeMessage(messages, 'document.connector.width.customMm', {
        value,
      }),
    },
  ] as const;
}

function ConnectorButton({
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
