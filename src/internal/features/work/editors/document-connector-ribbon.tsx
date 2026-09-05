import type { Editor } from '@tiptap/core';
import {
  ArrowDownRight,
  ArrowUpRight,
  Rows3,
  SquareDashedMousePointer,
  Trash2,
} from 'lucide-react';
import {
  DOCUMENT_CONNECTOR_DEFAULTS,
  DOCUMENT_CONNECTOR_LIMITS,
  documentConnectorProperties,
  type WorkDocumentConnectorArrow,
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
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

const layoutOptions = [
  { value: 'inline', label: '嵌入文字', icon: Rows3 },
  { value: 'floating', label: '浮于文字上方', icon: SquareDashedMousePointer },
] as const satisfies readonly {
  value: WorkDocumentConnectorLayout;
  label: string;
  icon: typeof Rows3;
}[];

const arrowOptions = [
  { value: 'none', label: '无箭头' },
  { value: 'triangle', label: '三角箭头' },
  { value: 'stealth', label: '隐形箭头' },
  { value: 'diamond', label: '菱形箭头' },
  { value: 'oval', label: '圆形箭头' },
  { value: 'open', label: '开放箭头' },
] as const satisfies readonly {
  value: WorkDocumentConnectorArrow;
  label: string;
}[];

const lineStyleOptions = [
  { value: 'solid', label: '实线' },
  { value: 'dash', label: '虚线' },
  { value: 'dot', label: '点线' },
  { value: 'dashDot', label: '点划线' },
] as const satisfies readonly {
  value: WorkDocumentConnectorLineStyle;
  label: string;
}[];

const widthOptions = [
  { value: '0.35', label: '细（0.35 mm）' },
  { value: '0.7', label: '中（0.7 mm）' },
  { value: '1.4', label: '粗（1.4 mm）' },
] as const;

const horizontalReferenceOptions = [
  { value: 'column', label: '栏' },
  { value: 'margin', label: '页边距' },
  { value: 'page', label: '页面' },
] as const satisfies readonly {
  value: WorkDocumentConnectorReference;
  label: string;
}[];

const verticalReferenceOptions = [
  { value: 'paragraph', label: '段落' },
  { value: 'margin', label: '页边距' },
  { value: 'page', label: '页面' },
] as const satisfies readonly {
  value: WorkDocumentConnectorVerticalReference;
  label: string;
}[];

export function DocumentConnectorRibbon({ editor }: { editor: Editor }) {
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
      <WorkOfficeRibbonGroup label="连接线" priority="high">
        <OfficeColorPicker
          ariaLabel="连接符线条颜色"
          triggerLabel="线条颜色"
          value={properties.lineColor}
          disabled={!selected}
          onValueChange={(lineColor) => update({ lineColor })}
        />
        <OfficeSelect
          ariaLabel="连接符线条粗细"
          value={String(properties.lineWidth)}
          options={widthOptionsForValue(String(properties.lineWidth))}
          disabled={!selected}
          onValueChange={(value) => update({ lineWidth: Number(value) })}
        />
        <OfficeSelect<WorkDocumentConnectorLineStyle>
          ariaLabel="连接符线条样式"
          value={properties.lineStyle}
          options={lineStyleOptions}
          disabled={!selected}
          onValueChange={(lineStyle) => update({ lineStyle })}
        />
        <OfficeSelect<WorkDocumentConnectorArrow>
          ariaLabel="连接符起点箭头"
          value={properties.startArrow}
          options={arrowOptions}
          disabled={!selected}
          onValueChange={(startArrow) => update({ startArrow })}
        />
        <ArrowUpRight size={18} aria-hidden="true" />
        <OfficeSelect<WorkDocumentConnectorArrow>
          ariaLabel="连接符终点箭头"
          value={properties.endArrow}
          options={arrowOptions}
          disabled={!selected}
          onValueChange={(endArrow) => update({ endArrow })}
        />
        <ArrowDownRight size={18} aria-hidden="true" />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="布局" priority="high">
        {layoutOptions.map((option) => {
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
      <WorkOfficeRibbonGroup label="大小">
        <ConnectorNumberField
          label="宽度（毫米）"
          ariaLabel="连接符宽度（毫米）"
          value={properties.width}
          min={DOCUMENT_CONNECTOR_LIMITS.width.min}
          max={DOCUMENT_CONNECTOR_LIMITS.width.max}
          disabled={!selected}
          onValueCommit={(width) => update({ width })}
        />
        <ConnectorNumberField
          label="高度（毫米）"
          ariaLabel="连接符高度（毫米）"
          value={properties.height}
          min={DOCUMENT_CONNECTOR_LIMITS.height.min}
          max={DOCUMENT_CONNECTOR_LIMITS.height.max}
          disabled={!selected}
          onValueCommit={(height) => update({ height })}
        />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="端点">
        <ConnectorNumberField
          label="起点 X（%）"
          ariaLabel="连接符起点 X（百分比）"
          value={properties.startX}
          min={DOCUMENT_CONNECTOR_LIMITS.endpoint.min}
          max={DOCUMENT_CONNECTOR_LIMITS.endpoint.max}
          disabled={!selected}
          onValueCommit={(startX) => update({ startX })}
        />
        <ConnectorNumberField
          label="起点 Y（%）"
          ariaLabel="连接符起点 Y（百分比）"
          value={properties.startY}
          min={DOCUMENT_CONNECTOR_LIMITS.endpoint.min}
          max={DOCUMENT_CONNECTOR_LIMITS.endpoint.max}
          disabled={!selected}
          onValueCommit={(startY) => update({ startY })}
        />
        <ConnectorNumberField
          label="终点 X（%）"
          ariaLabel="连接符终点 X（百分比）"
          value={properties.endX}
          min={DOCUMENT_CONNECTOR_LIMITS.endpoint.min}
          max={DOCUMENT_CONNECTOR_LIMITS.endpoint.max}
          disabled={!selected}
          onValueCommit={(endX) => update({ endX })}
        />
        <ConnectorNumberField
          label="终点 Y（%）"
          ariaLabel="连接符终点 Y（百分比）"
          value={properties.endY}
          min={DOCUMENT_CONNECTOR_LIMITS.endpoint.min}
          max={DOCUMENT_CONNECTOR_LIMITS.endpoint.max}
          disabled={!selected}
          onValueCommit={(endY) => update({ endY })}
        />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="位置">
        <div className="work-office-field work-document-connector-offset-field">
          <span>水平相对于</span>
          <OfficeSelect<WorkDocumentConnectorReference>
            ariaLabel="连接符水平相对于"
            value={properties.horizontalReference}
            options={horizontalReferenceOptions}
            disabled={!selected || properties.layout !== 'floating'}
            onValueChange={(horizontalReference) =>
              update({ horizontalReference })
            }
          />
        </div>
        <div className="work-office-field work-document-connector-offset-field">
          <span>垂直相对于</span>
          <OfficeSelect<WorkDocumentConnectorVerticalReference>
            ariaLabel="连接符垂直相对于"
            value={properties.verticalReference}
            options={verticalReferenceOptions}
            disabled={!selected || properties.layout !== 'floating'}
            onValueChange={(verticalReference) => update({ verticalReference })}
          />
        </div>
        <ConnectorNumberField
          label="水平偏移（毫米）"
          ariaLabel="连接符水平偏移（毫米）"
          value={properties.horizontalOffset ?? 0}
          min={DOCUMENT_CONNECTOR_LIMITS.offset.min}
          max={DOCUMENT_CONNECTOR_LIMITS.offset.max}
          disabled={!selected || properties.layout !== 'floating'}
          onValueCommit={(horizontalOffset) => update({ horizontalOffset })}
        />
        <ConnectorNumberField
          label="垂直偏移（毫米）"
          ariaLabel="连接符垂直偏移（毫米）"
          value={properties.verticalOffset ?? 0}
          min={DOCUMENT_CONNECTOR_LIMITS.offset.min}
          max={DOCUMENT_CONNECTOR_LIMITS.offset.max}
          disabled={!selected || properties.layout !== 'floating'}
          onValueCommit={(verticalOffset) => update({ verticalOffset })}
        />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="连接符">
        <ConnectorButton
          label="删除连接符"
          disabled={!selected}
          onClick={() => editor.commands.deleteDocumentConnector()}
        >
          <Trash2 size={18} />
        </ConnectorButton>
        <ConnectorButton
          label="恢复默认样式"
          disabled={!selected}
          onClick={() =>
            update({
              width: DOCUMENT_CONNECTOR_DEFAULTS.width,
              height: DOCUMENT_CONNECTOR_DEFAULTS.height,
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

function widthOptionsForValue(value: string) {
  if (widthOptions.some((option) => option.value === value))
    return widthOptions;
  return [...widthOptions, { value, label: `${value} 毫米` }] as const;
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
