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
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

const layoutOptions = [
  { value: 'inline', label: '嵌入文字', icon: Rows3 },
  { value: 'floating', label: '浮于文字上方', icon: SquareDashedMousePointer },
] as const satisfies readonly {
  value: WorkDocumentTextBoxLayout;
  label: string;
  icon: typeof Rows3;
}[];

const shapeOptions = [
  { value: 'rectangle', label: '矩形' },
  { value: 'roundedRectangle', label: '圆角矩形' },
  { value: 'ellipse', label: '椭圆' },
  { value: 'diamond', label: '菱形' },
  { value: 'triangle', label: '三角形' },
] as const satisfies readonly {
  value: WorkDocumentShapeType;
  label: string;
}[];

const horizontalReferenceOptions = [
  { value: 'column', label: '栏' },
  { value: 'margin', label: '页边距' },
  { value: 'page', label: '页面' },
] as const satisfies readonly {
  value: WorkDocumentTextBoxHorizontalReference;
  label: string;
}[];

const verticalReferenceOptions = [
  { value: 'paragraph', label: '段落' },
  { value: 'margin', label: '页边距' },
  { value: 'page', label: '页面' },
] as const satisfies readonly {
  value: WorkDocumentTextBoxVerticalReference;
  label: string;
}[];

const verticalAlignOptions = [
  { value: 'top', label: '顶端', icon: AlignVerticalJustifyStart },
  { value: 'center', label: '居中', icon: AlignVerticalJustifyCenter },
  { value: 'bottom', label: '底端', icon: AlignVerticalJustifyEnd },
] as const satisfies readonly {
  value: WorkDocumentTextBoxVerticalAlign;
  label: string;
  icon: typeof AlignVerticalJustifyStart;
}[];

const borderWidthOptions = [
  { value: '0', label: '无边框' },
  { value: '0.35', label: '细（0.35 mm）' },
  { value: '0.7', label: '中（0.7 mm）' },
  { value: '1.4', label: '粗（1.4 mm）' },
] as const;

export function DocumentTextBoxRibbon({ editor }: { editor: Editor }) {
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
      <WorkOfficeRibbonGroup label="形状" priority="high">
        <OfficeSelect<WorkDocumentShapeType>
          ariaLabel="文本框形状"
          value={properties.shapeType}
          options={shapeOptions}
          disabled={!selected}
          onValueChange={(shapeType) => update({ shapeType })}
        />
        <Shapes size={18} aria-hidden="true" />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="布局" priority="high">
        {layoutOptions.map((option) => {
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
      <WorkOfficeRibbonGroup label="大小">
        <div className="work-office-field work-document-text-box-size-field">
          <span>宽度（毫米）</span>
          <CommittedOfficeNumberField
            ariaLabel="文本框宽度（毫米）"
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
          <span>高度（毫米）</span>
          <CommittedOfficeNumberField
            ariaLabel="文本框高度（毫米）"
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
      <WorkOfficeRibbonGroup label="位置">
        <div className="work-office-field work-document-text-box-position-field">
          <span>水平相对于</span>
          <OfficeSelect<WorkDocumentTextBoxHorizontalReference>
            ariaLabel="文本框水平相对于"
            value={properties.horizontalReference}
            options={horizontalReferenceOptions}
            disabled={!selected || properties.layout !== 'floating'}
            onValueChange={(horizontalReference) =>
              update({ horizontalReference })
            }
          />
        </div>
        <div className="work-office-field work-document-text-box-position-field">
          <span>垂直相对于</span>
          <OfficeSelect<WorkDocumentTextBoxVerticalReference>
            ariaLabel="文本框垂直相对于"
            value={properties.verticalReference}
            options={verticalReferenceOptions}
            disabled={!selected || properties.layout !== 'floating'}
            onValueChange={(verticalReference) => update({ verticalReference })}
          />
        </div>
        <div className="work-office-field work-document-text-box-offset-field">
          <span>水平偏移（毫米）</span>
          <CommittedOfficeNumberField
            ariaLabel="文本框水平偏移（毫米）"
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
          <span>垂直偏移（毫米）</span>
          <CommittedOfficeNumberField
            ariaLabel="文本框垂直偏移（毫米）"
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
      <WorkOfficeRibbonGroup label="文本对齐">
        {verticalAlignOptions.map((option) => {
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
      <WorkOfficeRibbonGroup label="填充和轮廓">
        <OfficeColorPicker
          ariaLabel="文本框填充颜色"
          triggerLabel="填充"
          value={properties.fill}
          disabled={!selected}
          resetAction={{
            kind: 'automatic',
            label: '无填充',
            onSelect: () => update({ fill: 'transparent' }),
          }}
          onValueChange={(fill) => update({ fill })}
        />
        <OfficeColorPicker
          ariaLabel="文本框边框颜色"
          triggerLabel="边框"
          value={
            properties.borderColor === 'none'
              ? '#ffffff'
              : properties.borderColor
          }
          disabled={!selected}
          resetAction={{
            kind: 'none',
            label: '无轮廓',
            onSelect: () => update({ borderColor: 'none', borderWidth: 0 }),
          }}
          onValueChange={(borderColor) => update({ borderColor })}
        />
        <OfficeSelect
          ariaLabel="文本框边框粗细"
          value={String(properties.borderWidth)}
          options={borderWidthOptionsForValue(String(properties.borderWidth))}
          disabled={!selected}
          onValueChange={(value) =>
            update({
              borderWidth: Number(value),
              ...(value === '0' ? { borderColor: 'none' } : {}),
            })
          }
        />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="文本框">
        <TextBoxButton
          label="删除文本框"
          disabled={!selected}
          onClick={() => editor.commands.deleteDocumentTextBox()}
        >
          <Trash2 size={18} />
        </TextBoxButton>
        <TextBoxButton
          label="恢复默认样式"
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

function borderWidthOptionsForValue(value: string) {
  if (borderWidthOptions.some((option) => option.value === value)) {
    return borderWidthOptions;
  }
  return [...borderWidthOptions, { value, label: `${value} 毫米` }] as const;
}
