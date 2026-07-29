import type { Editor } from '@tiptap/core';
import { NodeSelection } from '@tiptap/pm/state';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Rows3,
  ScanText,
  TextWrap,
  Trash2,
} from 'lucide-react';
import {
  documentImageAlternativeText,
  documentImageLayoutOptions,
  type WorkDocumentImageAlignment,
  type WorkDocumentImageLayout,
} from '../work-document-image-layout';
import { OfficeSelect, useOfficeDialog } from './office-controls';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

const imageWrapDistanceOptions = [
  { value: '0', label: '无间距' },
  { value: '2', label: '2 毫米' },
  { value: '3', label: '3 毫米' },
  { value: '5', label: '5 毫米' },
  { value: '10', label: '10 毫米' },
] as const;

export function DocumentPictureRibbon({ editor }: { editor: Editor }) {
  const image = documentImageLayoutOptions(editor);
  const imageSelected = editor.isActive('image');
  const wrapDistanceValue = String(image.wrapDistance);
  const officeDialog = useOfficeDialog();
  const updateLayout = (layout: WorkDocumentImageLayout) =>
    editor.commands.setDocumentImageLayoutOptions({ layout });
  const updateAlignment = (alignment: WorkDocumentImageAlignment) =>
    editor.commands.setDocumentImageLayoutOptions({ alignment });

  return (
    <>
      <WorkOfficeRibbonGroup label="文字环绕">
        <PictureButton
          label="嵌入文字"
          active={image.layout === 'inline'}
          disabled={!imageSelected}
          onClick={() => updateLayout('inline')}
        >
          <Rows3 size={18} />
        </PictureButton>
        <PictureButton
          label="四周环绕"
          active={image.layout === 'square'}
          disabled={!imageSelected}
          onClick={() => updateLayout('square')}
        >
          <TextWrap size={18} />
        </PictureButton>
        <PictureButton
          label="上下环绕"
          active={image.layout === 'topBottom'}
          disabled={!imageSelected}
          onClick={() => updateLayout('topBottom')}
        >
          <Rows3 size={18} />
        </PictureButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="位置">
        <PictureButton
          label="左对齐"
          active={image.alignment === 'left'}
          disabled={!imageSelected}
          onClick={() => updateAlignment('left')}
        >
          <AlignLeft size={18} />
        </PictureButton>
        <PictureButton
          label="居中"
          active={image.alignment === 'center'}
          disabled={!imageSelected}
          onClick={() => updateAlignment('center')}
        >
          <AlignCenter size={18} />
        </PictureButton>
        <PictureButton
          label="右对齐"
          active={image.alignment === 'right'}
          disabled={!imageSelected}
          onClick={() => updateAlignment('right')}
        >
          <AlignRight size={18} />
        </PictureButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="与文字距离">
        <OfficeSelect
          ariaLabel="图片与文字距离"
          value={wrapDistanceValue}
          options={imageWrapDistanceOptionsForValue(wrapDistanceValue)}
          disabled={!imageSelected || image.layout === 'inline'}
          onValueChange={(value) =>
            editor.commands.setDocumentImageLayoutOptions({
              wrapDistance: Number(value),
            })
          }
        />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="图片">
        <PictureButton
          label="替代文字"
          disabled={!imageSelected}
          onClick={() => {
            const imagePosition = selectedDocumentImagePosition(editor);
            void officeDialog
              .prompt({
                title: '图片说明',
                description: '简要描述图片中的关键信息，留空表示装饰性图片。',
                fieldLabel: '图片替代文字',
                initialValue: documentImageAlternativeText(editor),
                multiline: true,
                confirmLabel: '保存',
                restoreFocusTarget: () => editor.view.dom,
              })
              .then((value) => {
                if (!restoreDocumentImageSelection(editor, imagePosition)) {
                  return;
                }
                if (value !== null) {
                  editor.commands.setDocumentImageAlternativeText(value);
                }
              });
          }}
        >
          <ScanText size={18} />
        </PictureButton>
        <PictureButton
          label="删除图片"
          disabled={!imageSelected}
          onClick={() => editor.chain().focus().deleteSelection().run()}
        >
          <Trash2 size={18} />
        </PictureButton>
      </WorkOfficeRibbonGroup>
      {officeDialog.dialog}
    </>
  );
}

function selectedDocumentImagePosition(editor: Editor): number | null {
  const selection = editor.state.selection;
  return selection instanceof NodeSelection && selection.node.type.name === 'image'
    ? selection.from
    : null;
}

function restoreDocumentImageSelection(
  editor: Editor,
  position: number | null,
): boolean {
  if (
    position === null ||
    editor.isDestroyed ||
    editor.state.doc.nodeAt(position)?.type.name !== 'image'
  ) {
    return false;
  }
  return editor.chain().setNodeSelection(position).focus().run();
}

function imageWrapDistanceOptionsForValue(value: string) {
  if (imageWrapDistanceOptions.some((option) => option.value === value)) {
    return imageWrapDistanceOptions;
  }
  return [...imageWrapDistanceOptions, { value, label: `${value} 毫米` }];
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
