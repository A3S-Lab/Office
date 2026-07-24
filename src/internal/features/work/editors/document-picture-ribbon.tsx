import type { Editor } from '@tiptap/core';
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
  setDocumentImageAlternativeText,
  setDocumentImageLayoutOptions,
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
  const officeDialog = useOfficeDialog();
  const updateLayout = (layout: WorkDocumentImageLayout) =>
    setDocumentImageLayoutOptions(editor, { layout });
  const updateAlignment = (alignment: WorkDocumentImageAlignment) =>
    setDocumentImageLayoutOptions(editor, { alignment });

  return (
    <>
      <WorkOfficeRibbonGroup label="文字环绕">
        <PictureButton
          label="嵌入文字"
          active={image.layout === 'inline'}
          onClick={() => updateLayout('inline')}
        >
          <Rows3 size={18} />
        </PictureButton>
        <PictureButton
          label="四周环绕"
          active={image.layout === 'square'}
          onClick={() => updateLayout('square')}
        >
          <TextWrap size={18} />
        </PictureButton>
        <PictureButton
          label="上下环绕"
          active={image.layout === 'topBottom'}
          onClick={() => updateLayout('topBottom')}
        >
          <Rows3 size={18} />
        </PictureButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="位置">
        <PictureButton
          label="左对齐"
          active={image.alignment === 'left'}
          onClick={() => updateAlignment('left')}
        >
          <AlignLeft size={18} />
        </PictureButton>
        <PictureButton
          label="居中"
          active={image.alignment === 'center'}
          onClick={() => updateAlignment('center')}
        >
          <AlignCenter size={18} />
        </PictureButton>
        <PictureButton
          label="右对齐"
          active={image.alignment === 'right'}
          onClick={() => updateAlignment('right')}
        >
          <AlignRight size={18} />
        </PictureButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="与文字距离">
        <OfficeSelect
          ariaLabel="图片与文字距离"
          value={String(image.wrapDistance)}
          options={imageWrapDistanceOptions}
          onValueChange={(value) =>
            setDocumentImageLayoutOptions(editor, {
              wrapDistance: Number(value),
            })
          }
        />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="图片">
        <PictureButton
          label="替代文字"
          onClick={() => {
            void officeDialog
              .prompt({
                title: '图片替代文字',
                description: '简要说明图片内容，便于读屏软件识别。',
                initialValue: documentImageAlternativeText(editor),
                confirmLabel: '保存',
              })
              .then((value) => {
                if (value !== null)
                  setDocumentImageAlternativeText(editor, value);
              });
          }}
        >
          <ScanText size={18} />
        </PictureButton>
        <PictureButton
          label="删除图片"
          onClick={() => editor.chain().focus().deleteSelection().run()}
        >
          <Trash2 size={18} />
        </PictureButton>
      </WorkOfficeRibbonGroup>
      {officeDialog.dialog}
    </>
  );
}

function PictureButton({
  label,
  active = false,
  onClick,
  children,
}: {
  label: string;
  active?: boolean;
  onClick: () => void;
  children: React.ReactNode;
}) {
  return (
    <WorkOfficeRibbonButton
      label={label}
      displayLabel
      active={active}
      onClick={onClick}
    >
      {children}
    </WorkOfficeRibbonButton>
  );
}
