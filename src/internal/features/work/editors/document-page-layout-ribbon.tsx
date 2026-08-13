import type { Editor } from '@tiptap/core';
import { Columns3, FilePlus2, Palette, Settings2 } from 'lucide-react';
import { normalizeDocumentColumns } from '../work-document-columns';
import {
  updateDocumentPageOrientation,
  updateDocumentPaperSizePreset,
} from '../work-document-page-size';
import type {
  WorkDocumentMargins,
  WorkDocumentSectionLayout,
} from '../work-types';
import { getDocumentCommandDefinition } from './document-command-catalog';
import type { DocumentLayoutPanelTab } from './document-layout-panel';
import { DocumentPaginationPopover } from './document-pagination-popover';
import { DocumentParagraphSpacingPopover } from './document-paragraph-spacing-popover';
import { OfficeColorPicker, OfficeSelect } from './office-controls';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

type DocumentMarginPreset =
  | 'custom'
  | 'moderate'
  | 'narrow'
  | 'normal'
  | 'wide';
type DocumentColumnPreset = '1' | '2' | '3' | 'more';

const documentMarginPresets = {
  normal: { top: 25.4, right: 25.4, bottom: 25.4, left: 25.4 },
  narrow: { top: 12.7, right: 12.7, bottom: 12.7, left: 12.7 },
  moderate: { top: 25.4, right: 19.1, bottom: 25.4, left: 19.1 },
  wide: { top: 25.4, right: 50.8, bottom: 25.4, left: 50.8 },
} as const satisfies Record<
  Exclude<DocumentMarginPreset, 'custom'>,
  WorkDocumentMargins
>;

export function DocumentPageLayoutRibbon({
  editor,
  layout,
  layoutOpen,
  pageColor,
  onLayoutChange,
  onOpenLayout,
  onToggleLayout,
  onPageColorChange,
  onInsertSection,
}: {
  editor: Editor;
  layout: WorkDocumentSectionLayout;
  layoutOpen: boolean;
  pageColor: string;
  onLayoutChange: (layout: WorkDocumentSectionLayout) => void;
  onOpenLayout: (target: DocumentLayoutPanelTab) => void;
  onToggleLayout: () => void;
  onPageColorChange: (color: string) => void;
  onInsertSection: () => void;
}) {
  const pageBreakCommand = getDocumentCommandDefinition('insertPageBreak');
  const marginPreset = documentMarginPreset(layout.margins);
  const columnPreset: DocumentColumnPreset =
    layout.columns.count >= 1 && layout.columns.count <= 3
      ? (String(layout.columns.count) as DocumentColumnPreset)
      : 'more';
  const update = (patch: Partial<WorkDocumentSectionLayout>) =>
    onLayoutChange({ ...layout, ...patch });

  return (
    <>
      <WorkOfficeRibbonGroup label="页面设置" priority="high">
        <div className="work-office-field work-document-page-setup-choice">
          <span>页边距</span>
          <OfficeSelect
            ariaLabel="页边距"
            value={marginPreset}
            options={[
              { value: 'normal', label: '普通', meta: '四边 25.4 mm' },
              { value: 'narrow', label: '窄', meta: '四边 12.7 mm' },
              {
                value: 'moderate',
                label: '适中',
                meta: '上下 25.4 · 左右 19.1 mm',
              },
              {
                value: 'wide',
                label: '宽',
                meta: '上下 25.4 · 左右 50.8 mm',
              },
              { value: 'custom', label: '自定义页边距', meta: '页面设置' },
            ]}
            onValueChange={(preset) => {
              if (preset === 'custom') {
                onOpenLayout('page');
                return;
              }
              update({ margins: { ...documentMarginPresets[preset] } });
            }}
          />
        </div>
        <div className="work-office-field work-document-page-setup-choice">
          <span>方向</span>
          <OfficeSelect
            ariaLabel="页面方向"
            value={layout.orientation}
            options={[
              { value: 'portrait', label: '纵向' },
              { value: 'landscape', label: '横向' },
            ]}
            onValueChange={(orientation) =>
              onLayoutChange(updateDocumentPageOrientation(layout, orientation))
            }
          />
        </div>
        <div className="work-office-field work-document-page-setup-choice">
          <span>纸张</span>
          <OfficeSelect
            ariaLabel="纸张大小"
            value={layout.pageSize}
            options={[
              { value: 'a3', label: 'A3' },
              { value: 'a4', label: 'A4' },
              { value: 'a5', label: 'A5' },
              { value: 'letter', label: 'Letter' },
              { value: 'legal', label: 'Legal' },
              { value: 'tabloid', label: 'Tabloid' },
              { value: 'custom', label: '自定义' },
            ]}
            onValueChange={(pageSize) => {
              if (pageSize === 'custom') {
                onOpenLayout('page');
                return;
              }
              onLayoutChange(updateDocumentPaperSizePreset(layout, pageSize));
            }}
          />
        </div>
        <div className="work-office-field work-document-page-setup-choice">
          <span>分栏</span>
          <OfficeSelect
            ariaLabel="分栏"
            value={columnPreset}
            options={[
              { value: '1', label: '一栏' },
              { value: '2', label: '两栏' },
              { value: '3', label: '三栏' },
              { value: 'more', label: '更多分栏' },
            ]}
            onValueChange={(preset) => {
              if (preset === 'more') {
                onOpenLayout('columns');
                return;
              }
              update({
                columns: normalizeDocumentColumns({
                  ...layout.columns,
                  count: Number(preset),
                  custom: undefined,
                }),
              });
            }}
          />
        </div>
        <WorkOfficeRibbonButton
          label="页面设置"
          active={layoutOpen}
          onClick={onToggleLayout}
        >
          <Settings2 size={19} />
        </WorkOfficeRibbonButton>
        <WorkOfficeRibbonButton
          label={pageBreakCommand.label}
          title={`${pageBreakCommand.label}（${pageBreakCommand.shortcut?.label}）`}
          aria-keyshortcuts={pageBreakCommand.shortcut?.aria}
          onClick={() =>
            editor.chain().focus().insertContent({ type: 'pageBreak' }).run()
          }
        >
          <FilePlus2 size={19} />
        </WorkOfficeRibbonButton>
        <WorkOfficeRibbonButton label="插入分节符" onClick={onInsertSection}>
          <Columns3 size={19} />
        </WorkOfficeRibbonButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="段落" priority="high">
        <DocumentParagraphSpacingPopover editor={editor} />
        <DocumentPaginationPopover editor={editor} />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="页面背景" priority="low">
        <OfficeColorPicker
          ariaLabel="页面颜色"
          className="work-document-page-color-picker"
          triggerLabel="页面颜色"
          triggerIcon={<Palette size={18} />}
          value={pageColor}
          onValueChange={onPageColorChange}
        />
      </WorkOfficeRibbonGroup>
    </>
  );
}

function documentMarginPreset(
  margins: WorkDocumentMargins,
): DocumentMarginPreset {
  for (const [preset, values] of Object.entries(documentMarginPresets)) {
    if (
      Object.keys(values).every(
        (side) =>
          Math.abs(
            margins[side as keyof WorkDocumentMargins] -
              values[side as keyof WorkDocumentMargins],
          ) < 0.05,
      )
    ) {
      return preset as DocumentMarginPreset;
    }
  }
  return 'custom';
}
