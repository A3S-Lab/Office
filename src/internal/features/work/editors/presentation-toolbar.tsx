import {
  AlignCenter,
  AlignHorizontalSpaceBetween,
  AlignLeft,
  AlignRight,
  AlignVerticalSpaceBetween,
  ArrowDownToLine,
  ArrowUpToLine,
  BarChart3,
  Bold,
  ClipboardPaste,
  Copy,
  Grid2X2,
  Group,
  Image,
  Italic,
  LayoutTemplate,
  Link2,
  MessageSquarePlus,
  MessagesSquare,
  PanelsTopLeft,
  Play,
  Plus,
  Redo2,
  Scissors,
  Square,
  SquarePlay,
  Trash2,
  Type,
  Underline,
  Undo2,
  Ungroup,
} from 'lucide-react';
import { useEffect, useRef, useState } from 'react';
import {
  DOCUMENT_LINK_VALIDATION_MESSAGE,
  normalizeDocumentHref,
} from '../work-document-links';
import { workSlideAnimationForElement } from '../work-presentation-animation';
import type {
  WorkSlide,
  WorkSlideElement,
  WorkSlideTextAlign,
} from '../work-types';
import {
  OfficeColorPicker,
  OfficeNumberField,
  OfficeSelect,
  useOfficeDialog,
} from './office-controls';
import {
  normalizeOfficeFontFamily,
  officeFontFamilies,
  officeFontFamilyLabel,
} from './office-font-families';
import { OfficeTableInsertPopover } from './office-table-insert-popover';
import type {
  PresentationEditorCanCommands,
  PresentationEditorCommands,
} from './presentation-command-types';
import { PresentationAnimationPanel } from './presentation-animation-panel';
import { PresentationTransitionPanel } from './presentation-transition-panel';
import {
  type WorkOfficeFileAction,
  WorkOfficeRibbon,
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

const presentationRibbonTabs = [
  { id: 'home', label: '开始' },
  { id: 'insert', label: '插入' },
  { id: 'design', label: '设计' },
  { id: 'transitions', label: '切换' },
  { id: 'animations', label: '动画' },
  { id: 'slideshow', label: '幻灯片放映', compactLabel: '放映' },
  { id: 'review', label: '审阅' },
  { id: 'view', label: '视图' },
] as const;

const basePresentationFontFamilyOptions = officeFontFamilies.map(
  ({ cssFamily, cssValue, group, label, name }) => ({
    value: cssValue,
    group,
    label,
    previewStyle: { fontFamily: cssFamily },
    searchText: `${name} ${label}`,
  }),
);

const presentationAlignmentOptions = [
  { value: 'none', label: '对象对齐', disabled: true },
  { value: 'left', label: '左对齐' },
  { value: 'center', label: '水平居中' },
  { value: 'right', label: '右对齐' },
  { value: 'top', label: '顶端对齐' },
  { value: 'middle', label: '垂直居中' },
  { value: 'bottom', label: '底端对齐' },
] as const;

export function PresentationToolbar({
  selectedSlide,
  selectedElement,
  selectedUnitCount,
  can,
  textFormattingAvailable,
  commentsOpen,
  commentCount,
  designOpen,
  editingDesign,
  background,
  transition,
  fileActions,
  viewMode = 'normal',
  commands,
}: {
  selectedSlide: WorkSlide;
  selectedElement: WorkSlideElement | null;
  selectedUnitCount: number;
  can: PresentationEditorCanCommands;
  textFormattingAvailable: boolean;
  commentsOpen: boolean;
  commentCount: number;
  designOpen: boolean;
  editingDesign: boolean;
  background?: string;
  transition: WorkSlide['transition'];
  fileActions?: readonly WorkOfficeFileAction[];
  viewMode?: 'normal' | 'sorter';
  commands: PresentationEditorCommands;
}) {
  const officeDialog = useOfficeDialog();
  const selectedFontSize = selectedElement
    ? String(selectedElement.fontSize)
    : '';
  const selectedFontSizeRef = useRef({
    elementId: selectedElement?.id ?? null,
    value: selectedFontSize,
  });
  const [fontSizeDraft, setFontSizeDraft] = useState(selectedFontSize);
  const fontFamilyValue = presentationFontFamilyValue(
    selectedElement?.fontFamily,
  );
  const selectedAnimation = workSlideAnimationForElement(
    selectedSlide,
    selectedElement?.id,
  );
  useEffect(() => {
    const previous = selectedFontSizeRef.current;
    const elementId = selectedElement?.id ?? null;
    selectedFontSizeRef.current = { elementId, value: selectedFontSize };
    setFontSizeDraft((draft) =>
      previous.elementId !== elementId || draft === previous.value
        ? selectedFontSize
        : draft,
    );
  }, [selectedElement?.id, selectedFontSize]);
  const commitFontSize = (value: string): void => {
    if (!selectedElement) return;
    const fontSize = normalizedPresentationFontSize(
      value,
      selectedElement.fontSize,
    );
    setFontSizeDraft(String(fontSize));
    if (fontSize === selectedElement.fontSize) return;
    commands.updateElement({ fontSize }, { restoreTextFocus: false });
  };
  return (
    <>
      <WorkOfficeRibbon
        ariaLabel="演示功能区"
        tabs={presentationRibbonTabs}
        defaultTab="home"
        fileActions={fileActions}
        className="work-presentation-ribbon"
        toolbarClassName="presentation-toolbar"
        panels={{
          home: (
            <>
              <WorkOfficeRibbonGroup label="撤销与恢复">
                <WorkOfficeRibbonButton
                  label="撤销"
                  title="撤销（Cmd/Ctrl+Z）"
                  aria-keyshortcuts="Control+Z Meta+Z"
                  disabled={!can.undo()}
                  onClick={commands.undo}
                >
                  <Undo2 size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label="重做"
                  title="重做（Cmd/Ctrl+Shift+Z 或 Cmd/Ctrl+Y）"
                  aria-keyshortcuts="Control+Shift+Z Meta+Shift+Z Control+Y Meta+Y"
                  disabled={!can.redo()}
                  onClick={commands.redo}
                >
                  <Redo2 size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label="幻灯片">
                <WorkOfficeRibbonButton
                  label="新建幻灯片"
                  title="新建幻灯片（Ctrl+M / ⌘⇧N）"
                  aria-keyshortcuts="Control+M Meta+Shift+N"
                  disabled={!can.addSlide()}
                  onClick={commands.addSlide}
                >
                  <Plus size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label="复制幻灯片"
                  disabled={!can.duplicateSlide()}
                  onClick={commands.duplicateSlide}
                >
                  <Copy size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label="删除幻灯片"
                  disabled={!can.deleteSlide()}
                  onClick={commands.deleteSlide}
                >
                  <Trash2 size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label="剪贴板">
                <WorkOfficeRibbonButton
                  label="复制"
                  title="复制（⌘/Ctrl+C）"
                  aria-keyshortcuts="Control+C Meta+C"
                  disabled={!can.copySelection()}
                  onClick={commands.copySelection}
                >
                  <Copy size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label="剪切"
                  title="剪切（⌘/Ctrl+X）"
                  aria-keyshortcuts="Control+X Meta+X"
                  disabled={!can.cutSelection()}
                  onClick={commands.cutSelection}
                >
                  <Scissors size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label="粘贴"
                  title="粘贴（⌘/Ctrl+V）"
                  aria-keyshortcuts="Control+V Meta+V"
                  disabled={!can.pasteSelection()}
                  onClick={commands.pasteSelection}
                >
                  <ClipboardPaste size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
              {selectedElement && (
                <>
                  {textFormattingAvailable && (
                    <WorkOfficeRibbonGroup label="字体">
                      <OfficeSelect
                        ariaLabel="演示字体"
                        value={fontFamilyValue}
                        options={presentationFontFamilyOptions(fontFamilyValue)}
                        onValueChange={(fontFamily) =>
                          commands.updateElement(
                            { fontFamily },
                            { restoreTextFocus: false },
                          )
                        }
                      />
                      <OfficeNumberField
                        ariaLabel="演示字号"
                        className="presentation-font-size-field"
                        min={8}
                        max={96}
                        step={1}
                        value={fontSizeDraft}
                        escapeConsumer={
                          fontSizeDraft !== String(selectedElement.fontSize)
                        }
                        onValueChange={setFontSizeDraft}
                        onCommit={commitFontSize}
                        onCancel={() =>
                          setFontSizeDraft(String(selectedElement.fontSize))
                        }
                      />
                      <WorkOfficeRibbonButton
                        label="加粗"
                        title="加粗（Cmd/Ctrl+B）"
                        aria-keyshortcuts="Control+B Meta+B"
                        displayLabel={false}
                        active={Boolean(selectedElement.bold)}
                        disabled={!can.toggleBold()}
                        onClick={commands.toggleBold}
                      >
                        <Bold size={15} />
                      </WorkOfficeRibbonButton>
                      <WorkOfficeRibbonButton
                        label="斜体"
                        title="斜体（Cmd/Ctrl+I）"
                        aria-keyshortcuts="Control+I Meta+I"
                        displayLabel={false}
                        active={Boolean(selectedElement.italic)}
                        disabled={!can.toggleItalic()}
                        onClick={commands.toggleItalic}
                      >
                        <Italic size={15} />
                      </WorkOfficeRibbonButton>
                      <WorkOfficeRibbonButton
                        label="下划线"
                        title="下划线（Cmd/Ctrl+U）"
                        aria-keyshortcuts="Control+U Meta+U"
                        displayLabel={false}
                        active={Boolean(selectedElement.underline)}
                        disabled={!can.toggleUnderline()}
                        onClick={commands.toggleUnderline}
                      >
                        <Underline size={15} />
                      </WorkOfficeRibbonButton>
                      {(
                        ['left', 'center', 'right'] as WorkSlideTextAlign[]
                      ).map((align) => (
                        <WorkOfficeRibbonButton
                          label={
                            align === 'left'
                              ? '左对齐'
                              : align === 'center'
                                ? '居中'
                                : '右对齐'
                          }
                          displayLabel={false}
                          active={selectedElement.align === align}
                          key={align}
                          onClick={() => commands.updateElement({ align })}
                        >
                          {align === 'left' ? (
                            <AlignLeft size={15} />
                          ) : align === 'center' ? (
                            <AlignCenter size={15} />
                          ) : (
                            <AlignRight size={15} />
                          )}
                        </WorkOfficeRibbonButton>
                      ))}
                      <OfficeColorPicker
                        compact
                        className="work-color-tool"
                        value={selectedElement.color}
                        ariaLabel="演示文字颜色"
                        onValueChange={(color) =>
                          commands.updateElement(
                            { color },
                            { restoreTextFocus: false },
                          )
                        }
                      />
                    </WorkOfficeRibbonGroup>
                  )}
                  <WorkOfficeRibbonGroup label="排列">
                    <OfficeSelect
                      ariaLabel={
                        selectedUnitCount > 1
                          ? '对齐所选对象'
                          : '元素对齐到幻灯片'
                      }
                      value="none"
                      options={presentationAlignmentOptions}
                      disabled={!can.alignElement('left')}
                      onValueChange={(alignment) => {
                        if (alignment === 'none') return;
                        commands.alignElement(alignment);
                      }}
                    />
                    {selectedUnitCount >= 3 && (
                      <>
                        <WorkOfficeRibbonButton
                          label="横向均匀分布"
                          displayLabel={false}
                          disabled={!can.distributeElements('horizontal')}
                          onClick={() =>
                            commands.distributeElements('horizontal')
                          }
                        >
                          <AlignHorizontalSpaceBetween size={17} />
                        </WorkOfficeRibbonButton>
                        <WorkOfficeRibbonButton
                          label="纵向均匀分布"
                          displayLabel={false}
                          disabled={!can.distributeElements('vertical')}
                          onClick={() =>
                            commands.distributeElements('vertical')
                          }
                        >
                          <AlignVerticalSpaceBetween size={17} />
                        </WorkOfficeRibbonButton>
                      </>
                    )}
                    <WorkOfficeRibbonButton
                      label="组合"
                      title="组合（⌘/Ctrl+G）"
                      aria-keyshortcuts="Control+G Meta+G"
                      displayLabel={false}
                      disabled={!can.groupElements()}
                      onClick={commands.groupElements}
                    >
                      <Group size={19} />
                    </WorkOfficeRibbonButton>
                    <WorkOfficeRibbonButton
                      label="取消组合"
                      title="取消组合（⌘/Ctrl+Shift+G）"
                      aria-keyshortcuts="Control+Shift+G Meta+Shift+G"
                      displayLabel={false}
                      disabled={!can.ungroupElements()}
                      onClick={commands.ungroupElements}
                    >
                      <Ungroup size={19} />
                    </WorkOfficeRibbonButton>
                    <WorkOfficeRibbonButton
                      label="下移一层"
                      displayLabel={false}
                      disabled={!can.reorderElement(-1)}
                      onClick={() => commands.reorderElement(-1)}
                    >
                      <ArrowDownToLine size={19} />
                    </WorkOfficeRibbonButton>
                    <WorkOfficeRibbonButton
                      label="上移一层"
                      displayLabel={false}
                      disabled={!can.reorderElement(1)}
                      onClick={() => commands.reorderElement(1)}
                    >
                      <ArrowUpToLine size={19} />
                    </WorkOfficeRibbonButton>
                  </WorkOfficeRibbonGroup>
                </>
              )}
            </>
          ),
          insert: (
            <>
              <WorkOfficeRibbonGroup label="文本与形状">
                <WorkOfficeRibbonButton
                  label="文本框"
                  disabled={!can.addElement('text')}
                  onClick={() => commands.addElement('text')}
                >
                  <Type size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label="形状"
                  disabled={!can.addElement('shape')}
                  onClick={() => commands.addElement('shape')}
                >
                  <Square size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label="内容">
                <WorkOfficeRibbonButton
                  label="图片"
                  disabled={!can.requestImage()}
                  onClick={commands.requestImage}
                >
                  <Image size={19} />
                </WorkOfficeRibbonButton>
                {!editingDesign && (
                  <>
                    <OfficeTableInsertPopover
                      label="表格"
                      disabled={!can.addTable({ rows: 1, columns: 1 })}
                      onInsert={commands.addTable}
                    />
                    <WorkOfficeRibbonButton
                      label="图表"
                      disabled={!can.addChart()}
                      onClick={commands.addChart}
                    >
                      <BarChart3 size={19} />
                    </WorkOfficeRibbonButton>
                  </>
                )}
              </WorkOfficeRibbonGroup>
              {selectedElement &&
                (selectedElement.type === 'text' ||
                  selectedElement.type === 'shape') && (
                  <WorkOfficeRibbonGroup label="链接">
                    <WorkOfficeRibbonButton
                      label="链接"
                      active={Boolean(selectedElement.href)}
                      onClick={() =>
                        void officeDialog
                          .prompt({
                            title: '链接地址',
                            description:
                              '为所选对象设置网页、邮箱或文档内链接。',
                            fieldLabel: '链接地址',
                            initialValue: selectedElement.href ?? 'https://',
                            placeholder: 'https://',
                            inputMode: 'url',
                            confirmLabel: '应用链接',
                            validate: (value) =>
                              value.trim() && !normalizeDocumentHref(value)
                                ? DOCUMENT_LINK_VALIDATION_MESSAGE
                                : null,
                          })
                          .then((href) => {
                            if (href !== null) {
                              const normalized = href.trim()
                                ? normalizeDocumentHref(href)
                                : undefined;
                              commands.updateElement({
                                href: normalized ?? undefined,
                              });
                            }
                          })
                      }
                    >
                      <Link2 size={19} />
                    </WorkOfficeRibbonButton>
                  </WorkOfficeRibbonGroup>
                )}
            </>
          ),
          design: (
            <>
              <WorkOfficeRibbonGroup label="母版">
                <WorkOfficeRibbonButton
                  label="母版和版式"
                  active={designOpen}
                  onClick={commands.toggleDesign}
                >
                  <LayoutTemplate size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label="背景">
                <OfficeColorPicker
                  compact
                  className="work-color-tool slide-background-tool"
                  value={background ?? selectedSlide.background}
                  ariaLabel={editingDesign ? '设计背景颜色' : '幻灯片背景颜色'}
                  onValueChange={commands.setBackground}
                />
              </WorkOfficeRibbonGroup>
            </>
          ),
          transitions: (
            <PresentationTransitionPanel
              slideId={selectedSlide.id}
              transition={transition}
              editable={can.setTransition(transition)}
              canApplyToAll={can.applyTransitionToAll}
              onChange={commands.setTransition}
              onApplyToAll={commands.applyTransitionToAll}
            />
          ),
          animations: (
            <PresentationAnimationPanel
              animation={selectedAnimation}
              canMove={can.moveEntranceAnimation}
              canPreview={can.previewAnimations()}
              editable={can.setEntranceAnimation(selectedAnimation?.effect)}
              onMove={commands.moveEntranceAnimation}
              onPreview={commands.previewAnimations}
              onSetEffect={commands.setEntranceAnimation}
              onUpdate={commands.updateEntranceAnimation}
            />
          ),
          slideshow: (
            <WorkOfficeRibbonGroup label="开始放映">
              <WorkOfficeRibbonButton
                label="从头开始放映"
                title="从头开始放映（F5）"
                aria-keyshortcuts="F5"
                data-presentation-slideshow-source="beginning"
                disabled={!can.startSlideshow('beginning')}
                onClick={() => commands.startSlideshow('beginning')}
              >
                <Play size={19} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label="从当前幻灯片放映"
                title="从当前幻灯片放映（Shift+F5）"
                aria-keyshortcuts="Shift+F5"
                data-presentation-slideshow-source="current"
                disabled={!can.startSlideshow('current')}
                onClick={() => commands.startSlideshow('current')}
              >
                <SquarePlay size={19} />
              </WorkOfficeRibbonButton>
            </WorkOfficeRibbonGroup>
          ),
          review: (
            <WorkOfficeRibbonGroup label="批注">
              <WorkOfficeRibbonButton
                label="新建批注"
                disabled={editingDesign || !can.addComment()}
                onClick={commands.addComment}
              >
                <MessageSquarePlus size={19} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label={`查看批注${commentCount ? `（${commentCount}）` : ''}`}
                disabled={editingDesign}
                active={commentsOpen}
                onClick={commands.toggleComments}
              >
                <MessagesSquare size={19} />
              </WorkOfficeRibbonButton>
            </WorkOfficeRibbonGroup>
          ),
          view: (
            <>
              <WorkOfficeRibbonGroup label="演示文稿视图">
                <WorkOfficeRibbonButton
                  label="普通视图"
                  active={viewMode === 'normal'}
                  onClick={() => commands.setViewMode('normal')}
                >
                  <PanelsTopLeft size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label="幻灯片浏览"
                  active={viewMode === 'sorter'}
                  onClick={() => commands.setViewMode('sorter')}
                >
                  <Grid2X2 size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label="母版">
                <WorkOfficeRibbonButton
                  label="母版视图"
                  active={designOpen}
                  onClick={commands.toggleDesign}
                >
                  <LayoutTemplate size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
            </>
          ),
        }}
      />
      {officeDialog.dialog}
    </>
  );
}

function normalizedPresentationFontSize(
  value: string,
  current: number,
): number {
  if (!value.trim()) return current;
  const number = Number(value);
  if (!Number.isFinite(number)) return current;
  return Math.min(96, Math.max(8, Math.round(number)));
}

function presentationFontFamilyValue(current: string | undefined): string {
  const value = current?.trim() || 'Aptos';
  const normalized = normalizePresentationFontFamily(value);
  return (
    basePresentationFontFamilyOptions.find(
      (option) => normalizePresentationFontFamily(option.value) === normalized,
    )?.value ?? value
  );
}

function presentationFontFamilyOptions(current: string) {
  if (
    basePresentationFontFamilyOptions.some((option) => option.value === current)
  ) {
    return basePresentationFontFamilyOptions;
  }
  return [
    ...basePresentationFontFamilyOptions,
    {
      value: current,
      group: '文档字体',
      label: officeFontFamilyLabel(current),
      previewStyle: { fontFamily: current },
      searchText: current,
    },
  ];
}

function normalizePresentationFontFamily(value: string): string {
  return normalizeOfficeFontFamily(value);
}
