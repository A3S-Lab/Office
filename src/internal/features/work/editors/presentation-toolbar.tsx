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
  Table2,
  Trash2,
  Type,
  Undo2,
  Underline,
  Ungroup,
} from 'lucide-react';
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
import type {
  PresentationEditorCanCommands,
  PresentationEditorCommands,
} from './presentation-command-types';
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
  { id: 'slideshow', label: '幻灯片放映' },
  { id: 'review', label: '审阅' },
  { id: 'view', label: '视图' },
] as const;

const presentationFontFamilyOptions = [
  { value: 'Aptos', label: 'Aptos' },
  { value: '"Microsoft YaHei"', label: '微软雅黑' },
  { value: 'SimSun', label: '宋体' },
  { value: 'Arial', label: 'Arial' },
  { value: '"Times New Roman"', label: 'Times New Roman' },
] as const;

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
                  disabled={!can.undo()}
                  onClick={commands.undo}
                >
                  <Undo2 size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label="重做"
                  title="重做（Cmd/Ctrl+Shift+Z）"
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
                  disabled={!can.copySelection()}
                  onClick={commands.copySelection}
                >
                  <Copy size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label="剪切"
                  title="剪切（⌘/Ctrl+X）"
                  disabled={!can.cutSelection()}
                  onClick={commands.cutSelection}
                >
                  <Scissors size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label="粘贴"
                  title="粘贴（⌘/Ctrl+V）"
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
                        value={selectedElement.fontFamily ?? 'Aptos'}
                        options={presentationFontFamilyOptions}
                        onValueChange={(fontFamily) =>
                          commands.updateElement(
                            { fontFamily },
                            { restoreTextFocus: false },
                          )
                        }
                      />
                      <div className="presentation-number-field work-office-field">
                        <span>字号</span>
                        <OfficeNumberField
                          ariaLabel="演示字号"
                          min={8}
                          max={96}
                          value={selectedElement.fontSize}
                          onValueChange={(value) =>
                            commands.updateElement(
                              { fontSize: Number(value) || 8 },
                              { restoreTextFocus: false },
                            )
                          }
                        />
                      </div>
                      <WorkOfficeRibbonButton
                        label="加粗"
                        title="加粗（Cmd/Ctrl+B）"
                        aria-keyshortcuts="Control+B Meta+B"
                        displayLabel={false}
                        active={Boolean(selectedElement.bold)}
                        onClick={() =>
                          commands.updateElement({
                            bold: !selectedElement.bold,
                          })
                        }
                      >
                        <Bold size={15} />
                      </WorkOfficeRibbonButton>
                      <WorkOfficeRibbonButton
                        label="斜体"
                        title="斜体（Cmd/Ctrl+I）"
                        aria-keyshortcuts="Control+I Meta+I"
                        displayLabel={false}
                        active={Boolean(selectedElement.italic)}
                        onClick={() =>
                          commands.updateElement({
                            italic: !selectedElement.italic,
                          })
                        }
                      >
                        <Italic size={15} />
                      </WorkOfficeRibbonButton>
                      <WorkOfficeRibbonButton
                        label="下划线"
                        title="下划线（Cmd/Ctrl+U）"
                        aria-keyshortcuts="Control+U Meta+U"
                        displayLabel={false}
                        active={Boolean(selectedElement.underline)}
                        onClick={() =>
                          commands.updateElement({
                            underline: !selectedElement.underline,
                          })
                        }
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
                      disabled={!can.groupElements()}
                      onClick={commands.groupElements}
                    >
                      <Group size={19} />
                    </WorkOfficeRibbonButton>
                    <WorkOfficeRibbonButton
                      label="取消组合"
                      title="取消组合（⌘/Ctrl+Shift+G）"
                      aria-keyshortcuts="Control+Shift+G Meta+Shift+G"
                      disabled={!can.ungroupElements()}
                      onClick={commands.ungroupElements}
                    >
                      <Ungroup size={19} />
                    </WorkOfficeRibbonButton>
                    <WorkOfficeRibbonButton
                      label="下移一层"
                      disabled={!can.reorderElement(-1)}
                      onClick={() => commands.reorderElement(-1)}
                    >
                      <ArrowDownToLine size={19} />
                    </WorkOfficeRibbonButton>
                    <WorkOfficeRibbonButton
                      label="上移一层"
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
                    <WorkOfficeRibbonButton
                      label="表格"
                      disabled={!can.addTable()}
                      onClick={commands.addTable}
                    >
                      <Table2 size={19} />
                    </WorkOfficeRibbonButton>
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
                            initialValue: selectedElement.href ?? 'https://',
                            placeholder: 'https://',
                            confirmLabel: '应用链接',
                          })
                          .then((href) => {
                            if (href !== null)
                              commands.updateElement({
                                href: href.trim() || undefined,
                              });
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
              transition={transition}
              onChange={commands.setTransition}
              onApplyToAll={commands.applyTransitionToAll}
            />
          ),
          slideshow: (
            <WorkOfficeRibbonGroup label="开始放映">
              <WorkOfficeRibbonButton
                label="从头开始放映"
                title="从头开始放映（F5）"
                aria-keyshortcuts="F5"
                disabled={!can.startSlideshow()}
                onClick={commands.startSlideshow}
              >
                <Play size={19} />
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
