import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowDownToLine,
  ArrowUpToLine,
  BarChart3,
  Bold,
  ClipboardPaste,
  Copy,
  Grid2X2,
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
import type { PresentationCommandDispatcher } from './presentation-command-controller';
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
  { value: 'none', label: '对齐到幻灯片', disabled: true },
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
  slideCount,
  canUndo,
  canRedo,
  commentsOpen,
  commentCount,
  designOpen,
  editingDesign,
  background,
  transition,
  canStartSlideshow,
  fileActions,
  viewMode = 'normal',
  onCommand,
}: {
  selectedSlide: WorkSlide;
  selectedElement: WorkSlideElement | null;
  slideCount: number;
  canUndo: boolean;
  canRedo: boolean;
  commentsOpen: boolean;
  commentCount: number;
  designOpen: boolean;
  editingDesign: boolean;
  background?: string;
  transition: WorkSlide['transition'];
  canStartSlideshow: boolean;
  fileActions?: readonly WorkOfficeFileAction[];
  viewMode?: 'normal' | 'sorter';
  onCommand: PresentationCommandDispatcher;
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
                  disabled={!canUndo}
                  onClick={() => onCommand({ type: 'history.undo' })}
                >
                  <Undo2 size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label="重做"
                  title="重做（Cmd/Ctrl+Shift+Z）"
                  disabled={!canRedo}
                  onClick={() => onCommand({ type: 'history.redo' })}
                >
                  <Redo2 size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label="幻灯片">
                <WorkOfficeRibbonButton
                  label="新建幻灯片"
                  title="新建幻灯片（Ctrl+M / ⌘⇧N）"
                  aria-keyshortcuts="Control+M Meta+Shift+N"
                  onClick={() => onCommand({ type: 'slide.add' })}
                >
                  <Plus size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label="复制幻灯片"
                  onClick={() => onCommand({ type: 'slide.duplicate' })}
                >
                  <Copy size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label="删除幻灯片"
                  disabled={slideCount === 1}
                  onClick={() => onCommand({ type: 'slide.delete' })}
                >
                  <Trash2 size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label="剪贴板">
                <WorkOfficeRibbonButton
                  label="复制"
                  title="复制（⌘/Ctrl+C）"
                  onClick={() => onCommand({ type: 'clipboard.copy' })}
                >
                  <Copy size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label="剪切"
                  title="剪切（⌘/Ctrl+X）"
                  onClick={() => onCommand({ type: 'clipboard.cut' })}
                >
                  <Scissors size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label="粘贴"
                  title="粘贴（⌘/Ctrl+V）"
                  onClick={() => onCommand({ type: 'clipboard.paste' })}
                >
                  <ClipboardPaste size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
              {selectedElement && (
                <>
                  {isPresentationTextElement(selectedElement) && (
                    <WorkOfficeRibbonGroup label="字体">
                      <OfficeSelect
                        ariaLabel="演示字体"
                        value={selectedElement.fontFamily ?? 'Aptos'}
                        options={presentationFontFamilyOptions}
                        onValueChange={(fontFamily) =>
                          onCommand({
                            type: 'element.update',
                            patch: { fontFamily },
                            restoreTextFocus: false,
                          })
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
                            onCommand({
                              type: 'element.update',
                              patch: { fontSize: Number(value) || 8 },
                              restoreTextFocus: false,
                            })
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
                          onCommand({
                            type: 'element.update',
                            patch: { bold: !selectedElement.bold },
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
                          onCommand({
                            type: 'element.update',
                            patch: { italic: !selectedElement.italic },
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
                          onCommand({
                            type: 'element.update',
                            patch: {
                              underline: !selectedElement.underline,
                            },
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
                          onClick={() =>
                            onCommand({
                              type: 'element.update',
                              patch: { align },
                            })
                          }
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
                          onCommand({
                            type: 'element.update',
                            patch: { color },
                            restoreTextFocus: false,
                          })
                        }
                      />
                    </WorkOfficeRibbonGroup>
                  )}
                  <WorkOfficeRibbonGroup label="排列">
                    <OfficeSelect
                      ariaLabel="元素对齐到幻灯片"
                      value="none"
                      options={presentationAlignmentOptions}
                      onValueChange={(alignment) => {
                        if (alignment === 'none') return;
                        onCommand({
                          type: 'element.align',
                          alignment,
                        });
                      }}
                    />
                    <WorkOfficeRibbonButton
                      label="下移一层"
                      onClick={() =>
                        onCommand({
                          type: 'element.reorder',
                          direction: -1,
                        })
                      }
                    >
                      <ArrowDownToLine size={19} />
                    </WorkOfficeRibbonButton>
                    <WorkOfficeRibbonButton
                      label="上移一层"
                      onClick={() =>
                        onCommand({
                          type: 'element.reorder',
                          direction: 1,
                        })
                      }
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
                  onClick={() =>
                    onCommand({
                      type: 'element.add',
                      elementType: 'text',
                    })
                  }
                >
                  <Type size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label="形状"
                  onClick={() =>
                    onCommand({
                      type: 'element.add',
                      elementType: 'shape',
                    })
                  }
                >
                  <Square size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label="内容">
                <WorkOfficeRibbonButton
                  label="图片"
                  onClick={() => onCommand({ type: 'image.request' })}
                >
                  <Image size={19} />
                </WorkOfficeRibbonButton>
                {!editingDesign && (
                  <>
                    <WorkOfficeRibbonButton
                      label="表格"
                      onClick={() => onCommand({ type: 'table.add' })}
                    >
                      <Table2 size={19} />
                    </WorkOfficeRibbonButton>
                    <WorkOfficeRibbonButton
                      label="图表"
                      onClick={() => onCommand({ type: 'chart.add' })}
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
                              onCommand({
                                type: 'element.update',
                                patch: {
                                  href: href.trim() || undefined,
                                },
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
                  onClick={() => onCommand({ type: 'design.toggle' })}
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
                  onValueChange={(color) =>
                    onCommand({ type: 'slide.background.set', color })
                  }
                />
              </WorkOfficeRibbonGroup>
            </>
          ),
          transitions: (
            <PresentationTransitionPanel
              transition={transition}
              onChange={(nextTransition) =>
                onCommand({
                  type: 'transition.set',
                  transition: nextTransition,
                })
              }
              onApplyToAll={() => onCommand({ type: 'transition.applyToAll' })}
            />
          ),
          slideshow: (
            <WorkOfficeRibbonGroup label="开始放映">
              <WorkOfficeRibbonButton
                label="从头开始放映"
                title="从头开始放映（F5）"
                aria-keyshortcuts="F5"
                disabled={!canStartSlideshow}
                onClick={() => onCommand({ type: 'slideshow.start' })}
              >
                <Play size={19} />
              </WorkOfficeRibbonButton>
            </WorkOfficeRibbonGroup>
          ),
          review: (
            <WorkOfficeRibbonGroup label="批注">
              <WorkOfficeRibbonButton
                label="新建批注"
                disabled={editingDesign}
                onClick={() => onCommand({ type: 'comment.add' })}
              >
                <MessageSquarePlus size={19} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label={`查看批注${commentCount ? `（${commentCount}）` : ''}`}
                disabled={editingDesign}
                active={commentsOpen}
                onClick={() => onCommand({ type: 'comments.toggle' })}
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
                  onClick={() =>
                    onCommand({ type: 'view.set', mode: 'normal' })
                  }
                >
                  <PanelsTopLeft size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label="幻灯片浏览"
                  active={viewMode === 'sorter'}
                  onClick={() =>
                    onCommand({ type: 'view.set', mode: 'sorter' })
                  }
                >
                  <Grid2X2 size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label="母版">
                <WorkOfficeRibbonButton
                  label="母版视图"
                  active={designOpen}
                  onClick={() => onCommand({ type: 'design.toggle' })}
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

function isPresentationTextElement(element: WorkSlideElement): boolean {
  return (
    element.type === 'text' ||
    element.type === 'shape' ||
    Boolean(element.text || element.textRuns?.length)
  );
}
