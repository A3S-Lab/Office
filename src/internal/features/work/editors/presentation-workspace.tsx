import type { Editor } from '@tiptap/core';
import { GalleryVerticalEnd } from 'lucide-react';
import {
  useId,
  useRef,
  useState,
  type PointerEvent,
  type RefObject,
} from 'react';
import { useDialogFocusScope } from '../../../design-system/primitives/overlay/dialog-focus-scope';
import type { OfficeKernelPresentationSnapGuide } from '../../../kernel/office-kernel-protocol';
import {
  isWorkspaceContextMenuKeyboardEvent,
  type WorkspaceContextMenuEvent,
} from '../../workspace/components/workspace-context-menu';
import type {
  WorkPresentationContent,
  WorkPresentationLayout,
  WorkPresentationMaster,
  WorkSlide,
  WorkSlideElement,
} from '../work-types';
import { presentationSelectionUnits } from '../work-presentation-groups';
import { OfficeTextArea } from './office-controls';
import { SlideChart } from './presentation-chart-canvas';
import type { PresentationEditorCommands } from './presentation-command-types';
import type { PresentationDesignMode } from './presentation-editor-types';
import {
  EditableSlideTable,
  SlideElementPreview,
  SlideElementTextPreview,
  SlideTablePreview,
  slideElementStyle,
} from './presentation-slide-canvas';
import {
  presentationElementCanEditContent,
  presentationSelectionBounds,
  selectedPresentationElements,
} from './presentation-selection';
import { PresentationThumbnailRail } from './presentation-thumbnail-rail';
import { PresentationTextEditor } from './presentation-text-editor';

export type PresentationWorkspaceCommands = Pick<
  PresentationEditorCommands,
  | 'addSlide'
  | 'deleteSlideById'
  | 'editElement'
  | 'exitEditing'
  | 'instantiatePlaceholder'
  | 'openComment'
  | 'selectElement'
  | 'selectSlide'
  | 'setViewMode'
  | 'updateElement'
  | 'updateNotes'
  | 'updateTextElement'
>;

export interface PresentationWorkspaceProps {
  activeBackground: string;
  activeCommentId: string | null;
  activeElements: WorkSlideElement[];
  aspectRatio: string;
  canvasName: string;
  canvasRef: RefObject<HTMLElement | null>;
  commands: PresentationWorkspaceCommands;
  content: WorkPresentationContent;
  designContent: WorkPresentationContent;
  designMode: PresentationDesignMode;
  inheritedElements: WorkSlideElement[];
  placeholderGuides: WorkSlideElement[];
  editingElementId: string | null;
  selectedElementIds: readonly string[];
  selectedLayout: WorkPresentationLayout | undefined;
  selectedMaster: WorkPresentationMaster | undefined;
  selectedSlide: WorkSlide;
  snapGuides: OfficeKernelPresentationSnapGuide[];
  viewMode: 'normal' | 'sorter';
  zoom: number;
  onBeginDrag: (
    event: PointerEvent,
    element: WorkSlideElement,
    mode: 'move' | 'resize',
  ) => void;
  onContinueDrag: (event: PointerEvent) => void;
  onDragCancel: () => void;
  onDragEnd: (event: PointerEvent) => void;
  onOpenContextMenu: (
    event: WorkspaceContextMenuEvent,
    slide: WorkSlide,
    slideIndex: number,
    element?: WorkSlideElement | null,
  ) => void;
  onTextEditorChange: (elementId: string, editor: Editor | null) => void;
  onTextSelectionChange: () => void;
}

export function PresentationWorkspace({
  activeBackground,
  activeCommentId,
  activeElements,
  aspectRatio,
  canvasName,
  canvasRef,
  commands,
  content,
  designContent,
  designMode,
  inheritedElements,
  placeholderGuides,
  editingElementId,
  selectedElementIds,
  selectedLayout,
  selectedMaster,
  selectedSlide,
  snapGuides,
  viewMode,
  zoom,
  onBeginDrag,
  onContinueDrag,
  onDragCancel,
  onDragEnd,
  onOpenContextMenu,
  onTextEditorChange,
  onTextSelectionChange,
}: PresentationWorkspaceProps) {
  const [mobileSlideNavigationOpen, setMobileSlideNavigationOpen] =
    useState(false);
  const mobileSlideNavigationId = useId();
  const mobileSlideNavigationToggleRef = useRef<HTMLButtonElement>(null);
  const mobileSlideNavigationCloseRef = useRef<HTMLButtonElement>(null);

  const closeMobileSlideNavigation = () => {
    setMobileSlideNavigationOpen(false);
  };
  useDialogFocusScope<HTMLElement>({
    active: mobileSlideNavigationOpen,
    onEscape: closeMobileSlideNavigation,
    initialFocus: () => mobileSlideNavigationCloseRef.current,
    getActiveScope: () =>
      document.getElementById(mobileSlideNavigationId) as HTMLElement | null,
    getIsolationExceptions: () => [
      document.querySelector<HTMLElement>(
        '.work-presentation-slide-navigation-backdrop',
      ),
    ],
    restoreFocusTarget: () => mobileSlideNavigationToggleRef.current,
  });

  const selectedElementSet = new Set(selectedElementIds);
  const selectedElements = selectedPresentationElements(
    activeElements,
    selectedElementIds,
  );
  const selectionUnits = presentationSelectionUnits(
    activeElements,
    selectedElementIds,
  );
  const selectionBounds = presentationSelectionBounds(selectedElements);
  const selectionTransformAnchor = selectedElements.at(-1);
  const selectionResizeLabel =
    selectionUnits.length === 1 && selectionUnits[0]?.groupId
      ? '缩放所选组合'
      : '缩放所选对象';
  if (viewMode === 'sorter') {
    return (
      <PresentationThumbnailRail
        aspectRatio={aspectRatio}
        content={content}
        designContent={designContent}
        selectedSlide={selectedSlide}
        viewMode={viewMode}
        zoom={zoom}
        onAddSlide={commands.addSlide}
        onDeleteSlide={commands.deleteSlideById}
        onOpenContextMenu={onOpenContextMenu}
        onSelectSlide={commands.selectSlide}
        onViewModeChange={commands.setViewMode}
      />
    );
  }

  const selectedSlideIndex = content.slides.findIndex(
    (slide) => slide.id === selectedSlide.id,
  );
  const selectSlide = (slideId: string, returnToSlideMode: boolean) => {
    commands.selectSlide(slideId, returnToSlideMode);
    if (mobileSlideNavigationOpen) closeMobileSlideNavigation();
  };
  return (
    <div
      className="work-presentation-layout"
      data-mobile-slide-navigation={
        mobileSlideNavigationOpen ? 'open' : 'closed'
      }
    >
      <button
        ref={mobileSlideNavigationToggleRef}
        type="button"
        className="work-presentation-slide-navigation-toggle"
        aria-label="打开幻灯片导航"
        aria-controls={mobileSlideNavigationId}
        aria-expanded={mobileSlideNavigationOpen}
        onClick={() => setMobileSlideNavigationOpen(true)}
      >
        <GalleryVerticalEnd size={15} />
        <span>第 {selectedSlideIndex + 1} 张</span>
      </button>
      <PresentationThumbnailRail
        aspectRatio={aspectRatio}
        content={content}
        designContent={designContent}
        mobileCloseButtonRef={mobileSlideNavigationCloseRef}
        mobileNavigationModal={mobileSlideNavigationOpen}
        mobileNavigationId={mobileSlideNavigationId}
        selectedSlide={selectedSlide}
        viewMode={viewMode}
        zoom={zoom}
        onAddSlide={commands.addSlide}
        onCloseMobileNavigation={closeMobileSlideNavigation}
        onDeleteSlide={commands.deleteSlideById}
        onOpenContextMenu={onOpenContextMenu}
        onSelectSlide={selectSlide}
        onViewModeChange={commands.setViewMode}
      />

      {mobileSlideNavigationOpen && (
        <button
          type="button"
          className="work-presentation-slide-navigation-backdrop"
          aria-label="关闭幻灯片导航遮罩"
          tabIndex={-1}
          onClick={closeMobileSlideNavigation}
        />
      )}

      <div
        className="work-slide-stage"
        onPointerMove={onContinueDrag}
        onPointerUp={onDragEnd}
        onPointerCancel={onDragCancel}
      >
        <section
          ref={canvasRef}
          className="work-slide-canvas interactive"
          aria-label={canvasName}
          style={{
            background: activeBackground,
            aspectRatio,
            width: `${zoom}%`,
            maxWidth: `${(1050 * zoom) / 100}px`,
          }}
          onPointerDown={() => commands.selectElement(null, false)}
          onContextMenu={(event) => {
            if (designMode !== 'slide') return;
            onOpenContextMenu(event, selectedSlide, selectedSlideIndex);
          }}
        >
          {inheritedElements.map((element) => (
            <SlideElementPreview
              element={element}
              key={`inherited:${element.id}`}
              origin="inherited"
            />
          ))}
          {placeholderGuides.map((definition) => (
            <button
              type="button"
              className="work-slide-placeholder-guide"
              key={`placeholder:${definition.placeholder?.key ?? definition.id}`}
              style={slideElementStyle(definition)}
              aria-label={`添加${definition.placeholder?.type === 'title' ? '标题' : '内容'}占位符`}
              onPointerDown={(event) => event.stopPropagation()}
              onClick={(event) => {
                event.stopPropagation();
                commands.instantiatePlaceholder(definition);
              }}
            >
              {definition.placeholder?.prompt ?? '单击添加内容'}
            </button>
          ))}
          {snapGuides.map((guide) => (
            <span
              aria-hidden="true"
              className={`work-slide-snap-guide ${guide.axis === 'x' ? 'vertical' : 'horizontal'}`}
              data-presentation-snap-guide={guide.axis}
              data-presentation-snap-source={guide.source}
              key={guide.axis}
              style={
                guide.axis === 'x'
                  ? { left: `${guide.position}%` }
                  : { top: `${guide.position}%` }
              }
            />
          ))}
          {selectionBounds && selectedElements.length > 1 && (
            <span
              className="work-slide-selection-frame"
              data-presentation-selection-frame
              style={{
                left: `${selectionBounds.left}%`,
                top: `${selectionBounds.top}%`,
                width: `${selectionBounds.width}%`,
                height: `${selectionBounds.height}%`,
              }}
            >
              <button
                type="button"
                aria-label={selectionResizeLabel}
                title={selectionResizeLabel}
                className="work-slide-selection-resize-handle"
                data-presentation-selection-control
                onPointerDown={(event) => {
                  if (!selectionTransformAnchor) return;
                  onBeginDrag(event, selectionTransformAnchor, 'resize');
                }}
              />
            </span>
          )}
          {activeElements.map((element) => {
            const selected = selectedElementSet.has(element.id);
            const editing = editingElementId === element.id;
            const label =
              element.altText?.trim() ||
              element.text?.trim() ||
              element.placeholder?.prompt?.trim() ||
              '幻灯片元素';
            return (
              <fieldset
                key={element.id}
                className={`work-slide-element ${element.type} ${element.placeholder ? 'placeholder' : ''} ${
                  selected ? 'selected' : ''
                } ${
                  selected && selectedElements.length > 1
                    ? 'multi-selected'
                    : ''
                } ${editing ? 'editing' : ''}`}
                // biome-ignore lint/a11y/noNoninteractiveTabindex: Slide objects are keyboard-selectable and support object commands.
                tabIndex={0}
                data-slide-element-id={element.id}
                data-slide-element-group-path={
                  element.groupIds?.length
                    ? element.groupIds.join('/')
                    : undefined
                }
                data-slide-element-origin={designMode}
                data-slide-element-selected={selected ? 'true' : 'false'}
                style={slideElementStyle(element)}
                onClick={(event) => {
                  if (
                    !editing &&
                    event.target instanceof HTMLElement &&
                    event.target.closest('a')
                  ) {
                    event.preventDefault();
                  }
                }}
                onDoubleClick={(event) => {
                  if (!presentationElementCanEditContent(element)) return;
                  event.preventDefault();
                  event.stopPropagation();
                  commands.editElement(element.id);
                }}
                onFocus={(event) => {
                  if (event.currentTarget !== event.target || selected) return;
                  commands.selectElement(element.id, false);
                }}
                onKeyDown={(event) => {
                  if (isWorkspaceContextMenuKeyboardEvent(event)) {
                    event.preventDefault();
                    event.stopPropagation();
                    onOpenContextMenu(
                      event,
                      selectedSlide,
                      selectedSlideIndex,
                      element,
                    );
                    return;
                  }
                  if (
                    event.key !== 'Enter' ||
                    editing ||
                    !presentationElementCanEditContent(element)
                  ) {
                    return;
                  }
                  event.preventDefault();
                  event.stopPropagation();
                  commands.editElement(element.id);
                }}
                onContextMenu={(event) => {
                  if (!selected) commands.selectElement(element.id, false);
                  if (designMode !== 'slide') return;
                  onOpenContextMenu(
                    event,
                    selectedSlide,
                    selectedSlideIndex,
                    element,
                  );
                }}
                onPointerDown={(event) => {
                  if (
                    event.target instanceof HTMLTextAreaElement ||
                    (event.target instanceof HTMLElement &&
                      event.target.closest('[data-slide-editor]'))
                  ) {
                    event.stopPropagation();
                    return;
                  }
                  if (event.shiftKey) {
                    event.preventDefault();
                    event.stopPropagation();
                    commands.selectElement(element.id, true);
                    return;
                  }
                  event.stopPropagation();
                  onBeginDrag(event, element, 'move');
                }}
              >
                <legend className="sr-only">{label}</legend>
                {element.type === 'image' && element.image ? (
                  <img
                    src={element.image.dataUrl}
                    alt={element.altText ?? element.image.name}
                    draggable={false}
                  />
                ) : element.type === 'table' && element.table ? (
                  editing ? (
                    <EditableSlideTable
                      element={element}
                      onChange={(rows) =>
                        commands.updateElement({
                          table: { ...element.table, rows },
                        })
                      }
                    />
                  ) : (
                    <SlideTablePreview element={element} />
                  )
                ) : element.type === 'chart' && element.chart ? (
                  <SlideChart
                    chart={element.chart}
                    label={element.altText ?? element.chart.title ?? '图表'}
                  />
                ) : element.textRuns?.length ||
                  element.text ||
                  element.type === 'text' ||
                  element.type === 'shape' ? (
                  editing ? (
                    <PresentationTextEditor
                      autoFocus
                      element={element}
                      onChange={(value) =>
                        commands.updateTextElement(element.id, value)
                      }
                      onEditorChange={(editor) =>
                        onTextEditorChange(element.id, editor)
                      }
                      onExitEditing={commands.exitEditing}
                      onSelectionChange={onTextSelectionChange}
                    />
                  ) : (
                    <SlideElementTextPreview
                      element={element}
                      showPlaceholder
                    />
                  )
                ) : null}
                {selected && !editing && (
                  <>
                    <span
                      className="work-slide-move-handle"
                      aria-hidden="true"
                      onPointerDown={(event) =>
                        onBeginDrag(event, element, 'move')
                      }
                    />
                    {selectedElements.length === 1 && (
                      <span
                        className="work-slide-resize-handle"
                        aria-hidden="true"
                        onPointerDown={(event) =>
                          onBeginDrag(event, element, 'resize')
                        }
                      />
                    )}
                  </>
                )}
              </fieldset>
            );
          })}
          {designMode === 'slide' &&
            (selectedSlide.comments ?? []).map((comment, index) => (
              <button
                type="button"
                className={`work-presentation-comment-pin ${comment.id === activeCommentId ? 'active' : ''}`}
                key={comment.id}
                aria-label={`打开演示批注 ${index + 1}`}
                style={{ left: `${comment.x}%`, top: `${comment.y}%` }}
                onPointerDown={(event) => event.stopPropagation()}
                onClick={(event) => {
                  event.stopPropagation();
                  commands.openComment(comment.id);
                }}
              >
                {index + 1}
              </button>
            ))}
        </section>
        <footer>
          <span>
            {designMode === 'layout'
              ? `布局：${selectedLayout?.name ?? ''}`
              : designMode === 'master'
                ? `母版：${selectedMaster?.name ?? ''}`
                : `幻灯片 ${selectedSlideIndex + 1} / ${content.slides.length}`}
            {selectedElements.length > 0 && (
              <>
                {' · '}
                <span aria-live="polite">
                  {presentationSelectionStatus(
                    selectedElements.length,
                    selectionUnits,
                  )}
                </span>
              </>
            )}
          </span>
          <span>
            {(content.width ?? 13.333).toFixed(2)} ×{' '}
            {(content.height ?? 7.5).toFixed(2)}
          </span>
        </footer>
        {designMode === 'slide' && (
          <div className="work-slide-notes">
            <span>演讲者备注</span>
            <OfficeTextArea
              aria-label="演讲者备注"
              value={selectedSlide.notes ?? ''}
              placeholder="添加演讲者备注"
              onChange={(event) => commands.updateNotes(event.target.value)}
            />
          </div>
        )}
      </div>
    </div>
  );
}

function presentationSelectionStatus(
  selectedElementCount: number,
  units: ReturnType<typeof presentationSelectionUnits>,
): string {
  if (units.length === 1 && units[0].groupId && selectedElementCount > 1) {
    return `已选择 1 组，共 ${selectedElementCount} 个对象`;
  }
  if (units.length === selectedElementCount) {
    return `已选择 ${selectedElementCount} 个对象`;
  }
  return `已选择 ${units.length} 项，共 ${selectedElementCount} 个对象`;
}
