import type { Editor } from '@tiptap/core';
import { Plus } from 'lucide-react';
import type { CSSProperties, MouseEvent, PointerEvent, RefObject } from 'react';
import type { OfficeKernelPresentationSnapGuide } from '../../../kernel/office-kernel-protocol';
import type {
  WorkPresentationContent,
  WorkPresentationLayout,
  WorkPresentationMaster,
  WorkSlide,
  WorkSlideElement,
} from '../work-types';
import { OfficeTextArea } from './office-controls';
import { SlideChart } from './presentation-chart-canvas';
import type { PresentationDesignMode } from './presentation-design-panel';
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
import { PresentationSlideThumbnail } from './presentation-slide-thumbnail';
import {
  PresentationTextEditor,
  type PresentationTextValue,
} from './presentation-text-editor';

export interface PresentationWorkspaceProps {
  activeBackground: string;
  activeCommentId: string | null;
  activeElements: WorkSlideElement[];
  aspectRatio: string;
  canvasName: string;
  canvasRef: RefObject<HTMLElement | null>;
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
  onAddSlide: () => void;
  onBeginDrag: (
    event: PointerEvent,
    element: WorkSlideElement,
    mode: 'move' | 'resize',
  ) => void;
  onContinueDrag: (event: PointerEvent) => void;
  onDeleteSlide: (slideId: string) => boolean;
  onDragCancel: () => void;
  onDragEnd: (event: PointerEvent) => void;
  onEditElement: (elementId: string) => void;
  onExitEditing: () => void;
  onInstantiatePlaceholder: (definition: WorkSlideElement) => void;
  onOpenAgentMenu: (
    event: MouseEvent,
    slide: WorkSlide,
    slideIndex: number,
    element?: WorkSlideElement | null,
  ) => void;
  onOpenComment: (commentId: string) => void;
  onSelectElement: (elementId: string | null, additive: boolean) => void;
  onSelectSlide: (slideId: string, returnToSlideMode: boolean) => void;
  onTextEditorChange: (elementId: string, editor: Editor | null) => void;
  onTextSelectionChange: () => void;
  onUpdateElement: (patch: Partial<WorkSlideElement>) => void;
  onUpdateNotes: (notes: string) => void;
  onUpdateTextElement: (
    elementId: string,
    value: PresentationTextValue,
  ) => void;
  onViewModeChange: (mode: 'normal' | 'sorter') => void;
}

export function PresentationWorkspace({
  activeBackground,
  activeCommentId,
  activeElements,
  aspectRatio,
  canvasName,
  canvasRef,
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
  onAddSlide,
  onBeginDrag,
  onContinueDrag,
  onDeleteSlide,
  onDragCancel,
  onDragEnd,
  onEditElement,
  onExitEditing,
  onInstantiatePlaceholder,
  onOpenAgentMenu,
  onOpenComment,
  onSelectElement,
  onSelectSlide,
  onTextEditorChange,
  onTextSelectionChange,
  onUpdateElement,
  onUpdateNotes,
  onUpdateTextElement,
  onViewModeChange,
}: PresentationWorkspaceProps) {
  const selectedElementSet = new Set(selectedElementIds);
  const selectedElements = selectedPresentationElements(
    activeElements,
    selectedElementIds,
  );
  const selectionBounds = presentationSelectionBounds(selectedElements);
  if (viewMode === 'sorter') {
    return (
      <section
        className="work-presentation-sorter"
        aria-label="幻灯片浏览视图"
        style={
          {
            '--work-presentation-sorter-width': `${Math.round(220 * (zoom / 100))}px`,
          } as CSSProperties
        }
      >
        {content.slides.map((slide, index) => (
          <PresentationSlideThumbnail
            key={slide.id}
            content={designContent}
            slide={slide}
            index={index}
            selected={slide.id === selectedSlide.id}
            aspectRatio={aspectRatio}
            variant="sorter"
            onSelect={() => onSelectSlide(slide.id, false)}
            onDelete={() => onDeleteSlide(slide.id)}
            onDoubleClick={() => onViewModeChange('normal')}
          />
        ))}
      </section>
    );
  }

  const selectedSlideIndex = content.slides.findIndex(
    (slide) => slide.id === selectedSlide.id,
  );
  return (
    <div className="work-presentation-layout">
      <aside className="work-slide-strip" aria-label="幻灯片">
        {content.slides.map((slide, index) => (
          <PresentationSlideThumbnail
            key={slide.id}
            content={designContent}
            slide={slide}
            index={index}
            selected={slide.id === selectedSlide.id}
            aspectRatio={aspectRatio}
            variant="strip"
            onSelect={() => onSelectSlide(slide.id, true)}
            onDelete={() => onDeleteSlide(slide.id)}
            onContextMenu={(event) => {
              onSelectSlide(slide.id, false);
              onOpenAgentMenu(event, slide, index);
            }}
          />
        ))}
        <button type="button" className="work-slide-add" onClick={onAddSlide}>
          <Plus size={15} />
          添加幻灯片
        </button>
      </aside>

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
          onPointerDown={() => onSelectElement(null, false)}
          onContextMenu={(event) => {
            if (designMode !== 'slide') return;
            onOpenAgentMenu(event, selectedSlide, selectedSlideIndex);
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
                onInstantiatePlaceholder(definition);
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
              aria-hidden="true"
              className="work-slide-selection-frame"
              data-presentation-selection-frame
              style={{
                left: `${selectionBounds.left}%`,
                top: `${selectionBounds.top}%`,
                width: `${selectionBounds.width}%`,
                height: `${selectionBounds.height}%`,
              }}
            />
          )}
          {activeElements.map((element) => {
            const selected = selectedElementSet.has(element.id);
            const editing = editingElementId === element.id;
            const label =
              element.altText?.trim() || element.text?.trim() || '幻灯片元素';
            return (
              <fieldset
                key={element.id}
                className={`work-slide-element ${element.type} ${element.placeholder ? 'placeholder' : ''} ${
                  selected ? 'selected' : ''
                } ${editing ? 'editing' : ''}`}
                // biome-ignore lint/a11y/noNoninteractiveTabindex: Slide objects are keyboard-selectable and support object commands.
                tabIndex={0}
                data-slide-element-id={element.id}
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
                  onEditElement(element.id);
                }}
                onFocus={(event) => {
                  if (event.currentTarget !== event.target || selected) return;
                  onSelectElement(element.id, false);
                }}
                onKeyDown={(event) => {
                  if (
                    event.key !== 'Enter' ||
                    editing ||
                    !presentationElementCanEditContent(element)
                  ) {
                    return;
                  }
                  event.preventDefault();
                  event.stopPropagation();
                  onEditElement(element.id);
                }}
                onContextMenu={(event) => {
                  if (!selected) onSelectElement(element.id, false);
                  if (designMode !== 'slide') return;
                  onOpenAgentMenu(
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
                    onSelectElement(element.id, true);
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
                        onUpdateElement({
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
                        onUpdateTextElement(element.id, value)
                      }
                      onEditorChange={(editor) =>
                        onTextEditorChange(element.id, editor)
                      }
                      onExitEditing={onExitEditing}
                      onSelectionChange={onTextSelectionChange}
                    />
                  ) : (
                    <SlideElementTextPreview element={element} />
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
                  onOpenComment(comment.id);
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
                  已选择 {selectedElements.length} 个对象
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
              onChange={(event) => onUpdateNotes(event.target.value)}
            />
          </div>
        )}
      </div>
    </div>
  );
}
