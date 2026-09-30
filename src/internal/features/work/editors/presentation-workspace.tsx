import type { Editor } from '@tiptap/core';
import { GalleryVerticalEnd } from 'lucide-react';
import {
  type PointerEvent,
  type RefObject,
  useEffect,
  useId,
  useRef,
  useState,
} from 'react';
import { useDialogFocusScope } from '../../../design-system/primitives/overlay/dialog-focus-scope';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeKernelPresentationSnapGuide } from '../../../kernel/office-kernel-protocol';
import {
  isWorkspaceContextMenuKeyboardEvent,
  type WorkspaceContextMenuEvent,
} from '../../workspace/components/workspace-context-menu';
import { presentationSelectionUnits } from '../work-presentation-groups';
import type { WorkPresentationDesignContent } from '../work-presentation-layouts';
import type {
  WorkPresentationContent,
  WorkPresentationLayout,
  WorkPresentationMaster,
  WorkSlide,
  WorkSlideElement,
} from '../work-types';
import { OfficeTextArea } from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import { SlideChart } from './presentation-chart-canvas';
import { PresentationCollaborationPresenceLayer } from './presentation-collaboration-presence';
import type { PresentationEditorCommands } from './presentation-command-types';
import type { PresentationDesignMode } from './presentation-editor-types';
import { PresentationObjectList } from './presentation-object-list-panel';
import { PresentationRotateHandle } from './presentation-rotate-handle';
import {
  presentationElementCanEditContent,
  presentationElementDisplayBox,
  presentationSelectionBounds,
  selectedPresentationElements,
} from './presentation-selection';
import {
  EditableSlideTable,
  SlideElementPreview,
  SlideElementTextPreview,
  SlideTablePreview,
  slideElementStyle,
} from './presentation-slide-canvas';
import { PresentationTextEditor } from './presentation-text-editor';
import { PresentationThumbnailRail } from './presentation-thumbnail-rail';

const presentationMobileNavigationQuery = '(max-width: 640px)';

export type PresentationWorkspaceCommands = Pick<
  PresentationEditorCommands,
  | 'addSlide'
  | 'deleteSlideById'
  | 'editElement'
  | 'moveSlide'
  | 'exitEditing'
  | 'instantiatePlaceholder'
  | 'openComment'
  | 'rotateSelection'
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
  designContent: WorkPresentationDesignContent;
  designMode: PresentationDesignMode;
  inheritedElements: WorkSlideElement[];
  placeholderGuides: WorkSlideElement[];
  editingElementId: string | null;
  selectedElementIds: readonly string[];
  selectedLayout: WorkPresentationLayout | undefined;
  selectedMaster: WorkPresentationMaster | undefined;
  selectedSlide: WorkSlide;
  snapGuides: OfficeKernelPresentationSnapGuide[];
  notesVisible: boolean;
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
  notesVisible,
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
  const messages = useOfficeMessages();
  const [mobileSlideNavigationOpen, setMobileSlideNavigationOpen] =
    useState(false);
  const mobileSlideNavigationModal = useMobileSlideNavigationModal();
  const mobileSlideNavigationModalOpen =
    mobileSlideNavigationOpen && mobileSlideNavigationModal;
  const mobileSlideNavigationId = useId();
  const mobileSlideNavigationToggleRef = useRef<HTMLButtonElement>(null);
  const mobileSlideNavigationCloseRef = useRef<HTMLButtonElement>(null);

  const closeMobileSlideNavigation = () => {
    setMobileSlideNavigationOpen(false);
  };
  useDialogFocusScope<HTMLElement>({
    active: mobileSlideNavigationModalOpen,
    onEscape: closeMobileSlideNavigation,
    initialFocus: () => mobileSlideNavigationCloseRef.current,
    getActiveScope: () =>
      document.getElementById(mobileSlideNavigationId) as HTMLElement | null,
    getIsolationExceptions: () => [
      document.querySelector<HTMLElement>(
        '.work-presentation-slide-navigation-backdrop',
      ),
    ],
    restoreFocusTarget: () =>
      mobileSlideNavigationModal
        ? mobileSlideNavigationToggleRef.current
        : (document
            .getElementById(mobileSlideNavigationId)
            ?.querySelector<HTMLElement>('[data-slide-thumbnail].active') ??
          null),
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
  const selectionBounds = presentationSelectionBounds(
    selectedElements.map((element) => {
      const box = presentationElementDisplayBox(element, activeElements);
      return { ...element, x: box.x, y: box.y };
    }),
  );
  const selectionTransformAnchor = selectedElements.at(-1);
  const selectionResizeLabel =
    selectionUnits.length === 1 && selectionUnits[0]?.groupId
      ? officeMessage(messages, 'presentation.workspace.zoomGroup')
      : officeMessage(messages, 'presentation.workspace.zoomObject');
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
        onMoveSlide={commands.moveSlide}
        onOpenContextMenu={onOpenContextMenu}
        onSelectSlide={commands.selectSlide}
        onViewModeChange={commands.setViewMode}
      />
    );
  }

  const selectedSlideIndex = content.slides.findIndex(
    (slide) => slide.id === selectedSlide.id,
  );
  const selectSlide = (
    slideId: string,
    returnToSlideMode: boolean,
    dismissMobileNavigation = true,
  ) => {
    commands.selectSlide(slideId, returnToSlideMode);
    if (mobileSlideNavigationOpen && dismissMobileNavigation)
      closeMobileSlideNavigation();
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
        aria-label={officeMessage(messages, 'presentation.workspace.openNav')}
        aria-controls={mobileSlideNavigationId}
        aria-expanded={mobileSlideNavigationOpen}
        onClick={() => setMobileSlideNavigationOpen(true)}
      >
        <GalleryVerticalEnd size={15} />
        <span>
          {officeMessage(messages, 'presentation.workspace.slideNumber', {
            index: String(selectedSlideIndex + 1),
          })}
        </span>
      </button>
      <PresentationThumbnailRail
        aspectRatio={aspectRatio}
        content={content}
        designContent={designContent}
        mobileCloseButtonRef={mobileSlideNavigationCloseRef}
        mobileNavigationModal={mobileSlideNavigationModalOpen}
        mobileNavigationId={mobileSlideNavigationId}
        selectedSlide={selectedSlide}
        viewMode={viewMode}
        zoom={zoom}
        onAddSlide={commands.addSlide}
        onCloseMobileNavigation={closeMobileSlideNavigation}
        onDeleteSlide={commands.deleteSlideById}
        onMoveSlide={commands.moveSlide}
        onOpenContextMenu={onOpenContextMenu}
        onSelectSlide={selectSlide}
        onViewModeChange={commands.setViewMode}
      />

      {mobileSlideNavigationModalOpen && (
        <button
          type="button"
          className="work-presentation-slide-navigation-backdrop"
          aria-label={officeMessage(
            messages,
            'presentation.workspace.closeNav',
          )}
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
        <div className="work-slide-stage-main">
          <PresentationObjectList
            elements={activeElements}
            selectedElementIds={selectedElementIds}
            listLabel="Slide objects"
            emptyLabel="No objects on this slide"
            onSelectElement={(elementId, additive) =>
              commands.selectElement(elementId, additive)
            }
          />
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
                aria-label={officeMessage(
                  messages,
                  definition.placeholder?.type === 'title'
                    ? 'presentation.workspace.addTitlePlaceholder'
                    : 'presentation.workspace.addContentPlaceholder',
                )}
                onPointerDown={(event) => event.stopPropagation()}
                onClick={(event) => {
                  event.stopPropagation();
                  commands.instantiatePlaceholder(definition);
                }}
              >
                {definition.placeholder?.prompt ??
                  officeMessage(messages, 'presentation.workspace.clickToAdd')}
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
                <PresentationRotateHandle
                  label={officeMessage(
                    messages,
                    'presentation.workspace.rotateHandle',
                  )}
                  onRotate={commands.rotateSelection}
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
                officeMessage(messages, 'presentation.workspace.slideElement');
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
                  data-slide-element-group-rotation={
                    element.groupRotation
                      ? String(element.groupRotation)
                      : undefined
                  }
                  data-slide-element-origin={designMode}
                  data-slide-element-selected={selected ? 'true' : 'false'}
                  style={slideElementStyle({
                    ...element,
                    ...presentationElementDisplayBox(element, activeElements),
                  })}
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
                    if (event.currentTarget !== event.target || selected)
                      return;
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
                      label={
                        element.altText ??
                        element.chart.title ??
                        officeMessage(
                          messages,
                          'presentation.workspace.chartFallback',
                        )
                      }
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
            {designMode === 'slide' && (
              <PresentationCollaborationPresenceLayer
                elements={activeElements}
                slideId={selectedSlide.id}
              />
            )}
            {designMode === 'slide' &&
              (selectedSlide.comments ?? []).map((comment, index) => (
                <button
                  type="button"
                  className={`work-presentation-comment-pin ${comment.id === activeCommentId ? 'active' : ''}`}
                  key={comment.id}
                  aria-label={officeMessage(
                    messages,
                    'presentation.workspace.openCommentAria',
                    { index: String(index + 1) },
                  )}
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
        </div>
        <footer>
          <span>
            {designMode === 'layout'
              ? officeMessage(messages, 'presentation.workspace.layoutMeta', {
                  name: selectedLayout?.name ?? '',
                })
              : designMode === 'master'
                ? officeMessage(messages, 'presentation.workspace.masterMeta', {
                    name: selectedMaster?.name ?? '',
                  })
                : officeMessage(messages, 'presentation.workspace.slideMeta', {
                    current: String(selectedSlideIndex + 1),
                    total: String(content.slides.length),
                  })}
            {selectedElements.length > 0 && (
              <>
                {' · '}
                <span aria-live="polite">
                  {presentationSelectionStatus(
                    selectedElements.length,
                    selectionUnits,
                    messages,
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
        {designMode === 'slide' && notesVisible && (
          <div className="work-slide-notes">
            <span>
              {officeMessage(messages, 'presentation.workspace.notesLabel')}
            </span>
            <OfficeTextArea
              aria-label={officeMessage(
                messages,
                'presentation.workspace.notesAria',
              )}
              value={selectedSlide.notes ?? ''}
              placeholder={officeMessage(
                messages,
                'presentation.workspace.notesPlaceholder',
              )}
              onChange={(event) => commands.updateNotes(event.target.value)}
            />
          </div>
        )}
      </div>
    </div>
  );
}

function useMobileSlideNavigationModal(): boolean {
  const [matches, setMatches] = useState(() =>
    mediaQueryMatches(presentationMobileNavigationQuery),
  );

  useEffect(() => {
    if (
      typeof window === 'undefined' ||
      typeof window.matchMedia !== 'function'
    ) {
      return;
    }
    const mediaQuery = window.matchMedia(presentationMobileNavigationQuery);
    const update = () => setMatches(mediaQuery.matches);
    update();
    if (typeof mediaQuery.addEventListener === 'function') {
      mediaQuery.addEventListener('change', update);
      return () => mediaQuery.removeEventListener('change', update);
    }
    mediaQuery.addListener(update);
    return () => mediaQuery.removeListener(update);
  }, []);

  return matches;
}

function mediaQueryMatches(query: string): boolean {
  if (
    typeof window === 'undefined' ||
    typeof window.matchMedia !== 'function'
  ) {
    return true;
  }
  return window.matchMedia(query).matches;
}

function presentationSelectionStatus(
  selectedElementCount: number,
  units: ReturnType<typeof presentationSelectionUnits>,
  messages: ReturnType<typeof useOfficeMessages>,
): string {
  if (units.length === 1 && units[0].groupId && selectedElementCount > 1) {
    return officeMessage(messages, 'presentation.workspace.selectedOneGroup', {
      count: String(selectedElementCount),
    });
  }
  if (units.length === selectedElementCount) {
    return officeMessage(messages, 'presentation.workspace.selectedObjects', {
      count: String(selectedElementCount),
    });
  }
  return officeMessage(messages, 'presentation.workspace.selectedMixed', {
    units: String(units.length),
    count: String(selectedElementCount),
  });
}
