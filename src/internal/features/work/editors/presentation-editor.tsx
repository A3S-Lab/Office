import type { Editor } from '@tiptap/core';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  type WorkspaceContextMenuEvent,
  WorkspaceContextMenu,
  workspaceContextMenuPosition,
} from '../../workspace/components/workspace-context-menu';
import { presentationAgentMenuItems } from '../components/work-editor-agent-menus';
import { applyPresentationAgentProposalChanges } from '../work-agent-proposal-apply';
import {
  presentationAgentProposalTargets,
  presentationAgentSelection,
  presentationNotesProposalTarget,
} from '../work-presentation-agent-context';
import {
  presentationSlideView,
  withPresentationDesign,
} from '../work-presentation-layouts';
import {
  canGroupPresentationElements,
  canUngroupPresentationElements,
  presentationSelectionUnits,
} from '../work-presentation-groups';
import type { WorkSlide, WorkSlideElement } from '../work-types';
import { OfficeFileInput } from './office-controls';
import { createPresentationArrangementController } from './presentation-arrangement-controller';
import { PresentationChartPanel } from './presentation-chart-panel';
import { createPresentationEditorExtensions } from './presentation-command-controller';
import type { PresentationEditorCommands } from './presentation-command-types';
import {
  PresentationCommentsPanel,
  presentationCommentCount,
} from './presentation-comments-panel';
import { presentationCoreContextMenuItems } from './presentation-context-menu';
import { PresentationDesignPanel } from './presentation-design-panel';
import { updatePresentationElements } from './presentation-editor-operations';
import {
  type PresentationObjectFocusState,
  type PresentationWorkspaceFocusState,
  presentationCommandsWithObjectFocus,
  restorePresentationWorkspaceFocus,
} from './presentation-editor-focus';
import {
  presentationElementSupportsTextFormatting,
  selectedPresentationElements,
} from './presentation-selection';
import type {
  PresentationAgentMenuState,
  PresentationDesignMode,
  PresentationEditorProps,
} from './presentation-editor-types';
import { PresentationPlayer } from './presentation-player';
import { PresentationStatusBar } from './presentation-status-bar';
import { useOfficeTaskPaneEscape } from './office-task-pane';
import {
  applyPresentationTextFormatting,
  presentationTextToolbarState,
} from './presentation-text-editor';
import { presentationElementToolbarState } from './presentation-text-formatting';
import { PresentationToolbar } from './presentation-toolbar';
import { PresentationWorkspace } from './presentation-workspace';
import { usePresentationClipboard } from './use-presentation-clipboard';
import { usePresentationDesignCommands } from './use-presentation-design-commands';
import { usePresentationElementCommands } from './use-presentation-element-commands';
import { usePresentationGeometry } from './use-presentation-geometry';
import { usePresentationHistory } from './use-presentation-history';
import { useOfficeEditorKeyboardShortcuts } from './use-office-editor-keyboard-shortcuts';
import {
  stepOfficeZoom,
  useOfficeEditorWheelZoom,
} from './use-office-editor-wheel-zoom';
import { useOfficeEditorRuntime } from './use-office-editor-runtime';
import { usePresentationReviewCommands } from './use-presentation-review-commands';
import { usePresentationSelection } from './use-presentation-selection';
import { usePresentationSlideCommands } from './use-presentation-slide-commands';
import {
  type PresentationTransformCommit,
  usePresentationTransform,
} from './use-presentation-transform';
import { WorkOfficePreviewBar } from './work-office-chrome';

export type { PresentationEditorProps } from './presentation-editor-types';

type PresentationTaskPane = 'comments' | 'design' | null;
type PresentationTextFormattingAttribute = 'bold' | 'italic' | 'underline';

export function PresentationEditor(props: PresentationEditorProps) {
  const { content, fileActions, preview } = props;
  if (preview) {
    return (
      <section className="work-presentation-editor preview">
        <WorkOfficePreviewBar
          ariaLabel="演示预览工具"
          label="只读预览"
          detail={`${content.slides.length} 张幻灯片`}
          fileActions={fileActions}
          className="work-presentation-ribbon"
        />
        <PresentationPlayer content={content} />
      </section>
    );
  }
  const initialSlide = content.slides[0];
  if (!initialSlide) return null;
  return <PresentationEditingSurface {...props} initialSlide={initialSlide} />;
}

function PresentationEditingSurface({
  initialSlide,
  content,
  preview,
  saveStatus = '已自动保存',
  fileActions,
  kernelWasmUrl,
  onChange,
  onAgentRequest,
  onStartSlideshow,
}: PresentationEditorProps & { initialSlide: WorkSlide }) {
  const contentRef = useRef(content);
  const presentationCommandsRef = useRef<PresentationEditorCommands | null>(
    null,
  );
  const [selectedSlideId, setSelectedSlideId] = useState(
    content.slides[0]?.id ?? '',
  );
  const [activeTextEditor, setActiveTextEditor] = useState<{
    elementId: string;
    editor: Editor;
  } | null>(null);
  const [, setTextSelectionVersion] = useState(0);
  const [taskPane, setTaskPane] = useState<PresentationTaskPane>(null);
  const commentsOpen = taskPane === 'comments';
  const designOpen = taskPane === 'design';
  const [designMode, setDesignMode] = useState<PresentationDesignMode>('slide');
  const [activeCommentId, setActiveCommentId] = useState<string | null>(null);
  const [agentMenu, setAgentMenu] = useState<PresentationAgentMenuState | null>(
    null,
  );
  const [viewMode, setViewMode] = useState<'normal' | 'sorter'>('normal');
  const [slideshowStartIndex, setSlideshowStartIndex] = useState<number | null>(
    null,
  );
  const [zoom, setZoom] = useState(90);
  const slideshowReturnFocusRef = useRef<HTMLElement | null>(null);
  const presentationRootRef = useRef<HTMLElement>(null);
  const canvasRef = useRef<HTMLElement>(null);
  const objectFocusStateRef = useRef<PresentationObjectFocusState>({
    editingElementId: null,
    selectedElementIds: [],
  });
  const workspaceFocusStateRef = useRef<PresentationWorkspaceFocusState>({
    editingElementId: null,
    selectedElementIds: [],
    selectedSlideId: selectedSlideId,
    viewMode: 'normal',
  });
  const imageInputRef = useRef<HTMLInputElement>(null);
  const geometry = usePresentationGeometry(kernelWasmUrl, !preview);
  const designContent = withPresentationDesign(content);
  const selectedSlide =
    designContent.slides.find((slide) => slide.id === selectedSlideId) ??
    designContent.slides[0] ??
    initialSlide;
  const selectedLayout =
    designContent.layouts?.find(
      (layout) => layout.id === selectedSlide?.layoutId,
    ) ?? designContent.layouts?.[0];
  const selectedMaster =
    designContent.masters?.find(
      (master) => master.id === selectedLayout?.masterId,
    ) ?? designContent.masters?.[0];
  const activeElements =
    designMode === 'layout'
      ? (selectedLayout?.elements ?? [])
      : designMode === 'master'
        ? (selectedMaster?.elements ?? [])
        : (selectedSlide?.elements ?? []);
  const selection = usePresentationSelection(activeElements);
  const selectedElements = selectedPresentationElements(
    activeElements,
    selection.selectedElementIds,
  );
  const selectionUnits = presentationSelectionUnits(
    activeElements,
    selection.selectedElementIds,
  );
  const selectedElementId = selection.selectedElementIds.at(-1) ?? null;
  const selectedElement =
    activeElements.find((element) => element.id === selectedElementId) ?? null;
  const singleSelectedElement =
    selectedElements.length === 1 ? selectedElement : null;
  const selectedTextEditor =
    activeTextEditor?.elementId === selection.editingElementId &&
    !activeTextEditor.editor.isDestroyed
      ? activeTextEditor.editor
      : null;
  objectFocusStateRef.current = {
    editingElementId: selection.editingElementId,
    selectedElementIds: selection.selectedElementIds,
  };
  workspaceFocusStateRef.current = {
    ...objectFocusStateRef.current,
    selectedSlideId: selectedSlide.id,
    viewMode,
  };
  const toolbarSelectedElement = selectedElement
    ? selectedTextEditor && !selectedTextEditor.state.selection.empty
      ? presentationTextToolbarState(selectedTextEditor, selectedElement)
      : presentationElementToolbarState(selectedElement)
    : null;
  const slideView = selectedSlide
    ? presentationSlideView(designContent, selectedSlide)
    : undefined;
  const activeBackground =
    designMode === 'layout'
      ? (selectedLayout?.background ?? selectedMaster?.background ?? '#ffffff')
      : designMode === 'master'
        ? (selectedMaster?.background ?? '#ffffff')
        : (slideView?.background ?? selectedSlide?.background ?? '#ffffff');
  const inheritedElements =
    designMode === 'slide'
      ? (slideView?.inheritedElements ?? [])
      : designMode === 'layout'
        ? (selectedMaster?.elements.filter((element) => !element.placeholder) ??
          [])
        : [];
  const placeholderGuides =
    designMode === 'slide'
      ? (slideView?.placeholderElements.filter(
          (definition) =>
            !activeElements.some(
              (element) =>
                element.placeholder?.key === definition.placeholder?.key,
            ),
        ) ?? [])
      : [];
  const canvasName =
    designMode === 'layout'
      ? `${selectedLayout?.name ?? '布局'}布局编辑画布`
      : designMode === 'master'
        ? `${selectedMaster?.name ?? '母版'}母版编辑画布`
        : `${selectedSlide?.name ?? '幻灯片'}编辑画布`;
  const activeTargetId =
    designMode === 'layout'
      ? selectedLayout?.id
      : designMode === 'master'
        ? selectedMaster?.id
        : selectedSlide?.id;
  const commitElementTransform = useCallback(
    (changes: readonly PresentationTransformCommit[]) => {
      if (!activeTargetId || !changes.length) return;
      const patches = new Map(
        changes.map((change) => [change.elementId, change.patch]),
      );
      updatePresentationElements(
        contentRef.current,
        designMode,
        activeTargetId,
        (elements) =>
          elements.map((element) => {
            const patch = patches.get(element.id);
            return patch ? { ...element, ...patch } : element;
          }),
        (next) => {
          presentationCommandsRef.current?.setPresentationContent(next);
        },
      );
    },
    [activeTargetId, designMode],
  );
  const transform = usePresentationTransform({
    canvasRef,
    elements: activeElements,
    geometry,
    onCommit: commitElementTransform,
    onSelect: (elementId) => selection.select(elementId),
    selectedElementIds: selection.selectedElementIds,
    snapTargets: [
      ...inheritedElements,
      ...placeholderGuides,
      ...activeElements,
    ],
  });
  const agentMenuSlide = agentMenu
    ? (content.slides.find((slide) => slide.id === agentMenu.slideId) ?? null)
    : null;
  const agentMenuElement =
    agentMenuSlide && agentMenu?.elementId
      ? (agentMenuSlide.elements.find(
          (element) => element.id === agentMenu.elementId,
        ) ?? null)
      : null;
  const aspectRatio = `${content.width ?? 13.333} / ${content.height ?? 7.5}`;
  contentRef.current = content;
  const openContextMenu = (
    event: WorkspaceContextMenuEvent,
    slide: WorkSlide,
    slideIndex: number,
    element?: WorkSlideElement | null,
  ) => {
    event.preventDefault();
    event.stopPropagation();
    const position = workspaceContextMenuPosition(event);
    setAgentMenu({
      x: position.x,
      y: position.y,
      selection: presentationAgentSelection(
        slide,
        slideIndex,
        content.slides.length,
        element,
      ),
      target: element ? 'element' : 'slide',
      slideId: slide.id,
      elementId: element?.id ?? null,
    });
  };

  useEffect(() => {
    if (!content.slides.some((slide) => slide.id === selectedSlideId)) {
      setSelectedSlideId(content.slides[0]?.id ?? '');
      selection.clear();
    }
  }, [content.slides, selectedSlideId, selection.clear]);

  const history = usePresentationHistory({
    content,
    onChange,
    selectedSlideId,
    onSelectSlide: (slideId) => {
      setSelectedSlideId(slideId);
      selection.clear();
    },
  });
  const presentationSlides = usePresentationSlideCommands({
    content,
    onChange,
    onClearSelection: selection.clear,
    onSelectSlide: setSelectedSlideId,
    selectedSlide,
  });

  const presentationElements = usePresentationElementCommands({
    activeElements,
    activeTargetId,
    content,
    contentRef,
    designMode,
    onChange,
    onEditElement: selection.edit,
    onSelectElements: selection.replace,
    selectedElementIds: selection.selectedElementIds,
    selectedElements,
    selectedSlide,
  });

  const clipboard = usePresentationClipboard({
    content,
    mode: designMode,
    targetId: activeTargetId,
    selectedSlide,
    selectedElements,
    onChange,
    onSelectSlide: setSelectedSlideId,
    onSelectElements: selection.replace,
  });
  const closeComments = useCallback(() => {
    setTaskPane((current) => (current === 'comments' ? null : current));
  }, []);
  const openComments = useCallback(() => {
    if (designMode !== 'slide') {
      setDesignMode('slide');
      selection.clear();
    }
    setTaskPane('comments');
  }, [designMode, selection.clear]);
  const toggleComments = useCallback(() => {
    if (commentsOpen) {
      closeComments();
      return;
    }
    openComments();
  }, [closeComments, commentsOpen, openComments]);
  const setDesignPanelOpen = useCallback((open: boolean) => {
    setTaskPane((current) =>
      open ? 'design' : current === 'design' ? null : current,
    );
  }, []);
  const presentationReview = usePresentationReviewCommands({
    content,
    onChange,
    onClearSelection: selection.clear,
    onCloseComments: closeComments,
    onOpenComments: openComments,
    onSelectComment: setActiveCommentId,
    onSelectSlide: setSelectedSlideId,
    selectedElement,
    selectedSlide,
  });

  const presentationDesign = usePresentationDesignCommands({
    activeElements,
    activeTargetId,
    designContent,
    designMode,
    designOpen,
    onChange,
    onClearSelection: selection.clear,
    onCloseComments: closeComments,
    onEditElement: selection.edit,
    onSetDesignMode: setDesignMode,
    onSetDesignOpen: setDesignPanelOpen,
    selectedLayout,
    selectedMaster,
    selectedSlide,
  });
  const chartPaneOpen =
    !designOpen &&
    !commentsOpen &&
    designMode === 'slide' &&
    singleSelectedElement?.type === 'chart' &&
    Boolean(singleSelectedElement.chart);
  const closeVisibleTaskPane = useCallback(() => {
    if (designOpen) {
      presentationDesign.close();
      return;
    }
    if (commentsOpen) {
      closeComments();
      return;
    }
    selection.clear();
  }, [
    closeComments,
    commentsOpen,
    designOpen,
    presentationDesign,
    selection.clear,
  ]);
  useOfficeTaskPaneEscape(
    Boolean(designOpen || commentsOpen || chartPaneOpen),
    closeVisibleTaskPane,
  );
  const arrangement = createPresentationArrangementController({
    getContent: () => contentRef.current,
    geometry,
    mode: designMode,
    onChange,
    selectedElementIds: selection.selectedElementIds,
    selectedElements,
    targetId: activeTargetId,
  });
  const canGroupSelection = canGroupPresentationElements(
    activeElements,
    selection.selectedElementIds,
  );
  const canUngroupSelection = canUngroupPresentationElements(
    activeElements,
    selection.selectedElementIds,
  );
  const textFormattingAvailable =
    selectedElements.length > 0 &&
    selectedElements.every(presentationElementSupportsTextFormatting);
  const toggleTextFormatting = useCallback(
    (attribute: PresentationTextFormattingAttribute): boolean => {
      if (
        selectedTextEditor &&
        !selectedTextEditor.state.selection.empty &&
        toolbarSelectedElement
      ) {
        const value = !toolbarSelectedElement[attribute];
        return applyPresentationTextFormatting(
          selectedTextEditor,
          attribute === 'bold'
            ? { bold: value }
            : attribute === 'italic'
              ? { italic: value }
              : { underline: value },
        );
      }
      if (attribute === 'bold') return clipboard.toggleBold();
      if (attribute === 'italic') return clipboard.toggleItalic();
      return clipboard.toggleUnderline();
    },
    [
      clipboard.toggleBold,
      clipboard.toggleItalic,
      clipboard.toggleUnderline,
      selectedTextEditor,
      toolbarSelectedElement,
    ],
  );
  const presentationExtensions = useMemo(
    createPresentationEditorExtensions,
    [],
  );
  const presentationEditor = useOfficeEditorRuntime(
    {
      clipboard: {
        canCopySelection:
          selectedElements.length > 0 ||
          (designMode === 'slide' && Boolean(selectedSlide)),
        canCutSelection:
          selectedElements.length > 0 ||
          (designMode === 'slide' && Boolean(selectedSlide)),
        canPasteSelection: true,
        copySelection: clipboard.copySelection,
        cutSelection: clipboard.cutSelection,
        pasteSelection: clipboard.pasteSelection,
      },
      design: {
        addPlaceholder: presentationDesign.addPlaceholder,
        applyLayout: presentationDesign.applyLayout,
        canDeleteLayout: (designContent.layouts?.length ?? 0) >= 2,
        close: presentationDesign.close,
        createLayout: presentationDesign.createLayout,
        deleteLayout: presentationDesign.deleteLayout,
        edit: presentationDesign.edit,
        renameLayout: presentationDesign.renameLayout,
        renameMaster: presentationDesign.renameMaster,
        setLayoutBackground: presentationDesign.setLayoutBackground,
        setMasterBackground: presentationDesign.setMasterBackground,
        toggleLayoutBackground: presentationDesign.toggleLayoutBackground,
      },
      document: {
        setContent: (next) => {
          contentRef.current = next;
          onChange(next);
        },
      },
      elements: {
        canAlignElement: selectionUnits.length >= 2,
        canDistributeElements: selectionUnits.length >= 3,
        canGroupElements: canGroupSelection,
        canReorderElement: selectedElements.length > 0,
        canUngroupElements: canUngroupSelection,
        canUpdateElement: selectedElements.length > 0,
        alignElement: arrangement.align,
        distributeElements: arrangement.distribute,
        groupElements: presentationElements.groupSelection,
        reorderElement: arrangement.reorder,
        ungroupElements: presentationElements.ungroupSelection,
        updateElement: (patch, options) => {
          if (
            selectedTextEditor &&
            !selectedTextEditor.state.selection.empty &&
            applyPresentationTextFormatting(selectedTextEditor, patch, {
              restoreFocus: options.restoreTextFocus,
            })
          ) {
            return;
          }
          presentationElements.updateElement(patch);
        },
        updateTextElement: presentationElements.updateTextElement,
      },
      history,
      insert: {
        enabled: Boolean(activeTargetId),
        addChart: presentationElements.addChart,
        addElement: presentationElements.addElement,
        addImage: presentationElements.addImage,
        addTable: presentationElements.addTable,
        instantiatePlaceholder: presentationElements.instantiatePlaceholder,
        requestImage: () => imageInputRef.current?.click(),
      },
      keyboard: {
        editingElementId: selection.editingElementId,
        selectedElement,
        selectedElementCount: selectedElements.length,
      },
      review: {
        canAddComment: Boolean(selectedSlide),
        addComment: presentationReview.addComment,
        closeComments: presentationReview.closeComments,
        deleteComment: presentationReview.deleteComment,
        locateComment: presentationReview.locateComment,
        openComment: presentationReview.openComment,
        toggleComments,
        updateComment: presentationReview.updateComment,
      },
      selection: {
        canDeleteSelection:
          selectedElements.length > 0 && Boolean(activeTargetId),
        canDuplicateSelection:
          selectedElements.length > 0 && Boolean(activeTargetId),
        canEditElement: (id) =>
          activeElements.some((element) => element.id === id),
        canExitEditing: Boolean(selection.editingElementId),
        canNudgeSelection:
          selectedElements.length > 0 && Boolean(activeTargetId),
        canToggleBold: textFormattingAvailable,
        canToggleItalic: textFormattingAvailable,
        canToggleUnderline: textFormattingAvailable,
        deleteSelection: clipboard.deleteSelection,
        duplicateSelection: clipboard.duplicateSelection,
        editElement: selection.edit,
        exitEditing: selection.exitEditing,
        nudgeSelection: clipboard.nudgeSelection,
        selectElement: selection.select,
        selectElements: selection.replace,
        toggleBold: () => toggleTextFormatting('bold'),
        toggleItalic: () => toggleTextFormatting('italic'),
        toggleUnderline: () => toggleTextFormatting('underline'),
      },
      slides: {
        canAddSlide: designMode === 'slide',
        canApplyTransitionToAll: (transition) =>
          designMode === 'slide' &&
          presentationSlides.canApplyTransitionToAll(transition),
        canDeleteSlide: designMode === 'slide' && content.slides.length > 1,
        canDuplicateSlide: designMode === 'slide' && Boolean(selectedSlide),
        canSetTransition: designMode === 'slide',
        addSlide: presentationSlides.addSlide,
        applyTransitionToAll: presentationSlides.applyTransitionToAll,
        deleteSlide: presentationSlides.deleteSlide,
        deleteSlideById: presentationSlides.deleteSlideById,
        duplicateSlide: presentationSlides.duplicateSlide,
        selectSlide: (slideId, returnToSlideMode) => {
          setSelectedSlideId(slideId);
          selection.clear();
          if (returnToSlideMode) setDesignMode('slide');
        },
        setBackground: presentationDesign.setActiveBackground,
        setTransition: presentationSlides.setTransition,
        updateNotes: presentationSlides.updateNotes,
      },
      view: {
        canStartSlideshow: designMode === 'slide' && content.slides.length > 0,
        setViewMode,
        startSlideshow: (source) => {
          const currentIndex = content.slides.findIndex(
            (slide) => slide.id === selectedSlide.id,
          );
          const activeElement = document.activeElement;
          slideshowReturnFocusRef.current =
            activeElement instanceof HTMLElement &&
            activeElement !== document.body
              ? activeElement
              : document.querySelector<HTMLElement>(
                  `[data-presentation-slideshow-source="${source}"]`,
                );
          setAgentMenu(null);
          setSlideshowStartIndex(
            source === 'current' ? Math.max(0, currentIndex) : 0,
          );
          onStartSlideshow?.();
        },
        toggleDesign: presentationDesign.toggleDesignPanel,
      },
    },
    presentationExtensions,
  );
  const presentationCommands = presentationEditor.commands;
  presentationCommandsRef.current = presentationCommands;
  const presentationCan = presentationEditor.can();
  const restoreObjectFocus = useCallback(
    () =>
      restorePresentationWorkspaceFocus(
        presentationRootRef.current,
        () => workspaceFocusStateRef.current,
      ),
    [],
  );
  const presentationToolbarCommands = useMemo(
    () =>
      presentationCommandsWithObjectFocus(
        presentationCommands,
        restoreObjectFocus,
      ),
    [presentationCommands, restoreObjectFocus],
  );
  useOfficeEditorKeyboardShortcuts(presentationEditor, {
    onHandled: restoreObjectFocus,
    scopeRef: presentationRootRef,
  });
  useOfficeEditorWheelZoom({
    scopeRef: presentationRootRef,
    onZoomIn: () => setZoom((current) => stepOfficeZoom(current, 'in')),
    onZoomOut: () => setZoom((current) => stepOfficeZoom(current, 'out')),
  });

  return (
    <section
      ref={presentationRootRef}
      className="work-presentation-editor"
      data-presentation-geometry-engine={geometry.engine ?? undefined}
      data-presentation-geometry-state={geometry.pending ? 'running' : 'idle'}
      data-presentation-transform-state={
        transform.dragging ? 'dragging' : 'idle'
      }
    >
      <OfficeFileInput
        ref={imageInputRef}
        accept="image/*"
        aria-label="插入图片"
        onChange={(event) => {
          const file = event.target.files?.[0];
          event.target.value = '';
          if (file) void presentationToolbarCommands.addImage(file);
        }}
      />
      <PresentationToolbar
        selectedSlide={selectedSlide}
        fileActions={fileActions}
        selectedElement={toolbarSelectedElement}
        selectedUnitCount={selectionUnits.length}
        can={presentationCan}
        textFormattingAvailable={textFormattingAvailable}
        commentsOpen={commentsOpen}
        commentCount={presentationCommentCount(content.slides)}
        designOpen={designOpen}
        editingDesign={designMode !== 'slide'}
        background={activeBackground}
        transition={selectedSlide.transition}
        viewMode={viewMode}
        commands={presentationToolbarCommands}
      />
      {designOpen && selectedLayout && selectedMaster && (
        <PresentationDesignPanel
          can={presentationCan}
          commands={presentationCommands}
          content={designContent}
          slide={selectedSlide}
          layout={selectedLayout}
          master={selectedMaster}
          mode={designMode}
        />
      )}
      {commentsOpen && designMode === 'slide' && (
        <PresentationCommentsPanel
          slides={content.slides}
          activeCommentId={activeCommentId}
          commands={presentationCommands}
        />
      )}
      {chartPaneOpen && singleSelectedElement.chart && (
        <PresentationChartPanel
          chart={singleSelectedElement.chart}
          onChange={(chart) =>
            presentationCommands.updateElement({
              chart,
              altText: chart.title || '演示图表',
            })
          }
          onDelete={presentationCommands.deleteSelection}
          onClose={selection.clear}
        />
      )}
      <PresentationWorkspace
        activeBackground={activeBackground}
        activeCommentId={activeCommentId}
        activeElements={transform.displayElements}
        aspectRatio={aspectRatio}
        canvasName={canvasName}
        canvasRef={canvasRef}
        commands={presentationCommands}
        content={content}
        designContent={designContent}
        designMode={designMode}
        editingElementId={selection.editingElementId}
        inheritedElements={inheritedElements}
        placeholderGuides={placeholderGuides}
        selectedElementIds={selection.selectedElementIds}
        selectedLayout={selectedLayout}
        selectedMaster={selectedMaster}
        selectedSlide={selectedSlide}
        viewMode={viewMode}
        zoom={zoom}
        snapGuides={transform.guides}
        onBeginDrag={transform.beginDrag}
        onContinueDrag={transform.continueDrag}
        onDragCancel={transform.cancelDrag}
        onDragEnd={transform.endDrag}
        onOpenContextMenu={openContextMenu}
        onTextEditorChange={(elementId, editor) =>
          setActiveTextEditor((current) =>
            editor
              ? { elementId, editor }
              : current?.elementId === elementId
                ? null
                : current,
          )
        }
        onTextSelectionChange={() =>
          setTextSelectionVersion((version) => version + 1)
        }
      />
      <PresentationStatusBar
        content={content}
        selectedSlide={selectedSlide}
        viewMode={viewMode}
        zoom={zoom}
        saveStatus={saveStatus}
        onViewModeChange={presentationToolbarCommands.setViewMode}
        onZoomChange={setZoom}
      />
      {designMode === 'slide' && agentMenu && (
        <WorkspaceContextMenu
          label={agentMenu.target === 'element' ? '演示对象操作' : '幻灯片操作'}
          x={agentMenu.x}
          y={agentMenu.y}
          items={[
            ...presentationCoreContextMenuItems({
              can: presentationCan,
              commands: presentationCommands,
              slideId: agentMenu.slideId,
              target: agentMenu.target,
            }),
            ...(onAgentRequest
              ? presentationAgentMenuItems(
                  agentMenu.selection,
                  agentMenu.target,
                  onAgentRequest,
                  agentMenuSlide
                    ? {
                        rewriteTargets: presentationAgentProposalTargets(
                          agentMenuSlide,
                          agentMenuElement,
                        ),
                        notesTarget:
                          presentationNotesProposalTarget(agentMenuSlide),
                        apply: (changes) => {
                          const outcome = applyPresentationAgentProposalChanges(
                            contentRef.current,
                            agentMenu.slideId,
                            changes,
                          );
                          if (outcome.result.appliedTargetIds.length)
                            presentationCommands.setPresentationContent(
                              outcome.content,
                            );
                          return outcome.result;
                        },
                      }
                    : undefined,
                ).map((item, index) =>
                  index === 0 ? { ...item, separatorBefore: true } : item,
                )
              : []),
          ]}
          onClose={() => setAgentMenu(null)}
        />
      )}
      {slideshowStartIndex !== null && (
        <div
          className="work-presentation-slideshow-layer"
          role="dialog"
          aria-label="幻灯片放映"
          aria-modal="true"
        >
          <PresentationPlayer
            autoFullscreen
            content={content}
            initialIndex={slideshowStartIndex}
            onExit={() => {
              const returnFocus = slideshowReturnFocusRef.current;
              setSlideshowStartIndex(null);
              restorePresentationSlideshowFocus(returnFocus);
            }}
          />
        </div>
      )}
      {presentationReview.dialog}
    </section>
  );
}

function restorePresentationSlideshowFocus(target: HTMLElement | null): void {
  let remainingAttempts = 30;
  const restore = () => {
    if (!target?.isConnected || remainingAttempts <= 0) return;
    remainingAttempts -= 1;
    const activeElement = document.activeElement;
    const canRestore =
      !activeElement ||
      activeElement === document.body ||
      activeElement === document.documentElement ||
      !activeElement.isConnected;
    if (canRestore) target.focus({ preventScroll: true });

    const nextActiveElement = document.activeElement;
    const slideshowOpen = Boolean(
      document.querySelector('[role="dialog"][aria-label="幻灯片放映"]'),
    );
    if (
      slideshowOpen ||
      nextActiveElement === target ||
      nextActiveElement === document.body ||
      nextActiveElement === document.documentElement ||
      !nextActiveElement?.isConnected
    ) {
      window.setTimeout(restore, 16);
    }
  };
  window.setTimeout(restore, 0);
}
