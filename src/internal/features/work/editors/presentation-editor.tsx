import type { Editor } from '@tiptap/core';
import { useCallback, useEffect, useRef, useState } from 'react';
import { WorkspaceContextMenu } from '../../workspace/components/workspace-context-menu';
import { presentationAgentMenuItems } from '../components/work-editor-agent-menus';
import { applyPresentationAgentProposalChanges } from '../work-agent-proposal-apply';
import {
  presentationAgentProposalTargets,
  presentationAgentSelection,
  presentationNotesProposalTarget,
} from '../work-presentation-agent-context';
import { createPresentationChartElement } from '../work-presentation-charts';
import {
  applyPresentationLayout,
  presentationSlideView,
  withPresentationDesign,
} from '../work-presentation-layouts';
import { createWorkId } from '../work-templates';
import type {
  WorkPresentationLayout,
  WorkPresentationMaster,
  WorkSlide,
  WorkSlideElement,
} from '../work-types';
import { OfficeFileInput, useOfficeDialog } from './office-controls';
import { PresentationChartPanel } from './presentation-chart-panel';
import { createPresentationCommandDispatcher } from './presentation-command-controller';
import {
  PresentationCommentsPanel,
  presentationCommentCount,
} from './presentation-comments-panel';
import {
  type PresentationDesignMode,
  PresentationDesignPanel,
} from './presentation-design-panel';
import {
  clamp,
  newPresentationElement,
  newPresentationImageElement,
  newPresentationTableElement,
  newSlide,
  structuredCopy,
  updatePresentationElements,
  updateSlide,
} from './presentation-editor-operations';
import type {
  PresentationAgentMenuState,
  PresentationDragState,
  PresentationEditorProps,
} from './presentation-editor-types';
import { PresentationPlayer } from './presentation-player';
import { PresentationStatusBar } from './presentation-status-bar';
import {
  applyPresentationTextFormatting,
  type PresentationTextValue,
  presentationTextToolbarState,
} from './presentation-text-editor';
import {
  applyPresentationElementFormattingPatch,
  presentationElementToolbarState,
} from './presentation-text-formatting';
import { PresentationToolbar } from './presentation-toolbar';
import { PresentationWorkspace } from './presentation-workspace';
import { usePresentationClipboard } from './use-presentation-clipboard';
import { usePresentationGeometry } from './use-presentation-geometry';
import { usePresentationHistory } from './use-presentation-history';
import { WorkOfficePreviewBar } from './work-office-chrome';

export type { PresentationEditorProps } from './presentation-editor-types';

export function PresentationEditor({
  content,
  preview,
  saveStatus = '已自动保存',
  fileActions,
  kernelWasmUrl,
  onChange,
  onAgentRequest,
  onStartSlideshow,
}: PresentationEditorProps) {
  const contentRef = useRef(content);
  const [selectedSlideId, setSelectedSlideId] = useState(
    content.slides[0]?.id ?? '',
  );
  const [selectedElementId, setSelectedElementId] = useState<string | null>(
    null,
  );
  const [activeTextEditor, setActiveTextEditor] = useState<{
    elementId: string;
    editor: Editor;
  } | null>(null);
  const [, setTextSelectionVersion] = useState(0);
  const [commentsOpen, setCommentsOpen] = useState(false);
  const [designOpen, setDesignOpen] = useState(false);
  const [designMode, setDesignMode] = useState<PresentationDesignMode>('slide');
  const [activeCommentId, setActiveCommentId] = useState<string | null>(null);
  const [agentMenu, setAgentMenu] = useState<PresentationAgentMenuState | null>(
    null,
  );
  const [viewMode, setViewMode] = useState<'normal' | 'sorter'>('normal');
  const [zoom, setZoom] = useState(90);
  const officeDialog = useOfficeDialog();
  const canvasRef = useRef<HTMLElement>(null);
  const imageInputRef = useRef<HTMLInputElement>(null);
  const dragRef = useRef<PresentationDragState | null>(null);
  const geometry = usePresentationGeometry(kernelWasmUrl, !preview);
  const designContent = withPresentationDesign(content);
  const selectedSlide =
    designContent.slides.find((slide) => slide.id === selectedSlideId) ??
    designContent.slides[0];
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
  const selectedElement =
    activeElements.find((element) => element.id === selectedElementId) ?? null;
  const selectedTextEditor =
    activeTextEditor?.elementId === selectedElementId &&
    !activeTextEditor.editor.isDestroyed
      ? activeTextEditor.editor
      : null;
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
  const openAgentMenu = (
    event: React.MouseEvent,
    slide: WorkSlide,
    slideIndex: number,
    element?: WorkSlideElement | null,
  ) => {
    if (!onAgentRequest) return;
    event.preventDefault();
    event.stopPropagation();
    setAgentMenu({
      x: event.clientX,
      y: event.clientY,
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
      setSelectedElementId(null);
    }
  }, [content.slides, selectedSlideId]);

  const addSlide = useCallback(() => {
    const slide = newSlide(content.slides.length + 1);
    onChange({ ...content, slides: [...content.slides, slide] });
    setSelectedSlideId(slide.id);
    setSelectedElementId(null);
  }, [content, onChange]);

  const history = usePresentationHistory({
    content,
    onChange,
    selectedSlideId,
    onSelectSlide: (slideId) => {
      setSelectedSlideId(slideId);
      setSelectedElementId(null);
    },
  });

  const clipboard = usePresentationClipboard({
    content,
    preview,
    mode: designMode,
    targetId: activeTargetId,
    selectedSlide,
    selectedElement,
    onChange,
    onSelectSlide: setSelectedSlideId,
    onSelectElement: setSelectedElementId,
    onUndo: history.undo,
    onRedo: history.redo,
    onAddSlide: addSlide,
    onStartSlideshow,
  });

  if (preview) {
    const player = <PresentationPlayer content={content} />;
    if (!fileActions?.length) return player;
    return (
      <section className="work-presentation-editor preview">
        <WorkOfficePreviewBar
          ariaLabel="演示预览工具"
          label="只读预览"
          detail={`${content.slides.length} 张幻灯片`}
          fileActions={fileActions}
          className="work-presentation-ribbon"
        />
        {player}
      </section>
    );
  }
  if (!selectedSlide) return null;

  const updateElement = (patch: Partial<WorkSlideElement>) => {
    if (!selectedElementId || !activeTargetId) return;
    updatePresentationElements(
      contentRef.current,
      designMode,
      activeTargetId,
      (elements) =>
        elements.map((element) =>
          element.id === selectedElementId
            ? applyPresentationElementFormattingPatch(element, patch)
            : element,
        ),
      (next) => {
        contentRef.current = next;
        onChange(next);
      },
    );
  };

  const updateTextElement = (
    elementId: string,
    value: PresentationTextValue,
  ) => {
    if (!activeTargetId) return;
    updatePresentationElements(
      contentRef.current,
      designMode,
      activeTargetId,
      (elements) =>
        elements.map((element) =>
          element.id === elementId ? { ...element, ...value } : element,
        ),
      (next) => {
        contentRef.current = next;
        onChange(next);
      },
    );
  };

  const addElement = (type: WorkSlideElement['type']) => {
    if (type !== 'text' && type !== 'shape') return;
    const element = newPresentationElement(type);
    if (!activeTargetId) return;
    updatePresentationElements(
      content,
      designMode,
      activeTargetId,
      (elements) => [...elements, element],
      onChange,
    );
    setSelectedElementId(element.id);
  };

  const addTable = () => {
    const element = newPresentationTableElement();
    updatePresentationElements(
      content,
      designMode,
      activeTargetId ?? selectedSlide.id,
      (elements) => [...elements, element],
      onChange,
    );
    setSelectedElementId(element.id);
  };

  const addChart = () => {
    const element = createPresentationChartElement();
    updatePresentationElements(
      content,
      'slide',
      selectedSlide.id,
      (elements) => [...elements, element],
      onChange,
    );
    setSelectedElementId(element.id);
  };

  const addImage = async (file: File) => {
    const element = await newPresentationImageElement(file);
    if (!activeTargetId) return;
    updatePresentationElements(
      content,
      designMode,
      activeTargetId,
      (elements) => [...elements, element],
      onChange,
    );
    setSelectedElementId(element.id);
  };

  const reorderElement = (direction: -1 | 1) => {
    if (!selectedElementId || !activeTargetId) return;
    updatePresentationElements(
      content,
      designMode,
      activeTargetId,
      (current) => {
        const elements = [...current];
        const index = elements.findIndex(
          (element) => element.id === selectedElementId,
        );
        const target = clamp(index + direction, 0, elements.length - 1);
        if (index < 0 || index === target) return current;
        [elements[index], elements[target]] = [
          elements[target],
          elements[index],
        ];
        return elements;
      },
      onChange,
    );
  };

  const addComment = () => {
    void officeDialog
      .prompt({ title: '批注内容', multiline: true, confirmLabel: '添加批注' })
      .then((text) => {
        if (!text?.trim()) return;
        const comment = {
          id: createWorkId('slide-comment'),
          author: 'A3S Work 用户',
          initials: 'AW',
          date: new Date().toISOString(),
          text: text.trim(),
          x: clamp(
            selectedElement ? selectedElement.x + selectedElement.width : 50,
            2,
            98,
          ),
          y: clamp(selectedElement ? selectedElement.y : 50, 2, 98),
        };
        updateSlide(
          content,
          selectedSlide.id,
          (slide) => ({
            ...slide,
            comments: [...(slide.comments ?? []), comment],
          }),
          onChange,
        );
        setActiveCommentId(comment.id);
        setCommentsOpen(true);
      });
  };

  const duplicateSlide = () => {
    const copy: WorkSlide = {
      ...structuredCopy(selectedSlide),
      id: createWorkId('slide'),
      name: `${selectedSlide.name} 副本`,
      elements: selectedSlide.elements.map((element) => ({
        ...structuredCopy(element),
        id: createWorkId('element'),
      })),
    };
    const index = content.slides.findIndex(
      (slide) => slide.id === selectedSlide.id,
    );
    const slides = [...content.slides];
    slides.splice(index + 1, 0, copy);
    onChange({ ...content, slides });
    setSelectedSlideId(copy.id);
    setSelectedElementId(null);
  };

  const deleteSlideById = (slideId: string): boolean => {
    if (content.slides.length === 1) return false;
    const index = content.slides.findIndex((slide) => slide.id === slideId);
    if (index < 0) return false;
    const slides = content.slides.filter((slide) => slide.id !== slideId);
    onChange({ ...content, slides });
    setSelectedSlideId(slides[Math.min(index, slides.length - 1)].id);
    setSelectedElementId(null);
    return true;
  };
  const deleteSlide = () => deleteSlideById(selectedSlide.id);

  const beginDrag = (
    event: React.PointerEvent,
    element: WorkSlideElement,
    mode: PresentationDragState['mode'],
  ) => {
    if (event.button !== 0) return;
    event.stopPropagation();
    event.currentTarget.setPointerCapture(event.pointerId);
    setSelectedElementId(element.id);
    dragRef.current = {
      elementId: element.id,
      mode,
      pointerId: event.pointerId,
      startX: event.clientX,
      startY: event.clientY,
      originX: element.x,
      originY: element.y,
      originWidth: element.width,
      originHeight: element.height,
    };
  };

  const continueDrag = (event: React.PointerEvent) => {
    const drag = dragRef.current;
    const bounds = canvasRef.current?.getBoundingClientRect();
    if (!drag || !bounds || drag.pointerId !== event.pointerId) return;
    const dx = ((event.clientX - drag.startX) / bounds.width) * 100;
    const dy = ((event.clientY - drag.startY) / bounds.height) * 100;
    const patch =
      drag.mode === 'move'
        ? {
            x: clamp(drag.originX + dx, 0, 100 - drag.originWidth),
            y: clamp(drag.originY + dy, 0, 100 - drag.originHeight),
          }
        : {
            width: clamp(drag.originWidth + dx, 4, 100 - drag.originX),
            height: clamp(drag.originHeight + dy, 4, 100 - drag.originY),
          };
    if (!activeTargetId) return;
    updatePresentationElements(
      content,
      designMode,
      activeTargetId,
      (elements) =>
        elements.map((element) =>
          element.id === drag.elementId ? { ...element, ...patch } : element,
        ),
      onChange,
    );
  };

  const toggleDesignPanel = () => {
    if (designOpen) {
      setDesignOpen(false);
      setDesignMode('slide');
      setSelectedElementId(null);
      return;
    }
    onChange(designContent);
    setDesignOpen(true);
    setCommentsOpen(false);
  };

  const updateLayout = (
    layoutId: string,
    update: (layout: WorkPresentationLayout) => WorkPresentationLayout,
  ) => {
    onChange({
      ...designContent,
      layouts: designContent.layouts?.map((layout) =>
        layout.id === layoutId ? update(structuredCopy(layout)) : layout,
      ),
    });
  };

  const updateMaster = (
    masterId: string,
    update: (master: WorkPresentationMaster) => WorkPresentationMaster,
  ) => {
    onChange({
      ...designContent,
      masters: designContent.masters?.map((master) =>
        master.id === masterId ? update(structuredCopy(master)) : master,
      ),
    });
  };

  const setActiveBackground = (background: string) => {
    if (designMode === 'layout' && selectedLayout) {
      updateLayout(selectedLayout.id, (layout) => ({ ...layout, background }));
      return;
    }
    if (designMode === 'master' && selectedMaster) {
      updateMaster(selectedMaster.id, (master) => ({ ...master, background }));
      return;
    }
    updateSlide(
      designContent,
      selectedSlide.id,
      (slide) => ({ ...slide, background, useLayoutBackground: false }),
      onChange,
    );
  };

  const createLayout = (copyCurrent: boolean) => {
    if (!selectedMaster || !selectedLayout) return;
    const id = createWorkId('layout');
    const layout: WorkPresentationLayout = copyCurrent
      ? {
          ...structuredCopy(selectedLayout),
          id,
          name: `${selectedLayout.name} 副本`,
          elements: selectedLayout.elements.map((element) => ({
            ...structuredCopy(element),
            id: createWorkId('element'),
          })),
        }
      : {
          id,
          name: `自定义布局 ${(designContent.layouts?.length ?? 0) + 1}`,
          masterId: selectedMaster.id,
          elements: [],
        };
    const next = applyPresentationLayout(
      {
        ...designContent,
        layouts: [...(designContent.layouts ?? []), layout],
      },
      selectedSlide.id,
      id,
    );
    onChange(next);
    setDesignMode('layout');
    setSelectedElementId(null);
  };

  const deleteLayout = () => {
    if (!selectedLayout || (designContent.layouts?.length ?? 0) < 2) return;
    const fallback = designContent.layouts?.find(
      (layout) => layout.id !== selectedLayout.id,
    );
    if (!fallback) return;
    onChange({
      ...designContent,
      layouts: designContent.layouts?.filter(
        (layout) => layout.id !== selectedLayout.id,
      ),
      slides: designContent.slides.map((slide) =>
        slide.layoutId === selectedLayout.id
          ? { ...slide, layoutId: fallback.id }
          : slide,
      ),
    });
    setDesignMode('slide');
    setSelectedElementId(null);
  };

  const addPlaceholder = (type: 'title' | 'body') => {
    if (designMode === 'slide' || !activeTargetId) return;
    const count = activeElements.filter(
      (element) => element.placeholder?.type === type,
    ).length;
    const prompt = type === 'title' ? '单击添加标题' : '单击添加内容';
    const element: WorkSlideElement = {
      id: createWorkId('element'),
      type: 'text',
      x: type === 'title' ? 8 : 10,
      y: type === 'title' ? 9 : 24,
      width: type === 'title' ? 84 : 80,
      height: type === 'title' ? 12 : 58,
      text: prompt,
      fontSize: type === 'title' ? 30 : 20,
      color: '#172033',
      fill: 'transparent',
      bold: type === 'title',
      align: 'left',
      placeholder: {
        key: count ? `type:${type}:${count + 1}` : `type:${type}`,
        type,
        prompt,
      },
    };
    updatePresentationElements(
      designContent,
      designMode,
      activeTargetId,
      (elements) => [...elements, element],
      onChange,
    );
    setSelectedElementId(element.id);
  };
  const presentationCommands = createPresentationCommandDispatcher({
    addChart,
    addComment,
    addElement,
    addSlide,
    addTable,
    alignElement: async (alignment) => {
      if (!selectedElement || !activeTargetId) return;
      const elementId = selectedElement.id;
      const targetId = activeTargetId;
      const targetMode = designMode;
      const aligned = await geometry.alignElement(selectedElement, alignment);
      if (!aligned) return;
      updatePresentationElements(
        contentRef.current,
        targetMode,
        targetId,
        (elements) =>
          elements.map((element) =>
            element.id === elementId
              ? { ...element, x: aligned.x, y: aligned.y }
              : element,
          ),
        onChange,
      );
    },
    applyTransitionToAll: () =>
      onChange({
        ...content,
        slides: content.slides.map((slide) => ({
          ...slide,
          transition: selectedSlide.transition
            ? structuredCopy(selectedSlide.transition)
            : undefined,
        })),
      }),
    copySelection: clipboard.copySelection,
    cutSelection: clipboard.cutSelection,
    deleteSlide,
    duplicateSlide,
    pasteSelection: clipboard.pasteSelection,
    redo: history.redo,
    reorderElement,
    requestImage: () => imageInputRef.current?.click(),
    setBackground: setActiveBackground,
    setTransition: (transition) =>
      updateSlide(
        content,
        selectedSlide.id,
        (slide) => ({ ...slide, transition }),
        onChange,
      ),
    setViewMode,
    startSlideshow: () => onStartSlideshow?.(),
    toggleComments: () => setCommentsOpen((value) => !value),
    toggleDesign: toggleDesignPanel,
    undo: history.undo,
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
      updateElement(patch);
    },
  });

  return (
    <section
      className="work-presentation-editor"
      data-presentation-geometry-engine={geometry.engine ?? undefined}
      data-presentation-geometry-state={geometry.pending ? 'running' : 'idle'}
    >
      <OfficeFileInput
        ref={imageInputRef}
        accept="image/*"
        aria-label="插入图片"
        onChange={(event) => {
          const file = event.target.files?.[0];
          event.target.value = '';
          if (file) void addImage(file);
        }}
      />
      <PresentationToolbar
        selectedSlide={selectedSlide}
        fileActions={fileActions}
        selectedElement={toolbarSelectedElement}
        slideCount={content.slides.length}
        canUndo={history.canUndo}
        canRedo={history.canRedo}
        commentsOpen={commentsOpen}
        commentCount={presentationCommentCount(content.slides)}
        designOpen={designOpen}
        editingDesign={designMode !== 'slide'}
        background={activeBackground}
        transition={selectedSlide.transition}
        canStartSlideshow={Boolean(onStartSlideshow)}
        viewMode={viewMode}
        onCommand={presentationCommands}
      />
      {designOpen && selectedLayout && selectedMaster && (
        <PresentationDesignPanel
          content={designContent}
          slide={selectedSlide}
          layout={selectedLayout}
          master={selectedMaster}
          mode={designMode}
          onApplyLayout={(layoutId) => {
            onChange(
              applyPresentationLayout(
                designContent,
                selectedSlide.id,
                layoutId,
              ),
            );
            setDesignMode('slide');
            setSelectedElementId(null);
          }}
          onToggleLayoutBackground={(enabled) =>
            updateSlide(
              designContent,
              selectedSlide.id,
              (slide) => ({ ...slide, useLayoutBackground: enabled }),
              onChange,
            )
          }
          onEditLayout={() => {
            setDesignMode('layout');
            setSelectedElementId(null);
          }}
          onEditMaster={() => {
            setDesignMode('master');
            setSelectedElementId(null);
          }}
          onCreateLayout={() => createLayout(false)}
          onDuplicateLayout={() => createLayout(true)}
          onDeleteLayout={deleteLayout}
          onRenameLayout={(name) =>
            updateLayout(selectedLayout.id, (layout) => ({ ...layout, name }))
          }
          onRenameMaster={(name) =>
            updateMaster(selectedMaster.id, (master) => ({ ...master, name }))
          }
          onSetLayoutBackground={(background) =>
            updateLayout(selectedLayout.id, (layout) => ({
              ...layout,
              background,
            }))
          }
          onSetMasterBackground={(background) =>
            updateMaster(selectedMaster.id, (master) => ({
              ...master,
              background,
            }))
          }
          onAddPlaceholder={addPlaceholder}
          onReturnToSlide={() => {
            setDesignMode('slide');
            setSelectedElementId(null);
          }}
          onClose={() => {
            setDesignOpen(false);
            setDesignMode('slide');
            setSelectedElementId(null);
          }}
        />
      )}
      {commentsOpen && designMode === 'slide' && (
        <PresentationCommentsPanel
          slides={content.slides}
          activeCommentId={activeCommentId}
          onLocate={(slideId, commentId) => {
            setSelectedSlideId(slideId);
            setSelectedElementId(null);
            setActiveCommentId(commentId);
          }}
          onChange={(slideId, commentId, text) =>
            updateSlide(
              content,
              slideId,
              (slide) => ({
                ...slide,
                comments: slide.comments?.map((comment) =>
                  comment.id === commentId ? { ...comment, text } : comment,
                ),
              }),
              onChange,
            )
          }
          onDelete={(slideId, commentId) => {
            updateSlide(
              content,
              slideId,
              (slide) => ({
                ...slide,
                comments: slide.comments?.filter(
                  (comment) => comment.id !== commentId,
                ),
              }),
              onChange,
            );
            if (activeCommentId === commentId) setActiveCommentId(null);
          }}
          onClose={() => setCommentsOpen(false)}
        />
      )}
      {designMode === 'slide' &&
        selectedElement?.type === 'chart' &&
        selectedElement.chart && (
          <PresentationChartPanel
            chart={selectedElement.chart}
            onChange={(chart) =>
              updateElement({ chart, altText: chart.title || '演示图表' })
            }
            onDelete={() => {
              updatePresentationElements(
                content,
                'slide',
                selectedSlide.id,
                (elements) =>
                  elements.filter(
                    (element) => element.id !== selectedElement.id,
                  ),
                onChange,
              );
              setSelectedElementId(null);
            }}
            onClose={() => setSelectedElementId(null)}
          />
        )}
      <PresentationWorkspace
        activeBackground={activeBackground}
        activeCommentId={activeCommentId}
        activeElements={activeElements}
        aspectRatio={aspectRatio}
        canvasName={canvasName}
        canvasRef={canvasRef}
        content={content}
        designContent={designContent}
        designMode={designMode}
        inheritedElements={inheritedElements}
        placeholderGuides={placeholderGuides}
        selectedElementId={selectedElementId}
        selectedLayout={selectedLayout}
        selectedMaster={selectedMaster}
        selectedSlide={selectedSlide}
        viewMode={viewMode}
        zoom={zoom}
        onAddSlide={addSlide}
        onBeginDrag={beginDrag}
        onContinueDrag={continueDrag}
        onDeleteSlide={deleteSlideById}
        onDragEnd={() => {
          dragRef.current = null;
        }}
        onInstantiatePlaceholder={(definition) => {
          const element: WorkSlideElement = {
            ...structuredCopy(definition),
            id: createWorkId('element'),
            text: '',
            textRuns: undefined,
          };
          updatePresentationElements(
            designContent,
            'slide',
            selectedSlide.id,
            (elements) => [...elements, element],
            onChange,
          );
          setSelectedElementId(element.id);
        }}
        onOpenAgentMenu={openAgentMenu}
        onOpenComment={(commentId) => {
          setActiveCommentId(commentId);
          setCommentsOpen(true);
        }}
        onSelectElement={setSelectedElementId}
        onSelectSlide={(slideId, returnToSlideMode) => {
          setSelectedSlideId(slideId);
          setSelectedElementId(null);
          if (returnToSlideMode) setDesignMode('slide');
        }}
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
        onUpdateElement={updateElement}
        onUpdateNotes={(notes) =>
          updateSlide(
            content,
            selectedSlide.id,
            (slide) => ({ ...slide, notes }),
            onChange,
          )
        }
        onUpdateTextElement={updateTextElement}
        onViewModeChange={setViewMode}
      />
      <PresentationStatusBar
        content={content}
        selectedSlide={selectedSlide}
        viewMode={viewMode}
        zoom={zoom}
        saveStatus={saveStatus}
        onViewModeChange={setViewMode}
        onZoomChange={setZoom}
      />
      {designMode === 'slide' && agentMenu && onAgentRequest && (
        <WorkspaceContextMenu
          label={
            agentMenu.target === 'element'
              ? '演示元素 AI 操作'
              : '幻灯片 AI 操作'
          }
          x={agentMenu.x}
          y={agentMenu.y}
          items={presentationAgentMenuItems(
            agentMenu.selection,
            agentMenu.target,
            onAgentRequest,
            agentMenuSlide
              ? {
                  rewriteTargets: presentationAgentProposalTargets(
                    agentMenuSlide,
                    agentMenuElement,
                  ),
                  notesTarget: presentationNotesProposalTarget(agentMenuSlide),
                  apply: (changes) => {
                    const outcome = applyPresentationAgentProposalChanges(
                      contentRef.current,
                      agentMenu.slideId,
                      changes,
                    );
                    if (outcome.result.appliedTargetIds.length)
                      onChange(outcome.content);
                    return outcome.result;
                  },
                }
              : undefined,
          )}
          onClose={() => setAgentMenu(null)}
        />
      )}
      {officeDialog.dialog}
    </section>
  );
}
