import { Plus, X } from 'lucide-react';
import {
  type CSSProperties,
  type PointerEvent as ReactPointerEvent,
  type RefObject,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import type { WorkspaceContextMenuEvent } from '../../workspace/components/workspace-context-menu';
import type { WorkPresentationDesignContent } from '../work-presentation-layouts';
import type {
  WorkPresentationContent,
  WorkSlide,
  WorkSlideElement,
} from '../work-types';
import { useOfficeMessages } from './office-messages-context';
import {
  PresentationSlideThumbnail,
  type PresentationThumbnailDropPosition,
  presentationThumbnailDragPassedThreshold,
  presentationThumbnailDropPosition,
} from './presentation-slide-thumbnail';
import { usePresentationThumbnailVisibility } from './use-presentation-thumbnail-visibility';
import { usePresentationThumbnailWindow } from './use-presentation-thumbnail-window';

export function PresentationThumbnailRail({
  aspectRatio,
  content,
  designContent,
  selectedSlide,
  viewMode,
  zoom,
  mobileCloseButtonRef,
  mobileNavigationModal = false,
  mobileNavigationId,
  onAddSlide,
  onCloseMobileNavigation,
  onDeleteSlide,
  onMoveSlide,
  onOpenContextMenu,
  onSelectSlide,
  onViewModeChange,
}: {
  aspectRatio: string;
  content: WorkPresentationContent;
  designContent: WorkPresentationDesignContent;
  selectedSlide: WorkSlide;
  viewMode: 'normal' | 'sorter';
  zoom: number;
  mobileCloseButtonRef?: RefObject<HTMLButtonElement | null>;
  mobileNavigationModal?: boolean;
  mobileNavigationId?: string;
  onAddSlide: () => void;
  onCloseMobileNavigation?: () => void;
  onDeleteSlide: (slideId: string) => boolean;
  onMoveSlide: (fromIndex: number, insertionIndex: number) => boolean;
  onOpenContextMenu: (
    event: WorkspaceContextMenuEvent,
    slide: WorkSlide,
    slideIndex: number,
    element?: WorkSlideElement | null,
  ) => void;
  onSelectSlide: (
    slideId: string,
    returnToSlideMode: boolean,
    dismissMobileNavigation?: boolean,
  ) => void;
  onViewModeChange: (mode: 'normal' | 'sorter') => void;
}) {
  const slideIds = useMemo(
    () => content.slides.map((slide) => slide.id),
    [content.slides],
  );
  const { viewportRef, visibleIds } = usePresentationThumbnailVisibility(
    slideIds,
    viewMode,
  );
  const thumbnailWindow = usePresentationThumbnailWindow({
    slideIds,
    selectedSlideId: selectedSlide.id,
    viewMode,
    viewportRef,
    zoom,
  });
  const visibleSlides = content.slides.slice(
    thumbnailWindow.start,
    thumbnailWindow.end,
  );

  const selectByIndex = (index: number, returnToSlideMode: boolean) => {
    const slide = content.slides[index];
    if (!slide) return;
    onSelectSlide(slide.id, returnToSlideMode, false);
    thumbnailWindow.requestFocus(index);
  };

  const deleteAndRetainFocus = (slide: WorkSlide, index: number) => {
    const nextFocusSlide =
      content.slides[index + 1] ?? content.slides[index - 1];
    if (!onDeleteSlide(slide.id)) return false;
    if (nextFocusSlide) thumbnailWindow.requestFocusById(nextFocusSlide.id);
    return true;
  };

  const dragRef = useRef<{
    pointerId: number;
    fromIndex: number;
    startY: number;
    active: boolean;
  } | null>(null);
  const [reorderArmed, setReorderArmed] = useState(false);
  const [dropIndicator, setDropIndicator] = useState<{
    index: number;
    position: PresentationThumbnailDropPosition;
  } | null>(null);
  const clearReorder = () => {
    dragRef.current = null;
    setDropIndicator(null);
    setReorderArmed(false);
  };

  useEffect(() => {
    if (!reorderArmed) return;
    const cancelIfReleasedOutside = (event: PointerEvent) => {
      const drag = dragRef.current;
      if (!drag || drag.pointerId !== event.pointerId) return;
      const landedOnThumbnail =
        event.target instanceof Element &&
        Boolean(event.target.closest('[data-slide-thumbnail]'));
      if (landedOnThumbnail) return;
      clearReorder();
    };
    window.addEventListener('pointerup', cancelIfReleasedOutside);
    window.addEventListener('pointercancel', cancelIfReleasedOutside);
    return () => {
      window.removeEventListener('pointerup', cancelIfReleasedOutside);
      window.removeEventListener('pointercancel', cancelIfReleasedOutside);
    };
  }, [reorderArmed]);

  const beginReorder = (
    index: number,
    event: ReactPointerEvent<HTMLButtonElement>,
  ) => {
    if (event.button !== 0) return;
    dragRef.current = {
      pointerId: event.pointerId,
      fromIndex: index,
      startY: event.clientY,
      active: false,
    };
    setReorderArmed(true);
  };

  const hoverReorder = (
    index: number,
    event: ReactPointerEvent<HTMLButtonElement>,
  ) => {
    const drag = dragRef.current;
    if (!drag || drag.pointerId !== event.pointerId) return;
    if (
      !drag.active &&
      !presentationThumbnailDragPassedThreshold(drag.startY, event.clientY)
    ) {
      return;
    }
    drag.active = true;
    const position = presentationThumbnailDropPosition(event);
    setDropIndicator((current) =>
      current?.index === index && current.position === position
        ? current
        : { index, position },
    );
  };

  const finishReorder = (
    index: number,
    event: ReactPointerEvent<HTMLButtonElement>,
  ): boolean => {
    const drag = dragRef.current;
    if (!drag || drag.pointerId !== event.pointerId) return false;
    const fromIndex = drag.fromIndex;
    const active =
      drag.active ||
      presentationThumbnailDragPassedThreshold(drag.startY, event.clientY);
    const position = presentationThumbnailDropPosition(event);
    clearReorder();
    if (!active) return false;
    const insertionIndex = position === 'before' ? index : index + 1;
    onMoveSlide(fromIndex, insertionIndex);
    return true;
  };

  const cancelReorder = (event: ReactPointerEvent<HTMLButtonElement>) => {
    const drag = dragRef.current;
    if (!drag || drag.pointerId !== event.pointerId) return;
    clearReorder();
  };

  const thumbnails = visibleSlides.map((slide, visibleIndex) => {
    const index = thumbnailWindow.start + visibleIndex;
    return (
      <PresentationSlideThumbnail
        key={slide.id}
        aspectRatio={aspectRatio}
        content={designContent}
        index={index}
        renderPreview={
          slide.id === selectedSlide.id || visibleIds.has(slide.id)
        }
        selected={slide.id === selectedSlide.id}
        slide={slide}
        slideCount={content.slides.length}
        variant={viewMode === 'sorter' ? 'sorter' : 'strip'}
        onContextMenu={(event) => {
          onSelectSlide(slide.id, false);
          onOpenContextMenu(event, slide, index);
        }}
        onDelete={() => deleteAndRetainFocus(slide, index)}
        dropPosition={
          dropIndicator?.index === index ? dropIndicator.position : null
        }
        onReorderPointerDown={(event) => beginReorder(index, event)}
        onReorderPointerMove={(event) => hoverReorder(index, event)}
        onReorderPointerUp={(event) => finishReorder(index, event)}
        onReorderPointerCancel={cancelReorder}
        onDoubleClick={
          viewMode === 'sorter' ? () => onViewModeChange('normal') : undefined
        }
        onFocus={() => onSelectSlide(slide.id, false, false)}
        onNavigate={(nextIndex) =>
          selectByIndex(nextIndex, viewMode === 'normal')
        }
        onSelect={() => onSelectSlide(slide.id, viewMode === 'normal')}
      />
    );
  });

  const list = (
    <>
      <ThumbnailSpacer
        height={thumbnailWindow.topSpacerHeight}
        position="before"
      />
      {thumbnails}
      <ThumbnailSpacer
        height={thumbnailWindow.bottomSpacerHeight}
        position="after"
      />
    </>
  );
  const messages = useOfficeMessages();
  const mobileModalAttributes = mobileNavigationModal
    ? ({ role: 'dialog', 'aria-modal': true } as const)
    : {};

  if (viewMode === 'sorter') {
    return (
      <section
        ref={viewportRef}
        className="work-presentation-sorter"
        aria-label={officeMessage(messages, 'presentation.thumb.sorterAria')}
        data-slide-count={content.slides.length}
        data-slide-window-end={thumbnailWindow.end}
        data-slide-window-start={thumbnailWindow.start}
        data-slide-windowed={thumbnailWindow.windowed ? 'true' : 'false'}
        style={
          {
            '--work-presentation-sorter-width': `${Math.round(220 * (zoom / 100))}px`,
          } as CSSProperties
        }
      >
        <div
          className="work-presentation-sorter-grid"
          data-slide-thumbnail-list
        >
          {list}
        </div>
      </section>
    );
  }

  return (
    <aside
      {...mobileModalAttributes}
      ref={viewportRef}
      id={mobileNavigationId}
      className="work-slide-strip"
      aria-label={officeMessage(messages, 'presentation.thumb.railAria')}
      data-slide-count={content.slides.length}
      data-slide-window-end={thumbnailWindow.end}
      data-slide-window-start={thumbnailWindow.start}
      data-slide-windowed={thumbnailWindow.windowed ? 'true' : 'false'}
    >
      <header className="work-slide-strip-header">
        <strong>
          {officeMessage(messages, 'presentation.thumb.railTitle')}
        </strong>
        <button
          ref={mobileCloseButtonRef}
          type="button"
          aria-label={officeMessage(messages, 'presentation.thumb.closeNav')}
          onClick={onCloseMobileNavigation}
        >
          <X size={16} />
        </button>
      </header>
      <div className="work-slide-thumbnail-list" data-slide-thumbnail-list>
        {list}
      </div>
      <button type="button" className="work-slide-add" onClick={onAddSlide}>
        <Plus size={15} />
        {officeMessage(messages, 'presentation.thumb.addSlide')}
      </button>
    </aside>
  );
}

function ThumbnailSpacer({
  height,
  position,
}: {
  height: number;
  position: 'before' | 'after';
}) {
  if (height <= 0) return null;
  return (
    <span
      aria-hidden="true"
      className="work-slide-thumbnail-spacer"
      data-slide-thumbnail-spacer={position}
      style={{ height: `${height}px` }}
    />
  );
}
