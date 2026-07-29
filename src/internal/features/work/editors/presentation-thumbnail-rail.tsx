import { Plus, X } from 'lucide-react';
import { useMemo, type CSSProperties, type RefObject } from 'react';
import type { WorkspaceContextMenuEvent } from '../../workspace/components/workspace-context-menu';
import type {
  WorkPresentationContent,
  WorkSlide,
  WorkSlideElement,
} from '../work-types';
import { PresentationSlideThumbnail } from './presentation-slide-thumbnail';
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
  onOpenContextMenu,
  onSelectSlide,
  onViewModeChange,
}: {
  aspectRatio: string;
  content: WorkPresentationContent;
  designContent: WorkPresentationContent;
  selectedSlide: WorkSlide;
  viewMode: 'normal' | 'sorter';
  zoom: number;
  mobileCloseButtonRef?: RefObject<HTMLButtonElement | null>;
  mobileNavigationModal?: boolean;
  mobileNavigationId?: string;
  onAddSlide: () => void;
  onCloseMobileNavigation?: () => void;
  onDeleteSlide: (slideId: string) => boolean;
  onOpenContextMenu: (
    event: WorkspaceContextMenuEvent,
    slide: WorkSlide,
    slideIndex: number,
    element?: WorkSlideElement | null,
  ) => void;
  onSelectSlide: (slideId: string, returnToSlideMode: boolean) => void;
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
    onSelectSlide(slide.id, returnToSlideMode);
    thumbnailWindow.requestFocus(index);
  };

  const deleteAndRetainFocus = (slide: WorkSlide, index: number) => {
    const nextFocusSlide =
      content.slides[index + 1] ?? content.slides[index - 1];
    if (!onDeleteSlide(slide.id)) return false;
    if (nextFocusSlide) thumbnailWindow.requestFocusById(nextFocusSlide.id);
    return true;
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
        onDoubleClick={
          viewMode === 'sorter' ? () => onViewModeChange('normal') : undefined
        }
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
  const mobileModalAttributes = mobileNavigationModal
    ? ({ role: 'dialog', 'aria-modal': true } as const)
    : {};

  if (viewMode === 'sorter') {
    return (
      <section
        ref={viewportRef}
        className="work-presentation-sorter"
        aria-label="幻灯片浏览视图"
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
      aria-label="幻灯片"
      data-slide-count={content.slides.length}
      data-slide-window-end={thumbnailWindow.end}
      data-slide-window-start={thumbnailWindow.start}
      data-slide-windowed={thumbnailWindow.windowed ? 'true' : 'false'}
    >
      <header className="work-slide-strip-header">
        <strong>幻灯片</strong>
        <button
          ref={mobileCloseButtonRef}
          type="button"
          aria-label="关闭幻灯片导航"
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
        添加幻灯片
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
