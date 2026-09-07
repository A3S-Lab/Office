import type { WorkspaceContextMenuEvent } from '../../workspace/components/workspace-context-menu';
import { isWorkspaceContextMenuKeyboardEvent } from '../../workspace/components/workspace-context-menu';
import type { WorkPresentationDesignContent } from '../work-presentation-layouts';
import type { WorkSlide } from '../work-types';
import { handlePresentationThumbnailKey } from './presentation-slide-thumbnail-keyboard';
import { SlideCanvas } from './presentation-slide-canvas';

export function PresentationSlideThumbnail({
  content,
  slide,
  index,
  selected,
  slideCount,
  aspectRatio,
  variant,
  renderPreview,
  onFocus,
  onSelect,
  onDelete,
  onNavigate,
  onContextMenu,
  onDoubleClick,
}: {
  content: WorkPresentationDesignContent;
  slide: WorkSlide;
  index: number;
  selected: boolean;
  slideCount: number;
  aspectRatio: string;
  variant: 'strip' | 'sorter';
  renderPreview: boolean;
  onFocus: () => void;
  onSelect: () => void;
  onDelete: () => boolean;
  onNavigate: (index: number) => void;
  onContextMenu?: (event: WorkspaceContextMenuEvent<HTMLButtonElement>) => void;
  onDoubleClick?: () => void;
}) {
  return (
    <button
      type="button"
      className={selected ? 'active' : ''}
      aria-label={`幻灯片 ${index + 1} / ${slideCount}：${slide.name}`}
      data-slide-thumbnail
      data-slide-id={slide.id}
      data-slide-index={index}
      data-slide-thumbnail-rendered={renderPreview ? 'true' : 'false'}
      onFocus={onFocus}
      onClick={onSelect}
      onContextMenu={onContextMenu}
      onDoubleClick={onDoubleClick}
      onKeyDown={(event) => {
        if (isWorkspaceContextMenuKeyboardEvent(event) && onContextMenu) {
          event.preventDefault();
          event.stopPropagation();
          onContextMenu(event);
          return;
        }
        handlePresentationThumbnailKey(event, {
          index,
          slideCount,
          onDelete,
          onNavigate,
          onActivate: onDoubleClick,
        });
      }}
    >
      {variant === 'strip' && <span>{index + 1}</span>}
      {renderPreview ? (
        <SlideCanvas
          designContent={content}
          slide={slide}
          interactive={false}
          aspectRatio={aspectRatio}
          showPlaceholders
        />
      ) : (
        <span
          aria-hidden="true"
          className="work-slide-canvas work-slide-thumbnail-placeholder"
          style={{ aspectRatio, background: slide.background }}
        />
      )}
      {variant === 'sorter' && (
        <>
          <span>{index + 1}</span>
          <strong>{slide.name}</strong>
        </>
      )}
    </button>
  );
}
