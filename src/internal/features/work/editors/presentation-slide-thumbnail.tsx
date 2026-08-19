import type { WorkspaceContextMenuEvent } from '../../workspace/components/workspace-context-menu';
import { isWorkspaceContextMenuKeyboardEvent } from '../../workspace/components/workspace-context-menu';
import type { WorkPresentationDesignContent } from '../work-presentation-layouts';
import type { WorkSlide } from '../work-types';
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
        handleThumbnailKey(event, index, slideCount, onDelete, onNavigate);
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

function handleThumbnailKey(
  event: React.KeyboardEvent<HTMLButtonElement>,
  index: number,
  slideCount: number,
  onDelete: () => boolean,
  onNavigate: (index: number) => void,
): void {
  if (event.key === 'Delete' || event.key === 'Backspace') {
    event.preventDefault();
    event.stopPropagation();
    onDelete();
    return;
  }
  if (
    ![
      'ArrowUp',
      'ArrowDown',
      'ArrowLeft',
      'ArrowRight',
      'Home',
      'End',
    ].includes(event.key)
  )
    return;
  event.preventDefault();
  event.stopPropagation();
  const nextIndex =
    event.key === 'Home'
      ? 0
      : event.key === 'End'
        ? slideCount - 1
        : event.key === 'ArrowUp' || event.key === 'ArrowLeft'
          ? Math.max(0, index - 1)
          : Math.min(slideCount - 1, index + 1);
  onNavigate(nextIndex);
}
