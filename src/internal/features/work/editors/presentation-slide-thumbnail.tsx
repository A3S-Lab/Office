import { type PointerEvent as ReactPointerEvent, useRef } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import type { WorkspaceContextMenuEvent } from '../../workspace/components/workspace-context-menu';
import { isWorkspaceContextMenuKeyboardEvent } from '../../workspace/components/workspace-context-menu';
import type { WorkPresentationDesignContent } from '../work-presentation-layouts';
import type { WorkSlide } from '../work-types';
import { useOfficeMessages } from './office-messages-context';
import { SlideCanvas } from './presentation-slide-canvas';
import { handlePresentationThumbnailKey } from './presentation-slide-thumbnail-keyboard';

export type PresentationThumbnailDropPosition = 'before' | 'after';

const REORDER_DRAG_THRESHOLD_PX = 4;

export function presentationThumbnailDropPosition(event: {
  clientY: number;
  currentTarget: EventTarget | null;
  target: EventTarget | null;
}): PresentationThumbnailDropPosition {
  const element = presentationThumbnailElement(event);
  if (!element) return 'before';
  const rect = element.getBoundingClientRect();
  const midpoint = rect.top + rect.height / 2;
  return event.clientY < midpoint ? 'before' : 'after';
}

export function presentationThumbnailDragPassedThreshold(
  startY: number,
  clientY: number,
): boolean {
  return Math.abs(clientY - startY) >= REORDER_DRAG_THRESHOLD_PX;
}

function presentationThumbnailElement(event: {
  currentTarget: EventTarget | null;
  target: EventTarget | null;
}): HTMLElement | null {
  const current = event.currentTarget;
  if (
    current instanceof HTMLElement &&
    current.hasAttribute('data-slide-thumbnail')
  ) {
    return current;
  }
  if (event.target instanceof Element) {
    const thumbnail = event.target.closest('[data-slide-thumbnail]');
    if (thumbnail instanceof HTMLElement) return thumbnail;
  }
  return current instanceof HTMLElement ? current : null;
}

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
  dropPosition = null,
  onReorderPointerDown,
  onReorderPointerMove,
  onReorderPointerUp,
  onReorderPointerCancel,
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
  dropPosition?: PresentationThumbnailDropPosition | null;
  onReorderPointerDown?: (event: ReactPointerEvent<HTMLButtonElement>) => void;
  onReorderPointerMove?: (event: ReactPointerEvent<HTMLButtonElement>) => void;
  onReorderPointerUp?: (event: ReactPointerEvent<HTMLButtonElement>) => boolean;
  onReorderPointerCancel?: (
    event: ReactPointerEvent<HTMLButtonElement>,
  ) => void;
}) {
  const messages = useOfficeMessages();
  const ignoreClickRef = useRef(false);
  return (
    <button
      type="button"
      className={selected ? 'active' : ''}
      aria-label={officeMessage(messages, 'presentation.thumb.slideAria', {
        index: String(index + 1),
        total: String(slideCount),
        name: slide.name,
      })}
      data-slide-thumbnail
      data-slide-id={slide.id}
      data-slide-index={index}
      data-slide-thumbnail-rendered={renderPreview ? 'true' : 'false'}
      data-slide-drop-position={dropPosition ?? undefined}
      onFocus={onFocus}
      onClick={() => {
        if (ignoreClickRef.current) {
          ignoreClickRef.current = false;
          return;
        }
        onSelect();
      }}
      onPointerDown={onReorderPointerDown}
      onPointerMove={onReorderPointerMove}
      onPointerUp={(event) => {
        if (onReorderPointerUp?.(event)) ignoreClickRef.current = true;
      }}
      onPointerCancel={onReorderPointerCancel}
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
