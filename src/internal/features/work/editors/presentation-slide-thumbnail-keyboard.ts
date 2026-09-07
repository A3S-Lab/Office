/**
 * Shared WPS-style keyboard contract for Presentation slide thumbnails.
 *
 * Sorter view activates the selected slide (return to normal) on Enter, matching
 * the WPS slide-sorter contract. Arrow/Home/End move between slides;
 * Delete/Backspace remove the current slide.
 */
export function handlePresentationThumbnailKey(
  event: {
    key: string;
    preventDefault: () => void;
    stopPropagation: () => void;
  },
  options: {
    index: number;
    slideCount: number;
    onDelete: () => boolean;
    onNavigate: (index: number) => void;
    onActivate?: () => void;
  },
): boolean {
  if (event.key === 'Enter' && options.onActivate) {
    event.preventDefault();
    event.stopPropagation();
    options.onActivate();
    return true;
  }
  if (event.key === 'Delete' || event.key === 'Backspace') {
    event.preventDefault();
    event.stopPropagation();
    options.onDelete();
    return true;
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
  ) {
    return false;
  }
  event.preventDefault();
  event.stopPropagation();
  const nextIndex =
    event.key === 'Home'
      ? 0
      : event.key === 'End'
        ? options.slideCount - 1
        : event.key === 'ArrowUp' || event.key === 'ArrowLeft'
          ? Math.max(0, options.index - 1)
          : Math.min(options.slideCount - 1, options.index + 1);
  options.onNavigate(nextIndex);
  return true;
}
