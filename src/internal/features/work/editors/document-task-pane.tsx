import { X } from 'lucide-react';
import { type KeyboardEvent, type ReactNode, useId } from 'react';
import { IconButton } from '../../../design-system/primitives';

export function DocumentTaskPane({
  ariaLabel,
  title,
  description,
  closeLabel,
  className,
  children,
  onClose,
}: {
  ariaLabel?: string;
  title: string;
  description?: ReactNode;
  closeLabel: string;
  className: string;
  children: ReactNode;
  onClose: () => void;
}) {
  const titleId = useId();
  const descriptionId = useId();
  const handleKeyDown = (event: KeyboardEvent<HTMLElement>) => {
    if (event.key !== 'Escape' || event.defaultPrevented) return;
    event.preventDefault();
    event.stopPropagation();
    onClose();
  };

  return (
    <aside
      className={`${className} work-document-task-pane`}
      aria-label={ariaLabel}
      aria-labelledby={ariaLabel ? undefined : titleId}
      aria-describedby={description ? descriptionId : undefined}
      onKeyDown={handleKeyDown}
    >
      <header className="work-document-task-pane-header">
        <div>
          <strong id={titleId}>{title}</strong>
          {description && <span id={descriptionId}>{description}</span>}
        </div>
        <IconButton className="close" label={closeLabel} onClick={onClose}>
          <X size={14} />
        </IconButton>
      </header>
      {children}
    </aside>
  );
}
