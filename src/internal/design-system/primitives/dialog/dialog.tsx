import { X } from 'lucide-react';
import { type ReactNode, useEffect, useId, useRef } from 'react';
import { createPortal } from 'react-dom';
import { IconButton } from '../button/icon-button';
import { useDialogFocusScope } from '../overlay/dialog-focus-scope';
import { officeOverlayPortalRoot } from '../overlay/portal-root';

export function Dialog({
  title,
  description,
  children,
  footer,
  onClose,
  closeDisabled = false,
  className,
  focusKey,
  restoreFocusTarget,
}: {
  title: string;
  description?: string;
  children: ReactNode;
  footer?: ReactNode;
  onClose: () => void;
  closeDisabled?: boolean;
  className?: string;
  focusKey?: string | number;
  restoreFocusTarget?: () => HTMLElement | null;
}) {
  const titleId = useId();
  const descriptionId = useId();
  const portalRootRef = useRef<HTMLElement | null>(null);
  if (!portalRootRef.current && typeof document !== 'undefined') {
    portalRootRef.current = officeOverlayPortalRoot(
      document,
      restoreFocusTarget?.(),
    );
  }
  const focusScope = useDialogFocusScope<HTMLElement>({
    onEscape: onClose,
    escapeDisabled: closeDisabled,
    restoreFocusTarget,
  });
  useEffect(() => {
    if (focusKey !== undefined) focusScope.focusInitial();
  }, [focusKey, focusScope.focusInitial]);
  const dialog = (
    <dialog
      open
      className="ds-dialog-backdrop"
      role="presentation"
      onCancel={(event) => {
        event.preventDefault();
        if (!closeDisabled) onClose();
      }}
    >
      <section
        ref={focusScope.scopeRef}
        className={`ds-dialog${className ? ` ${className}` : ''}`}
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        aria-describedby={description ? descriptionId : undefined}
        onKeyDown={focusScope.handleKeyDown}
      >
        <header>
          <div>
            <h2 id={titleId}>{title}</h2>
            {description && <p id={descriptionId}>{description}</p>}
          </div>
          <IconButton label="关闭" disabled={closeDisabled} onClick={onClose}>
            <X size={16} />
          </IconButton>
        </header>
        <div className="ds-dialog-body">{children}</div>
        {footer && <footer>{footer}</footer>}
      </section>
    </dialog>
  );
  return typeof document === 'undefined' || !portalRootRef.current
    ? dialog
    : createPortal(dialog, portalRootRef.current);
}
