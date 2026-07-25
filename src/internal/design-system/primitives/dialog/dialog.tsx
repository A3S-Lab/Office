import { X } from 'lucide-react';
import { type ReactNode, useEffect, useId, useRef } from 'react';
import { createPortal } from 'react-dom';
import { IconButton } from '../button/icon-button';
import { useDialogFocusScope } from '../overlay/dialog-focus-scope';

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
  const backdropRef = useRef<HTMLDialogElement>(null);
  useModalIsolation(backdropRef);
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
      ref={backdropRef}
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
  return typeof document === 'undefined'
    ? dialog
    : createPortal(dialog, document.body);
}

const inertCounts = new WeakMap<HTMLElement, number>();
const inertStates = new WeakMap<
  HTMLElement,
  { attribute: boolean; property: boolean }
>();

function useModalIsolation(ref: { current: HTMLDialogElement | null }) {
  useEffect(() => {
    const modal = ref.current;
    if (!modal || !document.body) return;
    const isolated = [...document.body.children].filter(
      (element): element is HTMLElement =>
        element instanceof HTMLElement && element !== modal,
    );
    for (const element of isolated) {
      const count = inertCounts.get(element) ?? 0;
      if (count === 0) {
        inertStates.set(element, {
          attribute: element.hasAttribute('inert'),
          property: element.inert,
        });
        element.inert = true;
        element.setAttribute('inert', '');
      }
      inertCounts.set(element, count + 1);
    }
    return () => {
      for (const element of isolated) {
        const count = inertCounts.get(element) ?? 0;
        if (count > 1) {
          inertCounts.set(element, count - 1);
          continue;
        }
        inertCounts.delete(element);
        const state = inertStates.get(element);
        inertStates.delete(element);
        element.inert = state?.property ?? false;
        if (state?.attribute) element.setAttribute('inert', '');
        else element.removeAttribute('inert');
      }
    };
  }, [ref]);
}
