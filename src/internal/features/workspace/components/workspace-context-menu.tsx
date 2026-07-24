import { Fragment, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
} from 'react';

const CONTEXT_MENU_MARGIN = 8;

export interface WorkspaceContextMenuItem {
  id: string;
  label: string;
  icon: ReactNode;
  onSelect(): void;
  shortcut?: string;
  ariaKeyShortcut?: string;
  danger?: boolean;
  disabled?: boolean;
  separatorBefore?: boolean;
}

export function WorkspaceContextMenu({
  label,
  x,
  y,
  items,
  onClose,
}: {
  label: string;
  x: number;
  y: number;
  items: readonly WorkspaceContextMenuItem[];
  onClose(): void;
}) {
  const menuRef = useRef<HTMLDivElement>(null);
  const restoreFocusRef = useRef<HTMLElement | null>(null);
  const portalRootRef = useRef<HTMLElement | null>(null);
  const [position, setPosition] = useState({ left: x, top: y });
  if (!portalRootRef.current && typeof document !== 'undefined') {
    portalRootRef.current = contextMenuPortalRoot(document, x, y);
  }

  const positionMenu = useCallback(() => {
    const menu = menuRef.current;
    if (!menu) return;
    const ownerDocument = menu.ownerDocument;
    const view = ownerDocument.defaultView;
    if (!view) return;
    const bounds = menu.getBoundingClientRect();
    const available = contextMenuAvailableBounds(
      ownerDocument,
      view,
      portalRootRef.current,
    );
    setPosition({
      left: clampCoordinate(
        x,
        available.left + CONTEXT_MENU_MARGIN,
        available.right - bounds.width - CONTEXT_MENU_MARGIN,
      ),
      top: clampCoordinate(
        y,
        available.top + CONTEXT_MENU_MARGIN,
        available.bottom - bounds.height - CONTEXT_MENU_MARGIN,
      ),
    });
  }, [x, y]);

  useLayoutEffect(() => {
    const menu = menuRef.current;
    if (!menu) return;
    const ownerDocument = menu.ownerDocument;
    const view = ownerDocument.defaultView;
    if (!restoreFocusRef.current) {
      restoreFocusRef.current =
        ownerDocument.activeElement instanceof HTMLElement
          ? ownerDocument.activeElement
          : null;
    }
    positionMenu();
    menu.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus();
    view?.addEventListener('resize', positionMenu);
    view?.visualViewport?.addEventListener('resize', positionMenu);
    view?.visualViewport?.addEventListener('scroll', positionMenu);
    return () => {
      view?.removeEventListener('resize', positionMenu);
      view?.visualViewport?.removeEventListener('resize', positionMenu);
      view?.visualViewport?.removeEventListener('scroll', positionMenu);
    };
  }, [positionMenu]);

  useEffect(() => {
    const menu = menuRef.current;
    if (!menu) return;
    const ownerDocument = menu.ownerDocument;
    const view = ownerDocument.defaultView;
    const closeFromOutside = (event: PointerEvent) => {
      const currentMenu = menuRef.current;
      if (!currentMenu || event.composedPath().includes(currentMenu)) return;
      onClose();
    };
    const closeFromScroll = (event: Event) => {
      if (menuRef.current?.contains(event.target as Node)) return;
      onClose();
    };
    ownerDocument.addEventListener('pointerdown', closeFromOutside, true);
    ownerDocument.addEventListener('scroll', closeFromScroll, true);
    view?.addEventListener('blur', onClose);
    return () => {
      ownerDocument.removeEventListener('pointerdown', closeFromOutside, true);
      ownerDocument.removeEventListener('scroll', closeFromScroll, true);
      view?.removeEventListener('blur', onClose);
    };
  }, [onClose]);

  const moveFocus = (direction: 1 | -1) => {
    const buttons = [
      ...(menuRef.current?.querySelectorAll<HTMLButtonElement>(
        'button:not(:disabled)',
      ) ?? []),
    ];
    if (!buttons.length) return;
    const current = buttons.indexOf(
      document.activeElement as HTMLButtonElement,
    );
    const next =
      current < 0 ? 0 : (current + direction + buttons.length) % buttons.length;
    buttons[next]?.focus();
  };
  const dismissAndRestoreFocus = () => {
    onClose();
    const restoreFocus = restoreFocusRef.current;
    if (restoreFocus?.isConnected) restoreFocus.focus({ preventScroll: true });
  };

  const portalRoot = portalRootRef.current;
  if (!portalRoot) return null;

  return createPortal(
    <div
      ref={menuRef}
      className="workspace-context-menu"
      role="menu"
      aria-label={label}
      style={position}
      onContextMenu={(event) => {
        event.preventDefault();
        event.stopPropagation();
      }}
      onKeyDown={(event) => {
        if (event.key === 'Escape') {
          event.preventDefault();
          dismissAndRestoreFocus();
        } else if (event.key === 'Tab') {
          onClose();
        } else if (event.key === 'ArrowDown') {
          event.preventDefault();
          moveFocus(1);
        } else if (event.key === 'ArrowUp') {
          event.preventDefault();
          moveFocus(-1);
        } else if (event.key === 'Home' || event.key === 'End') {
          event.preventDefault();
          const buttons = [
            ...(menuRef.current?.querySelectorAll<HTMLButtonElement>(
              'button:not(:disabled)',
            ) ?? []),
          ];
          (event.key === 'Home' ? buttons[0] : buttons.at(-1))?.focus();
        }
      }}
    >
      {items.map((item) => (
        <Fragment key={item.id}>
          {item.separatorBefore && (
            <hr className="workspace-context-menu-separator" />
          )}
          <button
            type="button"
            role="menuitem"
            className={item.danger ? 'danger' : undefined}
            disabled={item.disabled}
            aria-label={item.label}
            aria-keyshortcuts={item.ariaKeyShortcut}
            onClick={() => {
              onClose();
              const restoreFocus = restoreFocusRef.current;
              if (restoreFocus?.isConnected) {
                restoreFocus.focus({ preventScroll: true });
              }
              item.onSelect();
            }}
          >
            {item.icon}
            <span>{item.label}</span>
            {item.shortcut && <kbd>{item.shortcut}</kbd>}
          </button>
        </Fragment>
      ))}
    </div>,
    portalRoot,
  );
}

function contextMenuPortalRoot(
  ownerDocument: Document,
  x: number,
  y: number,
): HTMLElement {
  const elementAtPoint = ownerDocument.elementFromPoint?.(x, y);
  const activeElement = ownerDocument.activeElement;
  return (
    elementAtPoint?.closest<HTMLElement>('[data-a3s-office]') ??
    (activeElement instanceof HTMLElement
      ? activeElement.closest<HTMLElement>('[data-a3s-office]')
      : null) ??
    ownerDocument.body
  );
}

function contextMenuAvailableBounds(
  ownerDocument: Document,
  view: Window,
  portalRoot: HTMLElement | null,
) {
  const viewport = view.visualViewport;
  const viewportBounds = {
    left: viewport?.offsetLeft ?? 0,
    top: viewport?.offsetTop ?? 0,
    right: (viewport?.offsetLeft ?? 0) + (viewport?.width ?? view.innerWidth),
    bottom: (viewport?.offsetTop ?? 0) + (viewport?.height ?? view.innerHeight),
  };
  if (
    !portalRoot ||
    portalRoot === ownerDocument.body ||
    portalRoot === ownerDocument.documentElement
  ) {
    return viewportBounds;
  }
  const rootBounds = portalRoot.getBoundingClientRect();
  if (rootBounds.width <= 0 || rootBounds.height <= 0) return viewportBounds;
  return {
    left: Math.max(viewportBounds.left, rootBounds.left),
    top: Math.max(viewportBounds.top, rootBounds.top),
    right: Math.min(viewportBounds.right, rootBounds.right),
    bottom: Math.min(viewportBounds.bottom, rootBounds.bottom),
  };
}

function clampCoordinate(value: number, minimum: number, maximum: number) {
  return Math.max(minimum, Math.min(value, Math.max(minimum, maximum)));
}
