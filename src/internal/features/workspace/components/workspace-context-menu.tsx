import {
  Fragment,
  type KeyboardEvent,
  type MouseEvent,
  type ReactNode,
} from 'react';
import { createPortal } from 'react-dom';
import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
} from 'react';
import { matchesAriaKeyShortcuts } from '../../../keyboard-shortcuts';

const CONTEXT_MENU_MARGIN = 8;

const CONTEXT_MENU_TAB_STOP_SELECTOR = [
  'a[href]:not([tabindex="-1"])',
  'button:not(:disabled):not([tabindex="-1"])',
  'input:not(:disabled):not([type="hidden"]):not([tabindex="-1"])',
  'select:not(:disabled):not([tabindex="-1"])',
  'textarea:not(:disabled):not([tabindex="-1"])',
  '[contenteditable="true"]:not([tabindex="-1"])',
  '[tabindex]:not([tabindex="-1"])',
].join(', ');

export type WorkspaceContextMenuEvent<
  Target extends HTMLElement = HTMLElement,
> = MouseEvent<Target> | KeyboardEvent<Target>;

export function isWorkspaceContextMenuKeyboardEvent({
  key,
  shiftKey,
}: Pick<KeyboardEvent<HTMLElement>, 'key' | 'shiftKey'>): boolean {
  return key === 'ContextMenu' || (key === 'F10' && shiftKey);
}

export function workspaceContextMenuPosition(
  event: WorkspaceContextMenuEvent,
): { x: number; y: number } {
  if ('clientX' in event && (event.clientX !== 0 || event.clientY !== 0)) {
    return { x: event.clientX, y: event.clientY };
  }

  const target =
    event.target instanceof HTMLElement ? event.target : event.currentTarget;
  const bounds =
    contextMenuSelectionBounds(target.ownerDocument, target) ??
    target.getBoundingClientRect();
  return {
    x: bounds.left + bounds.width / 2,
    y: bounds.top + bounds.height / 2,
  };
}

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
  className = '',
  x,
  y,
  items,
  onClose,
  onRestoreFocus,
}: {
  label: string;
  className?: string;
  x: number;
  y: number;
  items: readonly WorkspaceContextMenuItem[];
  onClose(): void;
  onRestoreFocus?(): void;
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
    const requested = contextMenuRequestedPosition(
      x,
      y,
      restoreFocusRef.current,
    );
    setPosition({
      left: clampCoordinate(
        requested.x,
        available.left + CONTEXT_MENU_MARGIN,
        available.right - bounds.width - CONTEXT_MENU_MARGIN,
      ),
      top: clampCoordinate(
        requested.y,
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
    const activeElement = menuRef.current?.ownerDocument.activeElement;
    const current = buttons.indexOf(activeElement as HTMLButtonElement);
    const next =
      current < 0 ? 0 : (current + direction + buttons.length) % buttons.length;
    buttons[next]?.focus();
  };
  const restoreContextFocus = () => {
    if (onRestoreFocus) {
      onRestoreFocus();
      return;
    }
    const restoreFocus = restoreFocusRef.current;
    if (restoreFocus?.isConnected) restoreFocus.focus({ preventScroll: true });
  };
  const dismissAndRestoreFocus = () => {
    onClose();
    restoreContextFocus();
  };
  const dismissAndMoveFocus = (direction: -1 | 1) => {
    const menu = menuRef.current;
    const restoreFocus = restoreFocusRef.current;
    const next = adjacentContextMenuTabStop(restoreFocus, direction, menu);
    const view = menu?.ownerDocument.defaultView;
    onClose();
    const focus = () => {
      const target = next?.isConnected ? next : restoreFocus;
      if (target?.isConnected) target.focus({ preventScroll: true });
    };
    if (view) view.requestAnimationFrame(focus);
    else focus();
  };

  const portalRoot = portalRootRef.current;
  if (!portalRoot) return null;

  return createPortal(
    <div
      ref={menuRef}
      className={`work-office-context-menu workspace-context-menu ${className}`.trim()}
      role="menu"
      aria-label={label}
      aria-orientation="vertical"
      data-office-shortcuts="ignore"
      style={position}
      onContextMenu={(event) => {
        event.preventDefault();
        event.stopPropagation();
      }}
      onKeyDown={(event) => {
        const shortcutItem =
          !event.repeat && !event.nativeEvent.isComposing
            ? items.find(
                (item) =>
                  !item.disabled &&
                  matchesAriaKeyShortcuts(
                    event.nativeEvent,
                    item.ariaKeyShortcut,
                  ),
              )
            : undefined;
        if (shortcutItem) {
          event.preventDefault();
          event.stopPropagation();
          dismissAndRestoreFocus();
          shortcutItem.onSelect();
        } else if (event.key === 'Escape') {
          event.preventDefault();
          event.stopPropagation();
          dismissAndRestoreFocus();
        } else if (event.key === 'Tab') {
          event.preventDefault();
          event.stopPropagation();
          dismissAndMoveFocus(event.shiftKey ? -1 : 1);
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
            tabIndex={-1}
            className={item.danger ? 'danger' : undefined}
            disabled={item.disabled}
            aria-label={item.label}
            aria-keyshortcuts={item.ariaKeyShortcut}
            onClick={() => {
              onClose();
              restoreContextFocus();
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

function contextMenuSelectionBounds(
  ownerDocument: Document,
  target: HTMLElement,
): DOMRect | null {
  const selection = ownerDocument.getSelection?.();
  if (!selection || selection.isCollapsed || selection.rangeCount === 0) {
    return null;
  }
  const range = selection.getRangeAt(0);
  if (!target.contains(range.commonAncestorContainer)) return null;
  if (typeof range.getBoundingClientRect !== 'function') return null;
  const bounds = range.getBoundingClientRect();
  return bounds.width > 0 || bounds.height > 0 ? bounds : null;
}

function contextMenuRequestedPosition(
  x: number,
  y: number,
  anchor: HTMLElement | null,
): { x: number; y: number } {
  if (x !== 0 || y !== 0 || !anchor) return { x, y };
  const bounds =
    contextMenuSelectionBounds(anchor.ownerDocument, anchor) ??
    anchor.getBoundingClientRect();
  return {
    x: bounds.left + bounds.width / 2,
    y: bounds.top + bounds.height / 2,
  };
}

function adjacentContextMenuTabStop(
  origin: HTMLElement | null,
  direction: -1 | 1,
  menu: HTMLElement | null,
): HTMLElement | null {
  if (!origin) return null;
  const ownerDocument = origin.ownerDocument;
  const view = ownerDocument.defaultView;
  const tabStops = [
    ...ownerDocument.querySelectorAll<HTMLElement>(
      CONTEXT_MENU_TAB_STOP_SELECTOR,
    ),
  ].filter(
    (element) =>
      !menu?.contains(element) &&
      !element.closest('[hidden], [inert], [aria-hidden="true"]') &&
      element.getAttribute('aria-disabled') !== 'true' &&
      view?.getComputedStyle(element).display !== 'none' &&
      view?.getComputedStyle(element).visibility !== 'hidden',
  );
  const index = tabStops.indexOf(origin);
  if (index >= 0) return tabStops[index + direction] ?? null;

  const nodeApi = view?.Node;
  if (!nodeApi) return null;
  const relation =
    direction === 1
      ? nodeApi.DOCUMENT_POSITION_FOLLOWING
      : nodeApi.DOCUMENT_POSITION_PRECEDING;
  const ordered = direction === 1 ? tabStops : [...tabStops].reverse();
  return (
    ordered.find((element) =>
      Boolean(origin.compareDocumentPosition(element) & relation),
    ) ?? null
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
