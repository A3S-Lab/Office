import {
  type KeyboardEvent as ReactKeyboardEvent,
  useCallback,
  useEffect,
  useRef,
} from 'react';

const focusableSelector = [
  'button:not(:disabled)',
  'input:not(:disabled):not([type="hidden"])',
  'textarea:not(:disabled)',
  'select:not(:disabled)',
  'a[href]',
  '[contenteditable="true"]',
  '[tabindex]:not([tabindex="-1"])',
].join(', ');

export interface DialogFocusScopeOptions {
  active?: boolean;
  onEscape?: () => void;
  escapeDisabled?: boolean;
  passThroughCommandKeys?: readonly string[];
  initialFocus?: () => HTMLElement | null;
  getActiveScope?: () => HTMLElement | null;
  getIsolationExceptions?: () => readonly (HTMLElement | null | undefined)[];
  restoreFocus?: boolean;
  restoreFocusTarget?: () => HTMLElement | null;
}

interface ActiveFocusScope {
  focusInitial: () => void;
}

interface FocusScopeKeyboardEvent {
  key: string;
  shiftKey: boolean;
  metaKey: boolean;
  ctrlKey: boolean;
  altKey: boolean;
  defaultPrevented: boolean;
  preventDefault(): void;
  stopPropagation(): void;
}

const activeFocusScopes: ActiveFocusScope[] = [];
const inertCounts = new WeakMap<HTMLElement, number>();
const inertStates = new WeakMap<
  HTMLElement,
  { attribute: boolean; property: boolean }
>();

export function useDialogFocusScope<T extends HTMLElement>({
  active = true,
  onEscape,
  escapeDisabled = false,
  passThroughCommandKeys,
  initialFocus,
  getActiveScope,
  getIsolationExceptions,
  restoreFocus = true,
  restoreFocusTarget,
}: DialogFocusScopeOptions = {}) {
  const scopeRef = useRef<T>(null);
  const activeScopeRef = useRef<ActiveFocusScope | null>(null);
  const optionsRef = useRef<DialogFocusScopeOptions>({});
  optionsRef.current = {
    active,
    onEscape,
    escapeDisabled,
    passThroughCommandKeys,
    initialFocus,
    getActiveScope,
    getIsolationExceptions,
    restoreFocus,
    restoreFocusTarget,
  };

  const focusInitial = useCallback(
    (
      scope = optionsRef.current.getActiveScope?.() ??
        (scopeRef.current as HTMLElement | null),
    ) => {
      if (!scope) return;
      const configured = optionsRef.current.initialFocus?.();
      const target =
        configured && isAvailable(configured)
          ? configured
          : defaultInitialFocus(scope);
      target?.focus();
    },
    [],
  );

  useEffect(() => {
    if (!active) return;
    const scope =
      optionsRef.current.getActiveScope?.() ??
      (scopeRef.current as HTMLElement | null);
    if (!scope) return;
    const restoreTarget =
      optionsRef.current.restoreFocusTarget?.() ?? activeElement();
    const activeScope: ActiveFocusScope = {
      focusInitial,
    };
    activeScopeRef.current = activeScope;
    activeFocusScopes.push(activeScope);
    const releaseIsolation = isolateOutsideScope(
      scope,
      optionsRef.current.getIsolationExceptions?.() ?? [],
    );
    focusInitial();

    const handleDocumentKeyDown = (event: globalThis.KeyboardEvent) => {
      if (!isTopScope(activeScope)) return;
      const currentScope =
        optionsRef.current.getActiveScope?.() ??
        (scopeRef.current as HTMLElement | null);
      if (!currentScope) return;
      if (
        event.target instanceof Node &&
        currentScope.contains(event.target) &&
        event.key !== 'Escape' &&
        event.key !== 'Tab'
      ) {
        return;
      }
      handleScopedKeyDown(event, currentScope, optionsRef.current);
    };
    const handleDocumentFocus = (event: FocusEvent) => {
      if (!isTopScope(activeScope)) return;
      const currentScope =
        optionsRef.current.getActiveScope?.() ??
        (scopeRef.current as HTMLElement | null);
      if (
        currentScope &&
        event.target instanceof Node &&
        !currentScope.contains(event.target)
      ) {
        focusInitial(currentScope);
      }
    };
    const recoverFocusAfterContentChange = new MutationObserver(() => {
      if (!isTopScope(activeScope)) return;
      const currentScope =
        optionsRef.current.getActiveScope?.() ??
        (scopeRef.current as HTMLElement | null);
      const currentActiveElement = activeElement();
      if (
        currentScope?.isConnected &&
        (!currentActiveElement || !currentScope.contains(currentActiveElement))
      ) {
        focusInitial(currentScope);
      }
    });
    recoverFocusAfterContentChange.observe(scope, {
      childList: true,
      subtree: true,
    });
    document.addEventListener('keydown', handleDocumentKeyDown, true);
    document.addEventListener('focusin', handleDocumentFocus, true);

    return () => {
      recoverFocusAfterContentChange.disconnect();
      document.removeEventListener('keydown', handleDocumentKeyDown, true);
      document.removeEventListener('focusin', handleDocumentFocus, true);
      const wasTopScope = isTopScope(activeScope);
      const index = activeFocusScopes.indexOf(activeScope);
      if (index >= 0) activeFocusScopes.splice(index, 1);
      if (activeScopeRef.current === activeScope) activeScopeRef.current = null;
      releaseIsolation();
      if (!wasTopScope) return;
      const currentRestoreTarget =
        optionsRef.current.restoreFocusTarget?.() ?? restoreTarget;
      if (
        optionsRef.current.restoreFocus !== false &&
        currentRestoreTarget?.isConnected &&
        isAvailable(currentRestoreTarget)
      ) {
        currentRestoreTarget.focus({ preventScroll: true });
        return;
      }
      activeFocusScopes.at(-1)?.focusInitial();
    };
  }, [active, focusInitial]);

  const handleKeyDown = useCallback(
    (event: ReactKeyboardEvent<HTMLElement>) => {
      const activeScope = activeScopeRef.current;
      if (!activeScope || !isTopScope(activeScope) || event.defaultPrevented)
        return;
      const scope = optionsRef.current.getActiveScope?.() ?? scopeRef.current;
      if (scope) handleScopedKeyDown(event, scope, optionsRef.current);
    },
    [],
  );

  return { scopeRef, handleKeyDown, focusInitial };
}

function handleScopedKeyDown(
  event: FocusScopeKeyboardEvent,
  scope: HTMLElement,
  options: DialogFocusScopeOptions,
): void {
  if (event.key === 'Escape' && options.onEscape && !options.escapeDisabled) {
    event.preventDefault();
    event.stopPropagation();
    options.onEscape();
    return;
  }
  if (event.key !== 'Tab') {
    const commandKey = event.metaKey || event.ctrlKey;
    const normalizedKey = event.key.toLocaleLowerCase();
    const passThroughCommand = Boolean(
      commandKey && options.passThroughCommandKeys?.includes(normalizedKey),
    );
    if (
      commandKey &&
      !passThroughCommand &&
      ['f', 'h', 'k', 'n', 'p', 's'].includes(normalizedKey)
    )
      event.preventDefault();
    if ((commandKey || event.altKey) && !passThroughCommand)
      event.stopPropagation();
    return;
  }

  const focusable = focusableElements(scope);
  if (!focusable.length) {
    event.preventDefault();
    event.stopPropagation();
    scope.focus();
    return;
  }

  const activeIndex = focusable.indexOf(document.activeElement as HTMLElement);
  const nextIndex = event.shiftKey
    ? activeIndex <= 0
      ? focusable.length - 1
      : activeIndex - 1
    : activeIndex < 0 || activeIndex >= focusable.length - 1
      ? 0
      : activeIndex + 1;
  event.preventDefault();
  event.stopPropagation();
  focusable[nextIndex]?.focus();
}

function isTopScope(scope: ActiveFocusScope): boolean {
  return activeFocusScopes.at(-1) === scope;
}

function activeElement(): HTMLElement | null {
  return typeof document !== 'undefined' &&
    document.activeElement instanceof HTMLElement
    ? document.activeElement
    : null;
}

function isolateOutsideScope(
  scope: HTMLElement,
  exceptions: readonly (HTMLElement | null | undefined)[],
): () => void {
  const allowed = exceptions.filter((element): element is HTMLElement =>
    Boolean(element?.isConnected),
  );
  const isolated = new Set<HTMLElement>();
  let branch: HTMLElement | null = scope;
  while (branch) {
    const parent: HTMLElement | null = branch.parentElement;
    if (!parent) break;
    for (const child of parent.children) {
      if (!(child instanceof HTMLElement) || child === branch) continue;
      if (
        allowed.some((element) => child === element || child.contains(element))
      )
        continue;
      isolated.add(child);
    }
    if (parent === document.body) break;
    branch = parent;
  }

  for (const element of isolated) acquireInert(element);
  return () => {
    for (const element of isolated) releaseInert(element);
  };
}

function acquireInert(element: HTMLElement): void {
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

function releaseInert(element: HTMLElement): void {
  const count = inertCounts.get(element) ?? 0;
  if (count > 1) {
    inertCounts.set(element, count - 1);
    return;
  }
  inertCounts.delete(element);
  const state = inertStates.get(element);
  inertStates.delete(element);
  element.inert = state?.property ?? false;
  if (state?.attribute) element.setAttribute('inert', '');
  else element.removeAttribute('inert');
}

function defaultInitialFocus(scope: HTMLElement): HTMLElement | null {
  const configured = scope.querySelector<HTMLElement>('[data-autofocus]');
  if (configured && isAvailable(configured)) return configured;

  const field = scope.querySelector<HTMLElement>(
    'input:not(:disabled):not([type="hidden"]), textarea:not(:disabled), select:not(:disabled)',
  );
  if (field && isAvailable(field)) return field;

  const safeFooterAction = scope.querySelector<HTMLElement>(
    ':scope > footer button:not(:disabled)',
  );
  if (safeFooterAction && isAvailable(safeFooterAction))
    return safeFooterAction;
  return focusableElements(scope)[0] ?? null;
}

function focusableElements(scope: HTMLElement): HTMLElement[] {
  return [...scope.querySelectorAll<HTMLElement>(focusableSelector)].filter(
    isAvailable,
  );
}

function isAvailable(element: HTMLElement): boolean {
  const style =
    typeof window === 'undefined' ? null : window.getComputedStyle(element);
  return (
    !element.matches(':disabled') &&
    style?.display !== 'none' &&
    style?.visibility !== 'hidden' &&
    !element.closest('[hidden], [inert], [aria-hidden="true"]')
  );
}
