import {
  createContext,
  type ReactNode,
  useContext,
  useEffect,
  useRef,
} from 'react';

type OfficeEditorFocusOrigin = Element | null;

const OFFICE_EDITOR_INITIAL_FOCUS_MAX_FRAMES = 120;

const OfficeEditorFocusOriginContext = createContext<
  OfficeEditorFocusOrigin | undefined
>(undefined);

/**
 * Captures the focused opener before a lazy editor suspends so the mounted
 * editing surface can claim that focus without overriding later navigation.
 */
export function OfficeEditorFocusHandoff({
  children,
}: {
  children: ReactNode;
}) {
  const originRef = useRef<OfficeEditorFocusOrigin | undefined>(undefined);
  if (originRef.current === undefined) {
    originRef.current = activeFocusOrigin();
  }
  return (
    <OfficeEditorFocusOriginContext.Provider value={originRef.current}>
      {children}
    </OfficeEditorFocusOriginContext.Provider>
  );
}

/** Returns the pre-suspense opener, or captures the local mount origin. */
export function useOfficeEditorFocusOrigin(): OfficeEditorFocusOrigin {
  const handedOffOrigin = useContext(OfficeEditorFocusOriginContext);
  const localOriginRef = useRef<OfficeEditorFocusOrigin | undefined>(undefined);
  if (localOriginRef.current === undefined) {
    localOriginRef.current = activeFocusOrigin();
  }
  return handedOffOrigin === undefined
    ? localOriginRef.current
    : handedOffOrigin;
}

interface OfficeEditorInitialFocusOptions {
  enabled: boolean;
  getTarget: () => HTMLElement | null;
  isTargetReady?: (target: HTMLElement) => boolean;
}

/**
 * Claims first-open focus only after the editing target is ready, then keeps
 * it stable across the brief DOM replacement window used by rich editors.
 */
export function useOfficeEditorInitialFocus({
  enabled,
  getTarget,
  isTargetReady,
}: OfficeEditorInitialFocusOptions): void {
  const focusOrigin = useOfficeEditorFocusOrigin();
  const getTargetRef = useRef(getTarget);
  const isTargetReadyRef = useRef(isTargetReady);
  getTargetRef.current = getTarget;
  isTargetReadyRef.current = isTargetReady;

  useEffect(() => {
    if (!enabled) return;
    return stabilizeOfficeEditorInitialFocus(
      () => getTargetRef.current(),
      focusOrigin,
      (target) => isTargetReadyRef.current?.(target) ?? true,
    );
  }, [enabled, focusOrigin]);
}

/**
 * Returns a cleanup that cancels the bounded first-open focus handoff.
 * Deliberate pointer navigation, Tab navigation, or unrelated focus wins.
 */
export function stabilizeOfficeEditorInitialFocus(
  getTarget: () => HTMLElement | null,
  focusOrigin: Element | null = activeFocusOrigin(),
  isTargetReady: (target: HTMLElement) => boolean = () => true,
): () => void {
  if (
    typeof document === 'undefined' ||
    typeof requestAnimationFrame !== 'function' ||
    typeof cancelAnimationFrame !== 'function'
  ) {
    return () => undefined;
  }

  let frame = 0;
  let remainingFrames = OFFICE_EDITOR_INITIAL_FOCUS_MAX_FRAMES;
  let lastTarget: HTMLElement | null = null;
  let stopped = false;

  const stop = () => {
    if (stopped) return;
    stopped = true;
    if (frame) cancelAnimationFrame(frame);
    document.removeEventListener('pointerdown', handleOutsideInteraction, true);
    document.removeEventListener('click', handleOutsideInteraction, true);
    document.removeEventListener('keydown', handleKeyDown, true);
  };
  const handleOutsideInteraction = (event: Event) => {
    const target = getTarget();
    if (
      event.target instanceof Node &&
      target?.isConnected &&
      target.contains(event.target)
    ) {
      return;
    }
    stop();
  };
  const handleKeyDown = (event: KeyboardEvent) => {
    if (event.key === 'Tab') stop();
  };
  const canClaimFocus = (
    activeElement: Element | null,
    target: HTMLElement | null,
  ) =>
    activeElement === focusOrigin ||
    activeElement === document.body ||
    activeElement === document.documentElement ||
    !activeElement?.isConnected ||
    Boolean(target?.contains(activeElement)) ||
    Boolean(lastTarget?.contains(activeElement));

  const stabilize = () => {
    frame = 0;
    if (stopped || remainingFrames <= 0) {
      stop();
      return;
    }
    remainingFrames -= 1;

    const target = getTarget();
    const activeElement = document.activeElement;
    if (!canClaimFocus(activeElement, target)) {
      stop();
      return;
    }

    if (target?.isConnected && isTargetReady(target)) {
      const targetAlreadyFocused = target.contains(activeElement);
      if (!targetAlreadyFocused) target.focus({ preventScroll: true });

      if (target.contains(document.activeElement)) {
        lastTarget = target;
      }
    }

    frame = requestAnimationFrame(stabilize);
  };

  document.addEventListener('pointerdown', handleOutsideInteraction, true);
  // Synthetic clicks (tests, assistive technology, or agent automation) do not
  // necessarily emit pointerdown. They still represent deliberate navigation.
  document.addEventListener('click', handleOutsideInteraction, true);
  document.addEventListener('keydown', handleKeyDown, true);
  frame = requestAnimationFrame(stabilize);
  return stop;
}

function activeFocusOrigin(): OfficeEditorFocusOrigin {
  return typeof document === 'undefined' ? null : document.activeElement;
}
