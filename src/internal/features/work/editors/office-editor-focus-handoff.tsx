import { createContext, type ReactNode, useContext, useRef } from 'react';

type OfficeEditorFocusOrigin = Element | null;

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

function activeFocusOrigin(): OfficeEditorFocusOrigin {
  return typeof document === 'undefined' ? null : document.activeElement;
}
