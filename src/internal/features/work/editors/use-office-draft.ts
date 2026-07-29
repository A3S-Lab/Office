import {
  type Dispatch,
  type SetStateAction,
  useCallback,
  useRef,
  useState,
} from 'react';

export function useOfficeDraft<Value>(createInitial: () => Value) {
  const [initial] = useState(createInitial);
  const [draft, setDraftState] = useState(initial);
  const [baseline, setBaseline] = useState(initial);
  const draftRef = useRef(draft);
  const baselineRef = useRef(baseline);
  draftRef.current = draft;
  baselineRef.current = baseline;

  const setDraft = useCallback<Dispatch<SetStateAction<Value>>>((action) => {
    setDraftState((current) => {
      const next =
        typeof action === 'function'
          ? (action as (current: Value) => Value)(current)
          : action;
      draftRef.current = next;
      return next;
    });
  }, []);

  const replaceDraft = useCallback((next: Value) => {
    draftRef.current = next;
    baselineRef.current = next;
    setDraftState(next);
    setBaseline(next);
  }, []);

  const syncDraft = useCallback((next: Value, force = false) => {
    const replaceCurrent = force || draftRef.current === baselineRef.current;
    baselineRef.current = next;
    setBaseline(next);
    if (!replaceCurrent) return;
    draftRef.current = next;
    setDraftState(next);
  }, []);

  const cancelDraft = useCallback(() => {
    const next = baselineRef.current;
    draftRef.current = next;
    setDraftState(next);
  }, []);

  return {
    cancelDraft,
    dirty: draft !== baseline,
    draft,
    draftRef,
    replaceDraft,
    setDraft,
    syncDraft,
  };
}
