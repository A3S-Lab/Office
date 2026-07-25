import { useLayoutEffect, useMemo, useRef } from 'react';
import {
  createOfficeEditorRuntime,
  type OfficeEditorExtension,
  type OfficeEditorRuntime,
} from './office-editor-extension';

/**
 * Keeps one extension runtime mounted while React supplies the latest
 * controlled editor context. Callers should memoize the extension list.
 */
export function useOfficeEditorRuntime<Context, Commands>(
  context: Context,
  extensions: readonly OfficeEditorExtension<Context, Commands>[],
): OfficeEditorRuntime<Context, Commands> {
  const contextRef = useRef(context);
  contextRef.current = context;
  const runtime = useMemo(
    () =>
      createOfficeEditorRuntime(contextRef.current, extensions, {
        getCurrentContext: () => contextRef.current,
      }),
    [extensions],
  );

  useLayoutEffect(() => {
    runtime.mount();
    return () => runtime.unmount();
  }, [runtime]);

  useLayoutEffect(() => {
    runtime.updateContext(context);
  }, [context, runtime]);

  return runtime;
}
