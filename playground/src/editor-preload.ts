import type { OfficeArtifactKind } from '@a3s-lab/office/core';

const loadEditorWorkspaceModule = () => import('./editor-workspace');

interface PlaygroundEditorPreloadOptions {
  preloadRuntimeAssets?: boolean;
}

export function loadPlaygroundEditorWorkspace() {
  return loadEditorWorkspaceModule();
}

export async function preloadPlaygroundEditor(
  kind: OfficeArtifactKind,
  options: PlaygroundEditorPreloadOptions = {},
): Promise<void> {
  await Promise.all([
    loadEditorWorkspaceModule(),
    import('@a3s-lab/office/react').then(({ preloadOfficeEditor }) =>
      preloadOfficeEditor(kind, options),
    ),
  ]);
  if (typeof performance !== 'undefined') {
    performance.mark(
      `a3s-office.playground.${kind}.${
        options.preloadRuntimeAssets === true ? 'runtime' : 'module'
      }-preload-ready`,
    );
  }
}

export function warmPlaygroundEditor(
  kind: OfficeArtifactKind,
  options?: PlaygroundEditorPreloadOptions,
): void {
  void preloadPlaygroundEditor(
    kind,
    options ?? (kind === 'pdf' ? { preloadRuntimeAssets: true } : undefined),
  ).catch(() => undefined);
}
