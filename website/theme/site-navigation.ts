export function playgroundHrefFromDocsRoute(routePath: string): string {
  const normalized = routePath.split(/[?#]/, 1)[0] ?? '/';
  const segments = normalized.split('/').filter(Boolean);
  const directoryDepth = normalized.endsWith('/')
    ? segments.length
    : Math.max(0, segments.length - 1);
  // Documentation is mounted at `/docs/`, while the Playground remains a
  // sibling of that mount. The first parent traversal leaves the docs tree;
  // the remaining traversals account for the current page depth.
  return `../${'../'.repeat(directoryDepth)}playground/`;
}

export function playgroundAssetHrefFromDocsRoute(
  routePath: string,
  assetPath: string,
): string {
  const normalizedAssetPath = assetPath.replace(/^\/+/, '');
  return `${playgroundHrefFromDocsRoute(routePath)}${normalizedAssetPath}`;
}
