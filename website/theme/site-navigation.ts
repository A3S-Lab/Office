export function playgroundHrefFromDocsRoute(routePath: string): string {
  const normalized = routePath.split(/[?#]/, 1)[0] ?? '/';
  const segments = normalized.split('/').filter(Boolean);
  const directoryDepth = normalized.endsWith('/')
    ? segments.length
    : Math.max(0, segments.length - 1);
  return '../'.repeat(directoryDepth + 1);
}

export function playgroundAssetHrefFromDocsRoute(
  routePath: string,
  assetPath: string,
): string {
  const normalizedAssetPath = assetPath.replace(/^\/+/, '');
  return `${playgroundHrefFromDocsRoute(routePath)}${normalizedAssetPath}`;
}
