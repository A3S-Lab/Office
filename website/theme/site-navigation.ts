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

/**
 * Resolve a navigation href to a canonical site path for duplicate detection.
 * Rspress may rewrite the same destination as `/index.html`, `/`, or a
 * deployment-relative parent traversal depending on the active route.
 */
export function normalizeNavigationPath(
  pathname: string,
  href: string,
): string | null {
  const origin = 'https://a3s-office.invalid';
  const currentPath = pathname.split(/[?#]/, 1)[0] || '/';

  try {
    const resolved = new URL(href, `${origin}${currentPath}`);
    if (resolved.origin !== origin) return null;
    const withoutIndex = resolved.pathname.replace(/\/index\.html?$/, '/');
    if (withoutIndex === '/') return '/';
    return withoutIndex.replace(/\/+$/, '');
  } catch {
    return null;
  }
}
