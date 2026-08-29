/**
 * The global Office navigation is intentionally small and ordered.  Keep this
 * contract framework-neutral so the Rspress shell and the standalone
 * Playground cannot drift into different labels, destinations, or ordering.
 */
export const OFFICE_SITE_NAVIGATION_KEYS = [
  'home',
  'docs',
  'playground',
  'collaboration',
] as const;

export type OfficeSiteNavigationKey =
  (typeof OFFICE_SITE_NAVIGATION_KEYS)[number];

export type OfficeSiteNavigationLanguage = 'zh' | 'en';

export interface OfficeSiteNavigationItem {
  key: OfficeSiteNavigationKey;
  label: string;
  route: string;
  activeMatch: string;
}

/**
 * Return the one source of truth for the product-level navigation.
 *
 * `playgroundRoute` differs between the product Rspress build (which emits
 * `index.html`) and the docs/standalone build (which serves a directory
 * entry), so callers provide that one deployment detail while every other
 * part of the contract remains shared.
 */
export function officeSiteNavigationItems(
  language: OfficeSiteNavigationLanguage,
  playgroundRoute = '/playground/index.html',
): OfficeSiteNavigationItem[] {
  const english = language === 'en';
  return [
    {
      key: 'home',
      label: english ? 'Product home' : '产品首页',
      route: english ? '/en/index.html' : '/index.html',
      activeMatch: '^/$|^/en/?$',
    },
    {
      key: 'docs',
      label: english ? 'Docs' : '文档',
      route: english ? '/docs/en/index.html' : '/docs/index.html',
      activeMatch: '^/docs(?:/|$)',
    },
    {
      key: 'playground',
      label: 'Playground',
      route: playgroundRoute,
      activeMatch: '^/playground(?:/|$)',
    },
    {
      key: 'collaboration',
      label: english ? 'Collaboration' : '协作',
      route: english
        ? '/docs/en/components/collaboration.html'
        : '/docs/components/collaboration.html',
      activeMatch: '^/docs/(?:[^/]+/)?components/collaboration(?:\\.html)?',
    },
  ];
}

/**
 * Resolve a global navigation route from the standalone Playground mount.
 * `baseUri` is normally `/Office/playground/` (or `/playground/` locally),
 * so traversing one level first preserves a GitHub Pages project prefix.
 */
export function playgroundSiteNavigationHref(
  baseUri: string,
  route: string,
): string {
  const productRoot = new URL('../', baseUri);
  return new URL(route.replace(/^\/+/, ''), productRoot).href;
}

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

function normalizePath(pathname: string): string {
  const path = pathname.split(/[?#]/, 1)[0] ?? '/';
  if (!path || path === '/') return '/';
  return `/${path.replace(/^\/+|\/+$/g, '')}/`.replace('//', '/');
}

function normalizeBase(base: string): string {
  const normalized = normalizePath(base);
  return normalized === '/' ? '/' : normalized;
}

/**
 * Return the deployment root that contains the product, documentation, and
 * Playground mounts. The documentation Rspress build uses `/docs/` as its
 * own base, so its product root is the parent of that mount.
 */
export function productSiteBase(siteBase: string): string {
  const normalized = normalizeBase(siteBase);
  return normalized.endsWith('/docs/')
    ? normalized.slice(0, -'docs/'.length) || '/'
    : normalized;
}

function joinPath(base: string, route: string): string {
  const normalizedBase = normalizeBase(base);
  const normalizedRoute = `/${route.replace(/^\/+/, '')}`;
  if (normalizedBase === '/') return normalizedRoute;
  return `${normalizedBase.replace(/\/$/, '')}${normalizedRoute}`;
}

/**
 * Resolve the pathname exposed by React Router to a full deployment path.
 * During Rspress SSR the pathname can be base-relative, while the browser
 * runtime normally exposes the base-prefixed pathname. Supporting both keeps
 * server-rendered links and hydrated links identical.
 */
export function sitePathFromRoute(pathname: string, siteBase: string): string {
  const path = pathname.split(/[?#]/, 1)[0] || '/';
  const normalizedBase = normalizeBase(siteBase);
  const rootBase = productSiteBase(normalizedBase);

  if (normalizedBase === '/') return path.startsWith('/') ? path : `/${path}`;
  if (
    path === normalizedBase.replace(/\/$/, '') ||
    path.startsWith(normalizedBase)
  ) {
    return path;
  }
  if (
    rootBase !== '/' &&
    (path === rootBase.replace(/\/$/, '') || path.startsWith(rootBase))
  ) {
    return path;
  }
  // A few Rspress versions expose a docs route without the deployment
  // prefix. Reattach the outer root before falling back to the docs base.
  if (path === '/docs' || path.startsWith('/docs/')) {
    return joinPath(rootBase, path);
  }
  return joinPath(normalizedBase, path);
}

function relativePath(fromPath: string, targetPath: string): string {
  const from = fromPath.split('/').filter(Boolean);
  if (!fromPath.endsWith('/')) from.pop();
  const target = targetPath.split('/').filter(Boolean);
  const keepTrailingSlash = targetPath.endsWith('/');
  let common = 0;
  while (
    common < from.length &&
    common < target.length &&
    from[common] === target[common]
  ) {
    common += 1;
  }
  const up = '../'.repeat(from.length - common);
  const down = target.slice(common).join('/');
  const result = `${up}${down}`;
  if (!result) return './';
  return keepTrailingSlash && !result.endsWith('/') ? `${result}/` : result;
}

/**
 * Build a link between the product, docs, and Playground mounts. Links inside
 * the current Rspress base stay absolute and base-aware; links to a sibling
 * mount become relative so GitHub Pages prefixes such as `/Office/` survive.
 */
export function siteNavigationHref(
  pathname: string,
  siteBase: string,
  targetRoute: string,
): string {
  const current = sitePathFromRoute(pathname, siteBase);
  const rootBase = productSiteBase(siteBase);
  const target = joinPath(rootBase, targetRoute);
  const normalizedCurrentBase = normalizeBase(siteBase);
  if (
    normalizedCurrentBase === '/' ||
    target.startsWith(normalizedCurrentBase)
  ) {
    return target;
  }
  return relativePath(current, target);
}

/**
 * Resolve the active global route from the deployment-aware pathname.  Rspress
 * exposes a base-relative pathname during SSR on some versions, so deriving
 * the key from the site base keeps the first paint and the hydrated header in
 * agreement instead of briefly highlighting the product home on every docs
 * page.
 */
export function officeSiteNavigationActiveKey(
  pathname: string,
  siteBase: string,
): OfficeSiteNavigationKey | null {
  const current = sitePathFromRoute(pathname, siteBase)
    .replace(/\/index\.html?$/, '/')
    .replace(/\/{2,}/g, '/');
  const root = productSiteBase(siteBase);
  const relative =
    root === '/'
      ? current.replace(/^\/+/, '')
      : current.startsWith(root)
        ? current.slice(root.length)
        : current.replace(/^\/+/, '');
  const normalized = relative.replace(/^\/+|\/+$/g, '');

  if (!normalized || normalized === 'en') return 'home';
  if (normalized === 'playground' || normalized.startsWith('playground/')) {
    return 'playground';
  }
  if (
    /^docs\/(?:[^/]+\/)*components\/collaboration(?:\.html)?$/.test(normalized)
  ) {
    return 'collaboration';
  }
  if (normalized === 'docs' || normalized.startsWith('docs/')) return 'docs';
  return null;
}

export function isProductHomeRoute(
  pathname: string,
  siteBase: string,
): boolean {
  const current = sitePathFromRoute(pathname, siteBase).replace(
    /\/index\.html?$/,
    '/',
  );
  const root = productSiteBase(siteBase);
  return current === root || current === `${root}en/`;
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
    // Product routes use `/en/` while the default locale uses `/`. Treat the
    // locale prefix as presentation detail so the same destination cannot be
    // rendered twice in the desktop or mobile navigation.
    const withoutLocale = withoutIndex.replace(/^\/(?:en|zh)(?=\/|$)/, '');
    const canonical = withoutLocale || '/';
    if (canonical === '/') return '/';
    return canonical.replace(/\/+$/, '');
  } catch {
    return null;
  }
}
