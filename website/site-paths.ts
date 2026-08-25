const DEFAULT_SITE_BASE = '/';

export function siteBaseFromEnvironment(
  environment: NodeJS.ProcessEnv = process.env,
): string {
  return normalizeSiteBase(
    environment.A3S_OFFICE_SITE_BASE ??
      environment.A3S_OFFICE_PLAYGROUND_BASE ??
      DEFAULT_SITE_BASE,
  );
}

export function playgroundBaseFromSiteBase(siteBase: string): string {
  return `${normalizeSiteBase(siteBase)}playground/`;
}

export function normalizeSiteBase(value: string): string {
  const withLeadingSlash = value.startsWith('/') ? value : `/${value}`;
  return withLeadingSlash.endsWith('/')
    ? withLeadingSlash
    : `${withLeadingSlash}/`;
}
