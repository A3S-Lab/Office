const DOCUMENTATION_ENTRY_PATH = 'docs/';
const GETTING_STARTED_PATH = 'docs/guide/';

export function documentationEntryUrl(baseUri: string): string {
  return new URL(DOCUMENTATION_ENTRY_PATH, baseUri).href;
}

export function legacyDocsPath(hash: string): string | null {
  if (hash === '#guide/components') return 'docs/components/';
  if (hash === '#guide/api') return 'docs/components/document.html';
  if (
    hash === '#guide/automation' ||
    hash === '#guide/cli' ||
    hash === '#guide/skill' ||
    hash === '#cli' ||
    hash === '#skill'
  ) {
    return 'docs/automation/';
  }
  return hash === '#guide' ? GETTING_STARTED_PATH : null;
}
