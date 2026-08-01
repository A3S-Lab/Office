const DOCUMENTATION_ENTRY_PATH = 'docs/index.html';
const GETTING_STARTED_PATH = 'docs/guide/index.html';

export function documentationEntryUrl(baseUri: string): string {
  return new URL(DOCUMENTATION_ENTRY_PATH, baseUri).href;
}

export function legacyDocsPath(hash: string): string | null {
  if (hash === '#guide/components') return 'docs/components/index.html';
  if (hash === '#guide/api') return 'docs/components/document.html';
  if (
    hash === '#guide/automation' ||
    hash === '#guide/cli' ||
    hash === '#guide/skill' ||
    hash === '#cli' ||
    hash === '#skill'
  ) {
    return 'docs/automation/index.html';
  }
  return hash === '#guide' ? GETTING_STARTED_PATH : null;
}
