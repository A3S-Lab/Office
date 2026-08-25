const DOCUMENTATION_ENTRY_PATH = '../';
const GETTING_STARTED_PATH = '../guide/';
const COLLABORATION_SERVER_PATH = '../components/collaboration-server.html';

export function documentationEntryUrl(baseUri: string): string {
  return new URL(DOCUMENTATION_ENTRY_PATH, baseUri).href;
}

export function collaborationServerDocumentationUrl(baseUri: string): string {
  return new URL(COLLABORATION_SERVER_PATH, baseUri).href;
}

export function legacyDocsPath(hash: string): string | null {
  if (hash === '#guide/components') return '../components/';
  if (hash === '#guide/api') return '../components/document.html';
  if (
    hash === '#guide/automation' ||
    hash === '#guide/cli' ||
    hash === '#guide/skill' ||
    hash === '#cli' ||
    hash === '#skill'
  ) {
    return '../automation/';
  }
  return hash === '#guide' ? GETTING_STARTED_PATH : null;
}
