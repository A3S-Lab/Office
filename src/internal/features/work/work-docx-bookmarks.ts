import JSZip from 'jszip';
import {
  normalizeDocumentBookmarkName,
  normalizeDocumentBookmarkNativeId,
  type WorkDocumentBookmarkBoundaryKind,
} from './work-document-bookmarks';
import { descendants, parseXml } from './work-ooxml-package';

interface DocxBookmarkBoundaryPatch {
  marker: string;
  id: string;
  name: string;
  nativeId: number;
  kind: WorkDocumentBookmarkBoundaryKind;
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export class DocxBookmarkPatchCollector {
  readonly patches: DocxBookmarkBoundaryPatch[] = [];
  private nextMarker = 1;

  constructor(private readonly sourceHtml: string) {}

  register(element: HTMLElement): string | null {
    const id = element.dataset.bookmarkId?.trim() ?? '';
    const name = normalizeDocumentBookmarkName(element.dataset.bookmarkName);
    const nativeId = normalizeDocumentBookmarkNativeId(
      element.dataset.officeBookmarkId,
    );
    const kind = element.dataset.bookmarkKind === 'end' ? 'end' : 'start';
    if (!id || !name || nativeId === null) return null;
    let marker = '';
    do {
      marker = `__A3S_DOCX_BOOKMARK_${this.nextMarker}__`;
      this.nextMarker += 1;
    } while (this.sourceHtml.includes(marker));
    this.patches.push({ marker, id, name, nativeId, kind });
    return marker;
  }
}

export async function patchDocxBookmarks(
  buffer: ArrayBuffer,
  patches: readonly DocxBookmarkBoundaryPatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  const validPatches = completeBookmarkPatches(patches);
  if (!validPatches.size) return buffer;
  const archive = await JSZip.loadAsync(buffer);
  const entries = Object.values(archive.files).filter(
    (entry) =>
      !entry.dir &&
      /^word\/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/.test(
        entry.name,
      ),
  );
  const applied = new Set<string>();
  for (const entry of entries) {
    const document = parseXml(await entry.async('text'), entry.name);
    let changed = false;
    for (const text of descendants(document, 't')) {
      const marker = text.textContent ?? '';
      const patch = validPatches.get(marker);
      if (!patch) continue;
      const run = closestAncestor(text, 'r');
      if (!run?.parentNode) continue;
      const boundary = document.createElementNS(
        WORD_NAMESPACE,
        `w:bookmark${patch.kind === 'start' ? 'Start' : 'End'}`,
      );
      boundary.setAttributeNS(WORD_NAMESPACE, 'w:id', String(patch.nativeId));
      if (patch.kind === 'start') {
        boundary.setAttributeNS(WORD_NAMESPACE, 'w:name', patch.name);
      }
      run.parentNode.replaceChild(boundary, run);
      applied.add(marker);
      changed = true;
    }
    if (changed) {
      archive.file(entry.name, new XMLSerializer().serializeToString(document));
    }
  }
  const missing = Array.from(validPatches.keys()).filter(
    (marker) => !applied.has(marker),
  );
  if (missing.length) {
    throw new Error(
      `DOCX bookmark markers were not emitted: ${missing.join(', ')}.`,
    );
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function completeBookmarkPatches(
  patches: readonly DocxBookmarkBoundaryPatch[],
): Map<string, DocxBookmarkBoundaryPatch> {
  const byId = new Map<string, DocxBookmarkBoundaryPatch[]>();
  for (const patch of patches) {
    const matches = byId.get(patch.id) ?? [];
    matches.push(patch);
    byId.set(patch.id, matches);
  }
  return new Map(
    Array.from(byId.values()).flatMap((matches) => {
      const starts = matches.filter((patch) => patch.kind === 'start');
      const ends = matches.filter((patch) => patch.kind === 'end');
      if (starts.length !== 1 || ends.length !== 1) return [];
      const start = starts[0];
      const end = ends[0];
      if (start.name !== end.name || start.nativeId !== end.nativeId) {
        return [];
      }
      return [[start.marker, start] as const, [end.marker, end] as const];
    }),
  );
}

function closestAncestor(element: Element, localName: string): Element | null {
  let current: Element | null = element;
  while (current) {
    if (current.localName === localName) return current;
    current = current.parentElement;
  }
  return null;
}
