import {
  type ImportedDocxImageLayoutMarkerState,
  type ImportedDocxImageLayoutMarkers,
  markDocxImageLayouts,
} from './work-docx-image-layout-import';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import type { OoxmlPackage } from './work-ooxml-package';

interface MarkedNoteImagePart {
  document: Document;
  path: string;
}

export interface ImportedDocxNoteImageMarkers {
  markers: ImportedDocxImageLayoutMarkers;
  parts: MarkedNoteImagePart[];
}

export async function markDocxNoteImageLayouts(
  archive: OoxmlPackage,
  state: ImportedDocxImageLayoutMarkerState,
): Promise<ImportedDocxNoteImageMarkers> {
  const markers: ImportedDocxImageLayoutMarkers = { images: [] };
  const parts: MarkedNoteImagePart[] = [];
  for (const config of [
    { path: 'word/footnotes.xml', root: 'footnotes' },
    { path: 'word/endnotes.xml', root: 'endnotes' },
  ] as const) {
    if (!archive.has(config.path)) continue;
    const document = await archive.xml(config.path);
    const root = document.documentElement;
    if (
      root.localName !== config.root ||
      !DOCX_WORDPROCESSING_NAMESPACES.has(root.namespaceURI ?? '')
    ) {
      continue;
    }
    const partMarkers = markDocxImageLayouts(document, state);
    if (!partMarkers.images.length) continue;
    markers.images.push(...partMarkers.images);
    parts.push({ document, path: config.path });
  }
  return { markers, parts };
}
