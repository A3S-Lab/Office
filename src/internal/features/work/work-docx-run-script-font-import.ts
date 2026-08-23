import type {
  ImportedDocxRunFormatting,
  ImportedDocxRunFormattingChange,
  ImportedDocxRunFormattingMarkerState,
} from './work-docx-run-formatting-import';
import type {
  WorkDocumentScriptFontSegment,
  WorkDocumentScriptFontSlot,
} from './work-document-script-fonts';

interface DocxRunFormattingMarkers {
  endMarker: string;
  startMarker: string;
}

export function directDocxRunText(run: Element): string {
  return directDocxRunTextElements(run)
    .map((child) => child.textContent ?? '')
    .join('');
}

export function markSegmentedDocxRunFormatting(
  run: Element,
  state: ImportedDocxRunFormattingMarkerState,
  segments: readonly WorkDocumentScriptFontSegment[],
  formattingFor: (
    text: string,
    slot: WorkDocumentScriptFontSlot,
  ) => ImportedDocxRunFormatting,
  changeFor: (
    text: string,
    slot: WorkDocumentScriptFontSlot,
  ) => ImportedDocxRunFormattingChange | undefined,
  createMarkers: () => DocxRunFormattingMarkers,
): void {
  let runOffset = 0;
  for (const textElement of directDocxRunTextElements(run)) {
    const source = textElement.textContent ?? '';
    const elementFrom = runOffset;
    const elementTo = elementFrom + source.length;
    runOffset = elementTo;
    if (!source) continue;
    const replacement: string[] = [];
    for (const segment of segments) {
      const from = Math.max(elementFrom, segment.from);
      const to = Math.min(elementTo, segment.to);
      if (to <= from) continue;
      const text = source.slice(from - elementFrom, to - elementFrom);
      const formatting = formattingFor(text, segment.slot);
      const change = changeFor(text, segment.slot);
      if (!Object.keys(formatting).length && !change) {
        replacement.push(text);
        continue;
      }
      const { startMarker, endMarker } = createMarkers();
      replacement.push(startMarker, text, endMarker);
      state.markers.runs.push({
        startMarker,
        endMarker,
        formatting,
        ...(change ? { change } : {}),
      });
    }
    textElement.textContent = replacement.join('');
  }
}

function directDocxRunTextElements(run: Element): Element[] {
  return Array.from(run.children).filter(
    (child) => child.localName === 't' || child.localName === 'delText',
  );
}
