import type { CSSProperties } from 'react';
import {
  DOCUMENT_PAGE_BORDER_EDGES,
  documentPageBordersVisible,
  normalizeDocumentPageBorders,
  resolveDocumentPageBorders,
} from '../work-document-page-borders';
import { isDocumentParagraphArtBorderStyle } from '../work-document-paragraph-borders';
import type { WorkDocumentSectionLayout } from '../work-types';

export function WorkDocumentPageBorder({
  layout,
  sectionPage,
}: {
  layout: WorkDocumentSectionLayout;
  sectionPage: number;
}) {
  const source = normalizeDocumentPageBorders(layout.pageBorders);
  if (!source || !documentPageBordersVisible(source, sectionPage)) return null;
  const resolved = resolveDocumentPageBorders(source, layout.margins);
  if (!resolved) return null;
  const visibleEdges = DOCUMENT_PAGE_BORDER_EDGES.filter(
    (edge) => (resolved.edges[edge]?.width ?? 0) > 0,
  );
  if (!visibleEdges.length) return null;
  const shadow = visibleEdges
    .map((edge) => source.edges[edge])
    .find((border) => border?.shadow);
  const shadowColor = shadow
    ? (resolved.edges[
        visibleEdges.find((edge) => source.edges[edge] === shadow) ?? 'top'
      ]?.color ?? '#000000')
    : null;
  const art = visibleEdges.some((edge) => {
    const border = source.edges[edge];
    return border && isDocumentParagraphArtBorderStyle(border.style);
  });
  const frame = visibleEdges.some((edge) => source.edges[edge]?.frame);
  return (
    <div
      className={`work-document-page-border ${resolved.zOrder}`}
      data-document-page-border-art={String(art)}
      data-document-page-border-display={resolved.display}
      data-document-page-border-frame={String(frame)}
      data-document-page-border-offset-from={resolved.offsetFrom}
      data-document-page-border-z-order={resolved.zOrder}
      aria-hidden="true"
      style={
        {
          top: resolved.insets.top,
          right: resolved.insets.right,
          bottom: resolved.insets.bottom,
          left: resolved.insets.left,
          ...Object.fromEntries(
            DOCUMENT_PAGE_BORDER_EDGES.flatMap((edge) => {
              const border = resolved.edges[edge];
              if (!border || border.width <= 0) return [];
              const property = `border${edge[0]?.toUpperCase()}${edge.slice(1)}`;
              return [
                [
                  property,
                  `${formatPixels(border.width)}px ${border.style} ${border.color}`,
                ],
              ];
            }),
          ),
          ...(shadowColor
            ? { filter: `drop-shadow(2px 2px 0 ${shadowColor})` }
            : {}),
        } as CSSProperties
      }
    />
  );
}

function formatPixels(value: number): string {
  return Number(value.toFixed(3)).toString();
}
