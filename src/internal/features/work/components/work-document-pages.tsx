import { type CSSProperties, useLayoutEffect, useRef } from 'react';
import { millimetersToPixels } from '../work-document-layout';
import {
  documentColumnGridTemplate,
  documentUnequalColumnGroups,
  normalizeDocumentColumns,
} from '../work-document-columns';
import { resolveDocumentPageChrome } from '../work-document-page-chrome';
import { documentPageColor } from '../work-document-page-color';
import {
  documentPageDescriptors,
  type WorkDocumentPageDescriptor,
  type WorkDocumentPageSegment,
} from '../work-document-pages';
import type {
  WorkDocumentNote,
  WorkDocumentNoteKind,
} from '../work-document-notes';
import { layoutDocumentTabs } from '../work-document-tab-node';
import type { WorkDocumentContent } from '../work-types';
import { WorkDocumentPageBorder } from './work-document-page-border';

export function WorkDocumentPdfPages({
  content,
}: {
  content: WorkDocumentContent;
}) {
  return documentPageDescriptors(content).map((page) => (
    <DocumentPdfPage
      key={page.key}
      page={page}
      pageColor={documentPageColor(content.pageColor)}
    />
  ));
}

function DocumentPdfPage({
  page,
  pageColor,
}: {
  page: WorkDocumentPageDescriptor;
  pageColor: string;
}) {
  const pageRef = useRef<HTMLElement>(null);
  const layout = page.layout;
  const pageChrome = resolveDocumentPageChrome(
    layout,
    page.sectionPage,
    page.physicalPage,
  );
  const pageClass = `work-pdf-export-page document ${layout.pageSize} ${layout.orientation}`;
  const marginPixels = {
    top: millimetersToPixels(layout.margins.top),
    right: millimetersToPixels(layout.margins.right),
    bottom: millimetersToPixels(layout.margins.bottom),
    left: millimetersToPixels(layout.margins.left),
  };
  useLayoutEffect(() => {
    const element = pageRef.current;
    if (!element) return;
    let frame: number | null = null;
    const schedule = () => {
      if (frame !== null) cancelAnimationFrame(frame);
      frame = requestAnimationFrame(() => {
        frame = null;
        if (element.isConnected) layoutDocumentTabs(element);
      });
    };
    layoutDocumentTabs(element);
    const observer =
      typeof ResizeObserver === 'undefined'
        ? null
        : new ResizeObserver(schedule);
    observer?.observe(element);
    void document.fonts?.ready.then(schedule);
    return () => {
      if (frame !== null) cancelAnimationFrame(frame);
      observer?.disconnect();
    };
  }, [page]);
  return (
    <section
      ref={pageRef}
      className={pageClass}
      data-work-pdf-page=""
      data-pdf-orientation={layout.orientation}
      data-pdf-page-size={layout.pageSize}
      data-document-physical-page={page.physicalPage}
      data-document-page-number={page.pageNumber}
      data-document-blank-page={String(page.blank)}
      data-document-page-chrome={pageChrome.variant}
      data-document-comment-appearance="plain"
      aria-label={`文字打印预览第 ${page.physicalPage} 页`}
      style={
        {
          backgroundColor: pageColor,
          padding: `${marginPixels.top}px ${marginPixels.right}px ${marginPixels.bottom}px ${marginPixels.left}px`,
          '--work-document-page-margin-top': `${marginPixels.top}px`,
          '--work-document-page-margin-right': `${marginPixels.right}px`,
          '--work-document-page-margin-bottom': `${marginPixels.bottom}px`,
          '--work-document-page-margin-left': `${marginPixels.left}px`,
        } as CSSProperties
      }
    >
      <WorkDocumentPageBorder layout={layout} sectionPage={page.sectionPage} />
      {pageChrome.headerHtml && (
        <header>
          <div
            className="work-document-page-chrome-html"
            dangerouslySetInnerHTML={{ __html: pageChrome.headerHtml }}
          />
        </header>
      )}
      <div className="work-document-print-body">
        {page.blank ? (
          <span className="work-document-blank-page-label">
            此页按分节设置留空
          </span>
        ) : (
          page.segments.map((segment, index) => (
            <DocumentPageSegmentContent
              key={`${segment.sectionId}-${index}`}
              segment={segment}
            />
          ))
        )}
      </div>
      {page.footnotes.length > 0 && (
        <DocumentPageNotes kind="footnote" notes={page.footnotes} />
      )}
      {page.endnotes.length > 0 && (
        <DocumentPageNotes kind="endnote" notes={page.endnotes} />
      )}
      {(pageChrome.footerHtml || pageChrome.showPageNumber) && (
        <footer>
          {pageChrome.footerHtml ? (
            <div
              className="work-document-page-chrome-html"
              dangerouslySetInnerHTML={{ __html: pageChrome.footerHtml }}
            />
          ) : (
            <span />
          )}
          {pageChrome.showPageNumber && <span>{page.pageNumber}</span>}
        </footer>
      )}
    </section>
  );
}

function DocumentPageSegmentContent({
  segment,
}: {
  segment: WorkDocumentPageSegment;
}) {
  const columns = normalizeDocumentColumns(segment.columns);
  if (!columns.custom) {
    return (
      <article
        data-document-print-section={segment.sectionId}
        data-document-column-count={columns.count}
        data-document-column-layout="equal"
        style={{
          columnCount: columns.count,
          columnGap: `${columns.spacing}mm`,
          columnRule: columns.separator ? '1px solid #cbd0d9' : undefined,
        }}
        dangerouslySetInnerHTML={{ __html: segment.html }}
      />
    );
  }
  const groups = documentUnequalColumnGroups(segment.html, columns);
  return (
    <article
      className="work-document-custom-column-layout"
      data-document-print-section={segment.sectionId}
      data-document-column-count={columns.count}
      data-document-column-layout="custom"
      style={{ gridTemplateColumns: documentColumnGridTemplate(columns) }}
    >
      {groups.map((html, index) => (
        <div
          key={`${segment.sectionId}-column-${index + 1}`}
          data-document-column-index={index + 1}
          style={{
            gridColumn: index * 2 + 1,
            borderRight:
              columns.separator && index < groups.length - 1
                ? '1px solid #cbd0d9'
                : undefined,
          }}
          dangerouslySetInnerHTML={{ __html: html }}
        />
      ))}
    </article>
  );
}

function DocumentPageNotes({
  kind,
  notes,
}: {
  kind: WorkDocumentNoteKind;
  notes: WorkDocumentNote[];
}) {
  return (
    <section
      className={`work-document-page-notes ${kind}`}
      data-document-page-note-kind={kind}
      aria-label={kind === 'footnote' ? '脚注' : '尾注'}
    >
      {kind === 'endnote' && <h3>尾注</h3>}
      <ol>
        {notes.map((note) => (
          <li
            key={`${note.kind}-${note.id}`}
            value={note.number}
            data-document-note-id={note.id}
          >
            <div dangerouslySetInnerHTML={{ __html: note.html }} />
          </li>
        ))}
      </ol>
    </section>
  );
}
