import {
  type KeyboardEvent,
  useCallback,
  useEffect,
  useRef,
  useState,
} from 'react';
import type { WorkDocumentSectionLayout } from '../work-types';

export interface DocumentNavigationPage {
  backgroundColor?: string;
  physicalPage: number;
  pageNumber: number;
  orientation: WorkDocumentSectionLayout['orientation'];
  previewText: string;
  selectionPosition: number;
}

export function DocumentPageNavigation({
  currentPage,
  pages,
  onSelectPage,
}: {
  currentPage: number;
  pages: readonly DocumentNavigationPage[];
  onSelectPage: (page: DocumentNavigationPage) => void | Promise<void>;
}) {
  const pageRefs = useRef(new Map<number, HTMLButtonElement>());
  const [selectedPage, setSelectedPage] = useState(currentPage);
  const [rovingPage, setRovingPage] = useState(currentPage);

  useEffect(() => {
    setSelectedPage(currentPage);
    setRovingPage(currentPage);
  }, [currentPage]);

  useEffect(() => {
    if (pages.some((page) => page.physicalPage === rovingPage)) return;
    const fallback =
      pages.find((page) => page.physicalPage === currentPage) ?? pages[0];
    if (fallback) setRovingPage(fallback.physicalPage);
  }, [currentPage, pages, rovingPage]);

  useEffect(() => {
    pageRefs.current
      .get(selectedPage)
      ?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }, [selectedPage]);

  const focusPage = useCallback((pageNumber: number) => {
    setRovingPage(pageNumber);
    requestAnimationFrame(() => {
      pageRefs.current.get(pageNumber)?.focus({ preventScroll: true });
    });
  }, []);

  const handleKeyDown = (
    event: KeyboardEvent<HTMLButtonElement>,
    page: DocumentNavigationPage,
  ) => {
    const index = pages.findIndex(
      (candidate) => candidate.physicalPage === page.physicalPage,
    );
    const focusAt = (requestedIndex: number) => {
      const next = pages[requestedIndex];
      if (next) focusPage(next.physicalPage);
    };
    if (event.key === 'ArrowDown' || event.key === 'ArrowRight') {
      event.preventDefault();
      focusAt(Math.min(pages.length - 1, index + 1));
    } else if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') {
      event.preventDefault();
      focusAt(Math.max(0, index - 1));
    } else if (event.key === 'Home') {
      event.preventDefault();
      focusAt(0);
    } else if (event.key === 'End') {
      event.preventDefault();
      focusAt(pages.length - 1);
    }
  };

  return (
    <nav
      className="work-document-task-pane-body work-document-page-navigation"
      aria-label="文档页面"
    >
      {pages.length ? (
        <ol>
          {pages.map((page) => {
            const current = page.physicalPage === selectedPage;
            return (
              <li key={page.physicalPage}>
                <button
                  ref={(element) => {
                    if (element) {
                      pageRefs.current.set(page.physicalPage, element);
                    } else {
                      pageRefs.current.delete(page.physicalPage);
                    }
                  }}
                  type="button"
                  className={current ? 'active' : undefined}
                  aria-current={current ? 'page' : undefined}
                  aria-label={`第 ${page.physicalPage} 页`}
                  data-document-page-thumbnail={page.physicalPage}
                  tabIndex={page.physicalPage === rovingPage ? 0 : -1}
                  onFocus={() => setRovingPage(page.physicalPage)}
                  onKeyDown={(event) => handleKeyDown(event, page)}
                  onClick={() => {
                    setSelectedPage(page.physicalPage);
                    void onSelectPage(page);
                  }}
                >
                  <span
                    className={`work-document-page-thumbnail ${page.orientation}`}
                    aria-hidden="true"
                    style={{ backgroundColor: page.backgroundColor }}
                  >
                    <span>{page.previewText || '空白页'}</span>
                  </span>
                  <span className="work-document-page-thumbnail-label">
                    第 {page.physicalPage} 页
                    {page.pageNumber !== page.physicalPage && (
                      <small>页码 {page.pageNumber}</small>
                    )}
                  </span>
                </button>
              </li>
            );
          })}
        </ol>
      ) : (
        <div className="work-document-outline-empty">正在生成页面预览…</div>
      )}
    </nav>
  );
}
