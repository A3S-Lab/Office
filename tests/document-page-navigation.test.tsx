import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import {
  DOCUMENT_PAGE_WINDOW_LIMIT,
  DocumentPageNavigation,
  calculateDocumentPageRange,
  type DocumentNavigationPage,
} from '../src/internal/features/work/editors/document-page-navigation';

test('keeps short page collections fully mounted', () => {
  const pages = documentPages(8);
  const view = render(
    <DocumentPageNavigation
      currentPage={3}
      pages={pages}
      onSelectPage={() => undefined}
    />,
  );

  const navigation = screen.getByRole('navigation', { name: '文档页面' });
  expect(navigation).toHaveAttribute('data-document-page-windowed', 'false');
  expect(
    view.container.querySelectorAll('[data-document-page-thumbnail]'),
  ).toHaveLength(8);
  expect(
    within(navigation).getByRole('button', { name: '第 3 页' }),
  ).toHaveAttribute('aria-current', 'page');
  expect(
    view.container.querySelector('[data-document-page-spacer]'),
  ).toBeNull();
});

test('mounts a bounded page window with physical scroll spacers', () => {
  const view = render(
    <DocumentPageNavigation
      currentPage={128}
      pages={documentPages(240)}
      onSelectPage={() => undefined}
    />,
  );

  const navigation = screen.getByRole('navigation', { name: '文档页面' });
  const mountedPages = view.container.querySelectorAll(
    '[data-document-page-thumbnail]',
  );
  expect(navigation).toHaveAttribute('data-document-page-count', '240');
  expect(navigation).toHaveAttribute('data-document-page-windowed', 'true');
  expect(mountedPages.length).toBeGreaterThan(1);
  expect(mountedPages.length).toBeLessThanOrEqual(
    DOCUMENT_PAGE_WINDOW_LIMIT + 2,
  );
  expect(
    within(navigation).getByRole('button', { name: '第 128 页' }),
  ).toHaveAttribute('aria-current', 'page');
  expect(
    within(navigation).queryByRole('button', { name: '第 1 页' }),
  ).not.toBeInTheDocument();
  expect(
    within(navigation).queryByRole('button', { name: '第 240 页' }),
  ).not.toBeInTheDocument();

  const before = view.container.querySelector<HTMLElement>(
    '[data-document-page-spacer="before"]',
  );
  const after = view.container.querySelector<HTMLElement>(
    '[data-document-page-spacer="after"]',
  );
  expect(Number.parseFloat(before?.style.height ?? '0')).toBeGreaterThan(0);
  expect(Number.parseFloat(after?.style.height ?? '0')).toBeGreaterThan(0);
});

test('focuses Home and End destinations without waiting for an animation frame', async () => {
  const originalRequestAnimationFrame = globalThis.requestAnimationFrame;
  globalThis.requestAnimationFrame = (() => 1) as typeof requestAnimationFrame;

  try {
    const view = render(
      <DocumentPageNavigation
        currentPage={1}
        pages={documentPages(240)}
        onSelectPage={() => undefined}
      />,
    );
    const navigation = screen.getByRole('navigation', { name: '文档页面' });
    const firstPage = within(navigation).getByRole('button', {
      name: '第 1 页',
    });
    firstPage.focus();
    fireEvent.keyDown(firstPage, { key: 'End' });

    const lastPage = await within(navigation).findByRole('button', {
      name: '第 240 页',
    });
    await waitFor(() => expect(lastPage).toHaveFocus());
    expect(firstPage).toBeInTheDocument();
    expect(firstPage).toHaveAttribute('aria-current', 'page');
    expect(
      view.container.querySelectorAll('[data-document-page-thumbnail]').length,
    ).toBeLessThanOrEqual(DOCUMENT_PAGE_WINDOW_LIMIT + 2);

    fireEvent.keyDown(lastPage, { key: 'Home' });
    await waitFor(() => expect(firstPage).toHaveFocus());
  } finally {
    globalThis.requestAnimationFrame = originalRequestAnimationFrame;
  }
});

test('moves the mounted window on scroll while retaining current-page state', async () => {
  const selectedPages: number[] = [];
  const view = render(
    <DocumentPageNavigation
      currentPage={1}
      pages={documentPages(240)}
      onSelectPage={(page) => selectedPages.push(page.physicalPage)}
    />,
  );
  const navigation = screen.getByRole('navigation', { name: '文档页面' });

  fireEvent.scroll(navigation, { target: { scrollTop: 1_000_000 } });
  await waitFor(() =>
    expect(
      Number(navigation.getAttribute('data-document-page-window-start')),
    ).toBeGreaterThan(0),
  );

  const lastPage = within(navigation).getByRole('button', {
    name: '第 240 页',
  });
  expect(
    within(navigation).getByRole('button', { name: '第 1 页' }),
  ).toHaveAttribute('aria-current', 'page');
  expect(
    view.container.querySelectorAll('[data-document-page-thumbnail]').length,
  ).toBeLessThanOrEqual(DOCUMENT_PAGE_WINDOW_LIMIT + 2);

  fireEvent.click(lastPage);
  expect(selectedPages).toEqual([240]);
  expect(lastPage).toHaveAttribute('aria-current', 'page');
});

test('calculates stable bounded ranges for short and long documents', () => {
  expect(
    calculateDocumentPageRange({
      anchorIndex: 99,
      averageItemHeight: 230,
      pageCount: 300,
      viewportHeight: 720,
    }),
  ).toEqual({ end: 107, start: 95, windowed: true });
  expect(
    calculateDocumentPageRange({
      anchorIndex: 0,
      averageItemHeight: 230,
      pageCount: 8,
      viewportHeight: 720,
    }),
  ).toEqual({ end: 8, start: 0, windowed: false });
});

function documentPages(pageCount: number): DocumentNavigationPage[] {
  return Array.from({ length: pageCount }, (_, index) => ({
    orientation: index % 9 === 8 ? 'landscape' : 'portrait',
    pageNumber: index + 1,
    physicalPage: index + 1,
    previewText: `A3S Office page ${index + 1}`,
    selectionPosition: index * 10 + 1,
  }));
}
