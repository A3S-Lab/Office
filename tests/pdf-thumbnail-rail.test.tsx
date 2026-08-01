import type { PluginRegistry } from '@embedpdf/react-pdf-viewer';
import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useState } from 'react';
import {
  PDF_THUMBNAIL_WINDOW_LIMIT,
  PdfThumbnailRail,
  calculatePdfThumbnailRange,
} from '../src/internal/features/work/editors/pdf-thumbnail-rail';

test('shows page numbers, current-page state, and routes thumbnail selection', async () => {
  const selectedPages: number[] = [];
  const renderedPages: number[] = [];
  const urls = installObjectUrlFixture();
  const registry = createRegistry((pageIndex) => {
    renderedPages.push(pageIndex);
    return new Blob([`page-${pageIndex}`], { type: 'image/png' });
  });

  try {
    const view = render(
      <PdfThumbnailRail
        currentPage={2}
        registry={registry}
        totalPages={3}
        onSelectPage={(page) => selectedPages.push(page)}
      />,
    );

    expect(
      screen.getByRole('navigation', { name: 'PDF 页面缩略图' }),
    ).toBeInTheDocument();
    expect(screen.getByRole('button', { name: '第 2 页' })).toHaveAttribute(
      'aria-current',
      'page',
    );
    fireEvent.click(screen.getByRole('button', { name: '第 3 页' }));
    expect(selectedPages).toEqual([3]);

    await waitFor(() => expect(renderedPages).toEqual([0, 1, 2]));
    expect(
      view.container.querySelectorAll('[data-pdf-thumbnail-state="ready"]'),
    ).toHaveLength(3);

    view.unmount();
    expect(urls.revoked).toEqual(urls.created);
  } finally {
    urls.restore();
  }
});

test('moves focus with continuous thumbnail keyboard navigation', async () => {
  const urls = installObjectUrlFixture();
  const registry = createRegistry(
    (pageIndex) => new Blob([`page-${pageIndex}`], { type: 'image/png' }),
  );
  let view: ReturnType<typeof render> | null = null;

  function ControlledThumbnailRail() {
    const [currentPage, setCurrentPage] = useState(1);
    return (
      <PdfThumbnailRail
        currentPage={currentPage}
        registry={registry}
        totalPages={4}
        onSelectPage={setCurrentPage}
      />
    );
  }

  try {
    view = render(<ControlledThumbnailRail />);
    const firstPage = screen.getByRole('button', { name: '第 1 页' });
    firstPage.focus();

    fireEvent.keyDown(firstPage, { key: 'ArrowDown' });
    const secondPage = screen.getByRole('button', { name: '第 2 页' });
    await waitFor(() => expect(secondPage).toHaveFocus());
    expect(secondPage).toHaveAttribute('aria-current', 'page');

    fireEvent.keyDown(secondPage, { key: 'ArrowDown' });
    const thirdPage = screen.getByRole('button', { name: '第 3 页' });
    await waitFor(() => expect(thirdPage).toHaveFocus());
    expect(thirdPage).toHaveAttribute('aria-current', 'page');

    fireEvent.keyDown(thirdPage, { key: 'End' });
    const lastPage = screen.getByRole('button', { name: '第 4 页' });
    await waitFor(() => expect(lastPage).toHaveFocus());
    expect(lastPage).toHaveAttribute('aria-current', 'page');

    fireEvent.keyDown(lastPage, { key: 'Home' });
    await waitFor(() => expect(firstPage).toHaveFocus());
    expect(firstPage).toHaveAttribute('aria-current', 'page');
  } finally {
    view?.unmount();
    urls.restore();
  }
});

test('focuses a keyboard destination after long-document virtualization mounts it', async () => {
  const urls = installObjectUrlFixture();
  const registry = createRegistry(
    (pageIndex) => new Blob([`page-${pageIndex}`], { type: 'image/png' }),
  );
  let view: ReturnType<typeof render> | null = null;

  function ControlledLongThumbnailRail() {
    const [currentPage, setCurrentPage] = useState(1);
    return (
      <PdfThumbnailRail
        currentPage={currentPage}
        registry={registry}
        totalPages={240}
        onSelectPage={setCurrentPage}
      />
    );
  }

  try {
    view = render(<ControlledLongThumbnailRail />);
    const firstPage = screen.getByRole('button', { name: '第 1 页' });
    firstPage.focus();
    fireEvent.keyDown(firstPage, { key: 'End' });

    await waitFor(() => {
      const lastPage = screen.getByRole('button', { name: '第 240 页' });
      expect(lastPage).toHaveAttribute('aria-current', 'page');
      expect(lastPage).toHaveFocus();
    });
  } finally {
    view?.unmount();
    urls.restore();
  }
});

test('windows long documents around the current page and keeps it mounted', async () => {
  const renderedPages: number[] = [];
  const urls = installObjectUrlFixture();
  let view: ReturnType<typeof render> | null = null;
  const registry = createRegistry((pageIndex) => {
    renderedPages.push(pageIndex);
    return new Blob([`page-${pageIndex}`], { type: 'image/png' });
  });

  try {
    view = render(
      <PdfThumbnailRail
        currentPage={128}
        registry={registry}
        totalPages={240}
        onSelectPage={() => undefined}
      />,
    );
    let thumbnails = view.container.querySelectorAll(
      '[data-pdf-page-thumbnail]',
    );
    expect(thumbnails.length).toBeGreaterThan(1);
    expect(thumbnails.length).toBeLessThanOrEqual(PDF_THUMBNAIL_WINDOW_LIMIT);
    expect(screen.getByRole('button', { name: '第 128 页' })).toHaveAttribute(
      'aria-current',
      'page',
    );

    view.rerender(
      <PdfThumbnailRail
        currentPage={230}
        registry={registry}
        totalPages={240}
        onSelectPage={() => undefined}
      />,
    );
    await waitFor(() =>
      expect(screen.getByRole('button', { name: '第 230 页' })).toHaveAttribute(
        'aria-current',
        'page',
      ),
    );
    thumbnails = view.container.querySelectorAll('[data-pdf-page-thumbnail]');
    expect(thumbnails.length).toBeLessThanOrEqual(PDF_THUMBNAIL_WINDOW_LIMIT);
    await waitFor(() =>
      expect(renderedPages.length).toBeLessThanOrEqual(
        PDF_THUMBNAIL_WINDOW_LIMIT * 2,
      ),
    );
  } finally {
    view?.unmount();
    urls.restore();
  }
});

test('calculates bounded ranges without mounting every page', () => {
  expect(
    calculatePdfThumbnailRange({
      anchorIndex: 99,
      totalPages: 300,
      viewportHeight: 720,
    }),
  ).toEqual({ end: 108, start: 94, windowed: true });
  expect(
    calculatePdfThumbnailRange({
      anchorIndex: 0,
      totalPages: 3,
      viewportHeight: 720,
    }),
  ).toEqual({ end: 3, start: 0, windowed: false });
});

function createRegistry(render: (pageIndex: number) => Blob): PluginRegistry {
  const capabilities = {
    'document-manager': {
      getActiveDocumentId: () => 'document-1',
    },
    thumbnail: {
      forDocument: () => ({
        renderThumb: (pageIndex: number) => ({
          toPromise: () => Promise.resolve(render(pageIndex)),
        }),
      }),
    },
  };
  return {
    pluginsReady: () => Promise.resolve(),
    getPlugin: (id: keyof typeof capabilities) => ({
      provides: () => capabilities[id],
    }),
  } as unknown as PluginRegistry;
}

function installObjectUrlFixture(): {
  created: string[];
  restore: () => void;
  revoked: string[];
} {
  const created: string[] = [];
  const revoked: string[] = [];
  const originalCreate = URL.createObjectURL;
  const originalRevoke = URL.revokeObjectURL;
  Object.defineProperty(URL, 'createObjectURL', {
    configurable: true,
    value: () => {
      const url = `blob:pdf-thumbnail-${created.length + 1}`;
      created.push(url);
      return url;
    },
  });
  Object.defineProperty(URL, 'revokeObjectURL', {
    configurable: true,
    value: (url: string) => revoked.push(url),
  });
  return {
    created,
    revoked,
    restore: () => {
      Object.defineProperty(URL, 'createObjectURL', {
        configurable: true,
        value: originalCreate,
      });
      Object.defineProperty(URL, 'revokeObjectURL', {
        configurable: true,
        value: originalRevoke,
      });
    },
  };
}
