import { expect, test } from '@rstest/core';
import { render } from '@testing-library/react';
import { DocumentPageStack } from '../src/internal/features/work/editors/document-page-stack';

test('renders one measured paper surface for every paginated page', () => {
  const { container, rerender } = render(
    <DocumentPageStack
      pageColor="#fff2cc"
      pageCount={3}
      pageGap={28}
      pageHeight={1123}
    />,
  );

  const stack = container.querySelector('.work-document-page-stack');
  expect(stack).toHaveAttribute('aria-hidden', 'true');
  expect(stack).toHaveAttribute('data-page-count', '3');

  const pages = container.querySelectorAll('.work-document-page-sheet');
  expect(pages).toHaveLength(3);
  expect(pages[0]).toHaveStyle({
    backgroundColor: '#fff2cc',
    height: '1123px',
    top: '0px',
  });
  expect(pages[1]).toHaveStyle({ top: '1151px' });
  expect(pages[2]).toHaveStyle({ top: '2302px' });

  rerender(
    <DocumentPageStack
      pageColor="#ffffff"
      pageCount={2}
      pageGap={28}
      pageHeight={1123}
    />,
  );

  expect(container.querySelectorAll('.work-document-page-sheet')).toHaveLength(
    2,
  );
});

test('centers variable-size pages on cumulative physical offsets', () => {
  const layout = {
    pageSize: 'a4' as const,
    orientation: 'portrait' as const,
    margins: { top: 20, right: 20, bottom: 20, left: 20 },
    columns: { count: 1, spacing: 12, separator: false },
    breakAfter: 'nextPage' as const,
  };
  const page = (width: number, height: number, pageGap: number) => ({
    width,
    height,
    marginTop: 20,
    marginRight: 20,
    marginBottom: 20,
    marginLeft: 20,
    headerHeight: 10,
    footerHeight: 10,
    pageGap,
  });
  const { container } = render(
    <DocumentPageStack
      pageColor="#ffffff"
      pageCount={2}
      pageGap={28}
      pageHeight={400}
      pageWidth={300}
      pages={[
        { layout, page: page(300, 400, 20), sectionPage: 1 },
        { layout, page: page(500, 200, 30), sectionPage: 1 },
      ]}
    />,
  );

  const stack = container.querySelector('.work-document-page-stack');
  expect(stack).toHaveAttribute('data-page-surface-width', '500');
  expect(stack).toHaveAttribute('data-page-surface-height', '620');
  const sheets = container.querySelectorAll('.work-document-page-sheet');
  expect(sheets[0]).toHaveStyle({
    height: '400px',
    left: '100px',
    top: '0px',
    width: '300px',
  });
  expect(sheets[1]).toHaveStyle({
    height: '200px',
    left: '0px',
    top: '420px',
    width: '500px',
  });
});
