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
