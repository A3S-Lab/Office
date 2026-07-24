import { expect, test } from '@rstest/core';
import { render, screen } from '@testing-library/react';
import { CodeBlock } from '../playground/src/code-block';

test('renders integration examples with language-aware syntax tokens', () => {
  const { container } = render(
    <CodeBlock
      code={"import { DocumentEditor } from '@a3s-lab/office/react';"}
      language="tsx"
    />,
  );

  expect(container.querySelector('pre')).toHaveAttribute(
    'data-code-language',
    'tsx',
  );
  expect(container.querySelectorAll('.token').length).toBeGreaterThan(1);
  expect(screen.getByRole('button', { name: '复制代码' })).toBeVisible();
});
