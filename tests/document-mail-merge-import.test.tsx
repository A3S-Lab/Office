import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { createArtifact } from '../src/core';
import { DocumentEditor } from '../src/react';

test('imports two recipient rows from the insert ribbon and previews the first name', async () => {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  artifact.content.html = [
    '<section data-document-section="true"><p>',
    '<span data-document-field="true" data-field-id="name" data-field-kind="mergeField" data-field-instruction="MERGEFIELD 姓名" data-field-target-name="姓名" data-field-display="«姓名»">«姓名»</span>',
    '</p></section>',
  ].join('');

  render(
    <DocumentEditor
      content={artifact.content}
      onChange={() => undefined}
      preview={false}
      theme="light"
    />,
  );
  await screen.findByRole('textbox', { name: '文档正文' });
  fireEvent.click(screen.getByRole('tab', { name: '插入' }));
  const importRecipients = await screen.findByRole('button', {
    name: '导入收件人',
  });
  expect(screen.queryByRole('button', { name: '筛选收件人' })).toBeNull();
  fireEvent.click(importRecipients);

  const dialog = await screen.findByRole('dialog', { name: '导入收件人' });
  fireEvent.change(within(dialog).getByRole('textbox', { name: '收件人表' }), {
    target: { value: '姓名\n甲\n乙' },
  });
  fireEvent.click(within(dialog).getByRole('button', { name: '导入并预览' }));

  await waitFor(() => {
    expect(screen.getByRole('textbox', { name: '文档正文' })).toHaveTextContent(
      '甲',
    );
  });
  expect(
    screen.getByRole('textbox', { name: '文档正文' }),
  ).not.toHaveTextContent('乙');
  expect(screen.getByRole('button', { name: '筛选收件人' })).toBeEnabled();
});
