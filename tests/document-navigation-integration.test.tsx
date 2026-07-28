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

test('opens a persistent Word-style navigation pane from the View ribbon', async () => {
  const artifact = createArtifact('project-brief');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  const { unmount } = render(
    <DocumentEditor
      content={artifact.content}
      onChange={() => undefined}
      theme="light"
    />,
  );

  const editor = await screen.findByRole('textbox', { name: '文档正文' });
  fireEvent.click(screen.getByRole('tab', { name: '视图' }));
  fireEvent.click(await screen.findByRole('button', { name: '导航窗格' }));

  const pane = screen.getByRole('complementary', { name: '文档导航' });
  expect(
    within(pane).getByRole('button', { name: '背景与目标' }),
  ).toBeVisible();
  const homeTab = screen.getByRole('tab', { name: '开始' });
  fireEvent.click(homeTab);
  await waitFor(() => expect(homeTab).toHaveAttribute('aria-selected', 'true'));
  expect(pane).toBeVisible();

  const background = within(pane).getByRole('button', {
    name: '背景与目标',
  });
  fireEvent.click(background);
  await waitFor(() =>
    expect(background).toHaveAttribute('aria-current', 'location'),
  );

  fireEvent.keyDown(pane, { key: 'Escape' });
  await waitFor(() => expect(pane).not.toBeInTheDocument());
  await waitFor(() => expect(editor).toHaveFocus());
  unmount();
});

test('closes the Word comments pane with Escape from its ribbon trigger', async () => {
  const artifact = createArtifact('project-brief');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  const { unmount } = render(
    <DocumentEditor
      content={artifact.content}
      onChange={() => undefined}
      theme="light"
    />,
  );

  const editor = await screen.findByRole('textbox', { name: '文档正文' });
  fireEvent.click(screen.getByRole('tab', { name: '审阅' }));
  const commentsToggle = await screen.findByRole('button', {
    name: /^查看批注/,
  });
  fireEvent.click(commentsToggle);

  expect(await screen.findByLabelText('批注审阅')).toBeVisible();
  commentsToggle.focus();
  expect(commentsToggle).toHaveFocus();
  fireEvent.keyDown(commentsToggle, { key: 'Escape' });

  await waitFor(() =>
    expect(screen.queryByLabelText('批注审阅')).not.toBeInTheDocument(),
  );
  await waitFor(() => expect(editor).toHaveFocus());
  unmount();
});
