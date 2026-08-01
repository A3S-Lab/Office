import { Editor } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { DocumentCitationsPanel } from '../src/internal/features/work/editors/document-citations-panel';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import type {
  WorkDocumentCitationSource,
  WorkDocumentContent,
} from '../src/internal/features/work/work-types';

const sources: WorkDocumentCitationSource[] = [
  {
    id: 'source-1',
    tag: 'alpha2026',
    sourceType: 'Book',
    title: 'Alpha Architecture',
    year: '2026',
  },
  {
    id: 'source-2',
    tag: 'beta2025',
    sourceType: 'Report',
    title: 'Beta Systems',
    year: '2025',
  },
];

test('protects an edited citation source before changing the selection', async () => {
  const editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>References</p>',
  });
  const dirtyStates: boolean[] = [];
  const content: WorkDocumentContent = {
    type: 'document',
    html: editor.getHTML(),
    bibliography: { style: 'apa', sources },
  };

  try {
    render(
      <DocumentCitationsPanel
        editor={editor}
        content={content}
        onClose={() => undefined}
        onDirtyChange={(dirty) => dirtyStates.push(dirty)}
      />,
    );

    const title = screen.getByRole('textbox', { name: '文献标题' });
    title.focus();
    fireEvent.change(title, { target: { value: 'Edited architecture' } });
    await waitFor(() => expect(dirtyStates.at(-1)).toBe(true));
    expect(screen.getByRole('button', { name: '保存文献' })).toBeEnabled();
    expect(screen.getByRole('button', { name: '插入引文' })).toBeDisabled();

    const createButton = screen.getByRole('button', { name: '新建' });
    createButton.focus();
    fireEvent.click(createButton);
    expect(
      screen.getByRole('dialog', { name: '放弃未保存的更改？' }),
    ).toBeVisible();
    fireEvent.click(screen.getByRole('button', { name: '取消' }));
    await waitFor(() => expect(title).toHaveFocus());
    expect(title).toHaveValue('Edited architecture');

    const betaButton = screen.getByRole('button', { name: /Beta Systems/ });
    betaButton.focus();
    fireEvent.click(betaButton);
    fireEvent.click(screen.getByRole('button', { name: '放弃更改' }));
    await waitFor(() => expect(title).toHaveValue('Beta Systems'));
    expect(screen.getByRole('button', { name: '保存文献' })).toBeDisabled();
    expect(screen.getByRole('button', { name: '插入引文' })).toBeEnabled();
  } finally {
    editor.destroy();
  }
});

test('uses progressive disclosure and focuses a new citation source', async () => {
  const editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>References</p>',
  });
  const content: WorkDocumentContent = {
    type: 'document',
    html: editor.getHTML(),
    bibliography: { style: 'apa', sources },
  };

  try {
    render(
      <DocumentCitationsPanel
        editor={editor}
        content={content}
        onClose={() => undefined}
      />,
    );

    expect(screen.getByText('更多出版信息')).toBeInTheDocument();
    const publisher = screen.getByRole('textbox', {
      name: '出版者',
      hidden: true,
    });
    expect(publisher).not.toBeVisible();

    fireEvent.click(screen.getByText('更多出版信息'));
    expect(publisher).toBeVisible();

    fireEvent.click(screen.getByRole('button', { name: '新建' }));
    await waitFor(() =>
      expect(screen.getByRole('textbox', { name: '文献简称' })).toHaveFocus(),
    );
    expect(screen.getByRole('form', { name: '新建文献' })).toBeInTheDocument();
    expect(
      screen.queryByRole('button', { name: '删除文献' }),
    ).not.toBeInTheDocument();
  } finally {
    editor.destroy();
  }
});

test('moves citation validation to the invalid required field', async () => {
  const editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>References</p>',
  });
  const content: WorkDocumentContent = {
    type: 'document',
    html: editor.getHTML(),
    bibliography: { style: 'apa', sources: [] },
  };

  try {
    render(
      <DocumentCitationsPanel
        editor={editor}
        content={content}
        onClose={() => undefined}
      />,
    );

    const tag = screen.getByRole('textbox', { name: '文献简称' });
    const title = screen.getByRole('textbox', { name: '文献标题' });
    const save = screen.getByRole('button', { name: '保存文献' });

    fireEvent.change(tag, { target: { value: 'source2026' } });
    save.focus();
    fireEvent.click(save);

    const titleError = await screen.findByRole('alert');
    await waitFor(() => expect(title).toHaveFocus());
    expect(title).toHaveAttribute('aria-invalid', 'true');
    expect(title).toHaveAttribute('aria-describedby', titleError.id);
    expect(titleError).toHaveTextContent('请输入文献标题。');
    expect(tag).not.toHaveAttribute('aria-invalid');

    fireEvent.change(title, {
      target: { value: 'Agentic Office Workflows' },
    });
    expect(screen.queryByRole('alert')).not.toBeInTheDocument();
    expect(title).not.toHaveAttribute('aria-invalid');

    fireEvent.change(tag, { target: { value: 'invalid tag' } });
    fireEvent.click(save);

    const tagError = await screen.findByRole('alert');
    await waitFor(() => expect(tag).toHaveFocus());
    expect(tag).toHaveAttribute('aria-invalid', 'true');
    expect(tag).toHaveAttribute('aria-describedby', tagError.id);
    expect(tagError).toHaveTextContent('简称只能使用字母、数字');
    expect(title).not.toHaveAttribute('aria-invalid');
  } finally {
    editor.destroy();
  }
});

test('confirms before deleting a citation source', async () => {
  const editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>References</p>',
  });
  const content: WorkDocumentContent = {
    type: 'document',
    html: editor.getHTML(),
    bibliography: { style: 'apa', sources },
  };
  render(
    <DocumentCitationsPanel
      editor={editor}
      content={content}
      onClose={() => undefined}
    />,
  );

  try {
    const title = screen.getByRole('textbox', { name: '文献标题' });
    title.focus();
    fireEvent.change(title, { target: { value: 'Edited before deletion' } });
    const deleteButton = screen.getByRole('button', { name: '删除文献' });
    deleteButton.focus();
    fireEvent.click(deleteButton);
    expect(screen.getByRole('dialog', { name: '删除文献？' })).toBeVisible();
    fireEvent.click(screen.getByRole('button', { name: '取消' }));
    await waitFor(() => expect(title).toHaveFocus());
    expect(title).toHaveValue('Edited before deletion');

    fireEvent.click(deleteButton);
    fireEvent.click(
      within(screen.getByRole('dialog', { name: '删除文献？' })).getByRole(
        'button',
        { name: '删除' },
      ),
    );
    await waitFor(() =>
      expect(screen.getByRole('textbox', { name: '文献标题' })).toHaveValue(
        'Beta Systems',
      ),
    );
  } finally {
    editor.destroy();
  }
});
