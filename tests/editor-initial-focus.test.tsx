import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useState } from 'react';
import {
  createArtifact,
  type MarkdownContent,
  type PresentationContent,
  type SpreadsheetContent,
} from '../src/core';
import {
  DocumentEditor,
  MarkdownEditor,
  PresentationEditor,
  SpreadsheetEditor,
} from '../src/react';

test('focuses an editable document body when the editor first opens', async () => {
  render(<DocumentOpeningHarness />);

  const opener = screen.getByRole('button', { name: 'Open document' });
  opener.focus();
  fireEvent.click(opener);

  const body = await screen.findByRole('textbox', { name: '文档正文' });
  await waitFor(() => expect(body).toHaveFocus());
  expect(body).toHaveAttribute('contenteditable', 'true');
});

test('focuses the editable Markdown source and accepts the first change', async () => {
  render(<MarkdownOpeningHarness />);

  const opener = screen.getByRole('button', { name: 'Open Markdown' });
  opener.focus();
  fireEvent.click(opener);

  const source = await screen.findByLabelText('Markdown 源码');
  await waitFor(() => expect(source).toHaveFocus());
  expect(source).not.toHaveAttribute('readonly');

  fireEvent.change(source, { target: { value: 'First keystroke' } });
  expect(source).toHaveValue('First keystroke');
});

test('does not steal focus when the user moves on while an editor opens', async () => {
  render(<DocumentOpeningHarness showNextControl />);

  const opener = screen.getByRole('button', { name: 'Open document' });
  const next = screen.getByRole('button', { name: 'Next control' });
  opener.focus();
  fireEvent.click(opener);
  next.focus();

  await screen.findByRole('textbox', { name: '文档正文' });
  await waitFor(() => expect(next).toHaveFocus());
});

test('keeps opener focus when first-open focus is disabled', async () => {
  render(<DocumentOpeningHarness autoFocus={false} />);

  const opener = screen.getByRole('button', { name: 'Open document' });
  opener.focus();
  fireEvent.click(opener);

  await screen.findByRole('textbox', { name: '文档正文' });
  await waitFor(() => expect(opener).toHaveFocus());
});

test('focuses the spreadsheet grid and starts editing on its first printable key', async () => {
  const view = render(<SpreadsheetOpeningHarness />);

  const opener = screen.getByRole('button', { name: 'Open spreadsheet' });
  opener.focus();
  fireEvent.click(opener);

  const grid = await waitFor(() => {
    const target = view.container.querySelector<HTMLElement>(
      '.fortune-sheet-overlay',
    );
    expect(target).not.toBeNull();
    return target as HTMLElement;
  });
  await waitFor(() => expect(grid).toHaveFocus());

  fireEvent.keyDown(grid, { key: 'x', code: 'KeyX' });

  await waitFor(() => {
    const target = view.container.querySelector<HTMLElement>(
      '.luckysheet-cell-input',
    );
    expect(target).not.toBeNull();
    expect(target).toHaveFocus();
  });
});

test('focuses the empty title editor when a presentation first opens', async () => {
  render(<PresentationOpeningHarness />);

  const opener = screen.getByRole('button', { name: 'Open presentation' });
  opener.focus();
  fireEvent.click(opener);

  const title = await screen.findByRole('textbox', { name: '幻灯片文本' });
  await waitFor(() => expect(title).toHaveFocus());
  expect(title).toHaveAttribute('contenteditable', 'true');
});

test('keeps opener focus when presentation first-open focus is disabled', async () => {
  render(<PresentationOpeningHarness autoFocus={false} />);

  const opener = screen.getByRole('button', { name: 'Open presentation' });
  opener.focus();
  fireEvent.click(opener);

  await screen.findByRole('group', { name: '单击添加标题' });
  expect(screen.queryByRole('textbox', { name: '幻灯片文本' })).toBeNull();
  expect(opener).toHaveFocus();
});

test('does not steal focus when the user moves on while a presentation opens', async () => {
  render(<PresentationOpeningHarness showNextControl />);

  const opener = screen.getByRole('button', { name: 'Open presentation' });
  const next = screen.getByRole('button', {
    name: 'Next presentation control',
  });
  opener.focus();
  fireEvent.click(opener);
  next.focus();

  await screen.findByRole('textbox', { name: '幻灯片文本' });
  await waitFor(() => expect(next).toHaveFocus());
});

function DocumentOpeningHarness({
  autoFocus,
  showNextControl = false,
}: {
  autoFocus?: boolean;
  showNextControl?: boolean;
}) {
  const [open, setOpen] = useState(false);
  const artifact = createArtifact('blank-document');
  return (
    <>
      <button type="button" onClick={() => setOpen(true)}>
        Open document
      </button>
      {showNextControl && <button type="button">Next control</button>}
      {open && (
        <DocumentEditor
          autoFocus={autoFocus}
          content={artifact.content}
          onChange={() => undefined}
          theme="light"
        />
      )}
    </>
  );
}

function MarkdownOpeningHarness() {
  const [open, setOpen] = useState(false);
  const [content, setContent] = useState<MarkdownContent>({
    type: 'markdown',
    markdown: '',
  });
  return (
    <>
      <button type="button" onClick={() => setOpen(true)}>
        Open Markdown
      </button>
      {open && (
        <MarkdownEditor content={content} onChange={setContent} theme="light" />
      )}
    </>
  );
}

function SpreadsheetOpeningHarness() {
  const [open, setOpen] = useState(false);
  const [content, setContent] = useState<SpreadsheetContent>(blankSpreadsheet);
  return (
    <>
      <button type="button" onClick={() => setOpen(true)}>
        Open spreadsheet
      </button>
      {open && (
        <SpreadsheetEditor
          content={content}
          onChange={setContent}
          theme="light"
        />
      )}
    </>
  );
}

function PresentationOpeningHarness({
  autoFocus,
  showNextControl = false,
}: {
  autoFocus?: boolean;
  showNextControl?: boolean;
}) {
  const [open, setOpen] = useState(false);
  const [content, setContent] =
    useState<PresentationContent>(blankPresentation);
  return (
    <>
      <button type="button" onClick={() => setOpen(true)}>
        Open presentation
      </button>
      {showNextControl && (
        <button type="button">Next presentation control</button>
      )}
      {open && (
        <PresentationEditor
          autoFocus={autoFocus}
          content={content}
          onChange={setContent}
          theme="light"
        />
      )}
    </>
  );
}

function blankSpreadsheet(): SpreadsheetContent {
  const artifact = createArtifact('blank-spreadsheet');
  if (artifact.content.type !== 'spreadsheet') {
    throw new Error('Expected blank spreadsheet content.');
  }
  return artifact.content;
}

function blankPresentation(): PresentationContent {
  const artifact = createArtifact('blank-presentation');
  if (artifact.content.type !== 'presentation') {
    throw new Error('Expected blank presentation content.');
  }
  return artifact.content;
}
