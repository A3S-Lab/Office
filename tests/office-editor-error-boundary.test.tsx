import { expect, test } from '@rstest/core';
import { render, screen, waitFor } from '@testing-library/react';
import { createArtifact, type DocumentContent } from '../src/core';
import {
  DocumentEditor,
  type OfficeEditorHostError,
  OfficeEditorErrorBoundary,
  resolveOfficeMessages,
} from '../src/react';

const zhMessages = resolveOfficeMessages({ locale: 'zh-CN' });
const enMessages = resolveOfficeMessages({ locale: 'en-US' });

test('OfficeEditorErrorBoundary keeps the host mounted and notifies onError once', async () => {
  const calls: OfficeEditorHostError[] = [];
  const ThrowOnce = () => {
    throw new Error('synthetic render failure');
  };

  render(
    <div data-testid="host">
      <OfficeEditorErrorBoundary
        editor="document"
        messages={zhMessages}
        onError={(error) => {
          calls.push(error);
        }}
        title="Writer"
      >
        <ThrowOnce />
      </OfficeEditorErrorBoundary>
    </div>,
  );

  expect(screen.getByTestId('host')).toBeInTheDocument();
  expect(await screen.findByRole('alert')).toHaveAttribute(
    'data-office-editor-error',
    'render',
  );
  expect(calls).toHaveLength(1);
  expect(calls[0]?.editor).toBe('document');
  expect(calls[0]?.phase).toBe('render');
  expect(String(calls[0]?.error)).toMatch(/synthetic render failure/);
});

test('DocumentEditor onError is transparent when the editor mounts normally', async () => {
  const calls: OfficeEditorHostError[] = [];
  const artifact = createArtifact('blank-document');

  render(
    <DocumentEditor
      artifactId={artifact.id}
      content={artifact.content as DocumentContent}
      onChange={() => undefined}
      onError={(error) => {
        calls.push(error);
      }}
      preview
      theme="light"
    />,
  );

  expect(await screen.findByLabelText('文字预览')).toBeInTheDocument();
  expect(document.querySelector('[data-a3s-office]')).toBeTruthy();
  expect(calls).toHaveLength(0);
});

test('OfficeEditorErrorBoundary classifies chunk-load failures', async () => {
  const calls: OfficeEditorHostError[] = [];
  const ThrowChunk = () => {
    const error = new Error('Loading chunk 9 failed.');
    error.name = 'ChunkLoadError';
    throw error;
  };

  render(
    <OfficeEditorErrorBoundary
      editor="spreadsheet"
      messages={enMessages}
      onError={(error) => {
        calls.push(error);
      }}
      title="Spreadsheet"
    >
      <ThrowChunk />
    </OfficeEditorErrorBoundary>,
  );

  expect(await screen.findByRole('alert')).toHaveAttribute(
    'data-office-editor-error',
    'chunk-load',
  );
  expect(screen.getByRole('button', { name: 'Retry' })).toBeInTheDocument();
  await waitFor(() => expect(calls).toHaveLength(1));
  expect(calls[0]).toMatchObject({
    editor: 'spreadsheet',
    phase: 'chunk-load',
  });
});
