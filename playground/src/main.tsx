import { StrictMode, useCallback, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import {
  createArtifact,
  downloadArtifact,
  importOfficeFile,
  OFFICE_FILE_ACCEPT,
  readSourceBlob,
  registerSourceBlob,
  type DocumentContent,
  type MarkdownContent,
  type OfficeArtifact,
  type PresentationContent,
  type SpreadsheetContent,
} from '@a3s-lab/office/core';
import {
  DocumentEditor,
  MarkdownEditor,
  PdfViewer,
  PresentationEditor,
  SpreadsheetEditor,
} from '@a3s-lab/office/react';
import '@a3s-lab/office/styles.css';
import './playground.css';

type EditableKind = 'document' | 'markdown' | 'spreadsheet' | 'presentation';

const templateByKind: Record<EditableKind, string> = {
  document: 'blank-document',
  markdown: 'blank-markdown',
  presentation: 'blank-presentation',
  spreadsheet: 'blank-spreadsheet',
};

function Playground() {
  const [artifact, setArtifact] = useState<OfficeArtifact>(() =>
    createArtifact('blank-document'),
  );
  const [preview, setPreview] = useState(false);
  const [message, setMessage] = useState('Ready');
  const fileInput = useRef<HTMLInputElement>(null);

  const newArtifact = (kind: EditableKind) => {
    setArtifact(createArtifact(templateByKind[kind]));
    setPreview(false);
    setMessage(`Created a blank ${kind}`);
  };

  const replaceContent = (
    content:
      | DocumentContent
      | MarkdownContent
      | SpreadsheetContent
      | PresentationContent,
  ) => {
    setArtifact((current) => ({
      ...current,
      content,
      kind: content.type,
      revision: current.revision + 1,
      updatedAt: Date.now(),
    }));
    setMessage('Edited locally');
  };

  const importFile = async (file: File) => {
    try {
      const imported = await importOfficeFile(file);
      setArtifact(imported);
      setPreview(false);
      setMessage(`Imported ${file.name}`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : 'Import failed');
    }
  };

  const loadPdf = useCallback(() => readSourceBlob(artifact), [artifact]);
  const savePdf = useCallback(
    async (pdf: Blob) => {
      registerSourceBlob(artifact.id, pdf);
      setMessage('Saved PDF changes in this browser session');
      return true;
    },
    [artifact.id],
  );
  const playgroundAssetUrl = (fileName: string) =>
    new URL(fileName, document.baseURI).href;

  return (
    <main className="playground-shell">
      <header className="playground-header">
        <div>
          <strong>A3S Office</strong>
          <span>{artifact.title}</span>
        </div>
        <nav aria-label="Create an editor">
          <button type="button" onClick={() => newArtifact('document')}>
            Document
          </button>
          <button type="button" onClick={() => newArtifact('markdown')}>
            Markdown
          </button>
          <button type="button" onClick={() => newArtifact('spreadsheet')}>
            Spreadsheet
          </button>
          <button type="button" onClick={() => newArtifact('presentation')}>
            Presentation
          </button>
        </nav>
        <div className="playground-actions">
          <input
            ref={fileInput}
            hidden
            type="file"
            accept={OFFICE_FILE_ACCEPT}
            onChange={(event) => {
              const file = event.target.files?.[0];
              event.target.value = '';
              if (file) void importFile(file);
            }}
          />
          <button type="button" onClick={() => fileInput.current?.click()}>
            Import
          </button>
          <button
            type="button"
            disabled={artifact.kind === 'pdf'}
            onClick={() => setPreview((value) => !value)}
          >
            {preview ? 'Edit' : 'Preview'}
          </button>
          <button
            type="button"
            onClick={() =>
              void downloadArtifact(
                artifact,
                artifact.kind === 'presentation'
                  ? {
                      pptxRuntimeUrl: playgroundAssetUrl('pptxgen.bundle.js'),
                    }
                  : undefined,
              ).catch((error: unknown) =>
                setMessage(
                  error instanceof Error ? error.message : 'Export failed',
                ),
              )
            }
          >
            Export
          </button>
        </div>
      </header>

      <section className="playground-editor" aria-label="Editor playground">
        {artifact.content.type === 'document' && (
          <DocumentEditor
            content={artifact.content}
            onChange={replaceContent}
            preview={preview}
          />
        )}
        {artifact.content.type === 'spreadsheet' && (
          <SpreadsheetEditor
            content={artifact.content}
            onChange={replaceContent}
            preview={preview}
          />
        )}
        {artifact.content.type === 'markdown' && (
          <MarkdownEditor
            content={artifact.content}
            onChange={replaceContent}
            preview={preview}
          />
        )}
        {artifact.content.type === 'presentation' && (
          <PresentationEditor
            content={artifact.content}
            onChange={replaceContent}
            onStartSlideshow={() => setPreview(true)}
            preview={preview}
          />
        )}
        {artifact.content.type === 'pdf' && (
          <PdfViewer
            fileName={artifact.source?.name ?? `${artifact.title}.pdf`}
            loadSource={loadPdf}
            onSave={savePdf}
            sourceKey={`${artifact.id}:${artifact.revision}`}
            wasmUrl={playgroundAssetUrl('pdfium.wasm')}
          />
        )}
      </section>

      <footer className="playground-status" aria-live="polite">
        <span>{message}</span>
        <span>Local browser demo · no server required</span>
      </footer>
    </main>
  );
}

const root = document.getElementById('root');
if (!root) throw new Error('Playground root element is missing.');

createRoot(root).render(
  <StrictMode>
    <Playground />
  </StrictMode>,
);
