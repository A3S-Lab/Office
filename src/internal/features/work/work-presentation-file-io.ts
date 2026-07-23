import {
  downloadBlob,
  fileNameWithoutExtension,
  safeFileName,
} from './work-file-download';
import { createWorkArtifact } from './work-templates';
import type { WorkArtifact } from './work-types';

const PPTX_RUNTIME_FILE_NAME = 'pptxgen.bundle.js';

type PptxConstructor = typeof import('pptxgenjs').default;

const presentationRuntimePromises = new Map<string, Promise<PptxConstructor>>();

declare global {
  interface Window {
    PptxGenJS?: PptxConstructor;
  }
}

export interface WorkPresentationExportOptions {
  pptxRuntimeUrl?: string;
}

export const defaultPptxRuntimeUrl = siblingAssetUrl(
  import.meta.url,
  PPTX_RUNTIME_FILE_NAME,
);

export async function importWorkPresentationFile(
  file: File,
): Promise<WorkArtifact> {
  const { importPptxPresentation } = await import('./work-pptx-import');
  const imported = await importPptxPresentation(file);
  const artifact = createWorkArtifact('blank-presentation');
  artifact.title = fileNameWithoutExtension(file.name);
  artifact.content = imported.content;
  artifact.compatibility = imported.compatibility;
  return artifact;
}

export async function exportWorkPresentationArtifact(
  artifact: WorkArtifact,
  options?: WorkPresentationExportOptions,
): Promise<void> {
  downloadBlob(
    await createWorkPresentationBlob(artifact, options),
    `${safeFileName(artifact.title)}.pptx`,
  );
}

export async function createWorkPresentationBlob(
  artifact: WorkArtifact,
  options?: WorkPresentationExportOptions,
): Promise<Blob> {
  if (artifact.content.type !== 'presentation')
    throw new Error('当前文件不是演示文稿。');
  const PptxGenJS = await loadPresentationRuntime(
    options?.pptxRuntimeUrl ?? defaultPptxRuntimeUrl,
  );
  const { createPptxBlob } = await import('./work-pptx-export');
  return createPptxBlob(artifact, PptxGenJS);
}

function loadPresentationRuntime(runtimeUrl: string): Promise<PptxConstructor> {
  if (typeof window === 'undefined' || typeof document === 'undefined') {
    return Promise.reject(
      new Error('PowerPoint export is only available in a browser.'),
    );
  }
  if (window.PptxGenJS) return Promise.resolve(window.PptxGenJS);

  const resolvedUrl = new URL(runtimeUrl, document.baseURI).href;
  const pending = presentationRuntimePromises.get(resolvedUrl);
  if (pending) return pending;

  const runtime = new Promise<PptxConstructor>((resolve, reject) => {
    const existing = Array.from(
      document.querySelectorAll<HTMLScriptElement>(
        'script[data-a3s-pptx-runtime]',
      ),
    ).find((script) => script.src === resolvedUrl);
    const script = existing ?? document.createElement('script');
    const finish = () => {
      if (window.PptxGenJS) {
        resolve(window.PptxGenJS);
        return;
      }
      reject(new Error('PowerPoint export runtime did not initialize.'));
    };
    script.addEventListener('load', finish, { once: true });
    script.addEventListener(
      'error',
      () =>
        reject(
          new Error(
            `PowerPoint export runtime could not be loaded from ${resolvedUrl}.`,
          ),
        ),
      { once: true },
    );
    if (!existing) {
      script.src = resolvedUrl;
      script.dataset.a3sPptxRuntime = 'true';
      document.head.append(script);
    }
  }).catch((error) => {
    presentationRuntimePromises.delete(resolvedUrl);
    throw error;
  });
  presentationRuntimePromises.set(resolvedUrl, runtime);
  return runtime;
}

function siblingAssetUrl(moduleUrl: string, fileName: string): string {
  return `${moduleUrl.slice(0, moduleUrl.lastIndexOf('/') + 1)}${fileName}`;
}
