import { documentContentLayoutProperties } from './work-document-section';
import {
  downloadBlob,
  fileNameWithoutExtension,
  safeFileName,
} from './work-file-download';
import type { WorkFileImportContext } from './work-file-import';
import {
  readVerifiedWorkSourceBytes,
  rememberWorkSourceBlob,
  workSourceFingerprint,
} from './work-repository';
import { createWorkArtifact } from './work-templates';
import type { WorkArtifact, WorkDocumentContent } from './work-types';

const DOCX_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

export async function importWorkDocumentFile(
  file: File,
  extension: string,
  context?: WorkFileImportContext,
): Promise<WorkArtifact> {
  let html: string;
  if (extension === 'docx') {
    const importStartedAt = documentImportNow();
    const arrayBuffer = context?.bytes ?? (await file.arrayBuffer());
    context?.controller.report('parsing', 0.1);
    const { OoxmlPackage } = await import('./work-ooxml-package');
    const packageStartedAt = documentImportNow();
    const sourcePackage = await OoxmlPackage.load(arrayBuffer).catch(
      () => null,
    );
    recordDocumentImportMeasure(
      'a3s-office.document.package',
      packageStartedAt,
      documentImportNow(),
    );
    const largeSimpleStartedAt = documentImportNow();
    const {
      largeSimpleDocxPackageIsCandidate,
      largeSimpleDocxCompatibilityReport,
      largeSimpleDocxContent,
      parseLargeSimpleDocxPackage,
    } = await import('./work-docx-large-document-import');
    let largeSimple = null;
    let largeSimplePath = 'ineligible';
    if (sourcePackage && largeSimpleDocxPackageIsCandidate(sourcePackage)) {
      const { parseLargeSimpleDocxInWorker } = await import(
        './work-document-import-worker-client'
      );
      const inflateStartedAt = documentImportNow();
      const documentXmlBytes = await sourcePackage
        .bytes('word/document.xml')
        .catch(() => null);
      recordDocumentImportMeasure(
        'a3s-office.document.large-simple-inflate',
        inflateStartedAt,
        documentImportNow(),
      );
      const workerAttempt = documentXmlBytes
        ? await parseLargeSimpleDocxInWorker(
            transferableDocumentXmlBuffer(documentXmlBytes),
            {},
            context?.controller.signal,
          )
        : null;
      if (workerAttempt?.status === 'accepted') {
        largeSimple = workerAttempt.result;
        largeSimplePath = 'worker';
      } else if (!workerAttempt) {
        largeSimple = await parseLargeSimpleDocxPackage(sourcePackage).catch(
          () => null,
        );
        largeSimplePath = largeSimple ? 'synchronous-fallback' : 'ineligible';
      }
    }
    recordDocumentImportMeasure(
      'a3s-office.document.large-simple-inspection',
      largeSimpleStartedAt,
      documentImportNow(),
      { accepted: Boolean(largeSimple), path: largeSimplePath },
    );
    if (largeSimple) {
      context?.controller.report('parsing', 0.8);
      const modelStartedAt = documentImportNow();
      const { createSchemaValidatedWorkDocumentModel } = await import(
        './work-document-model'
      );
      const artifact = createWorkArtifact('blank-document');
      artifact.title = fileNameWithoutExtension(file.name);
      const fastContent = largeSimpleDocxContent(largeSimple);
      artifact.content = {
        ...fastContent,
        model: createSchemaValidatedWorkDocumentModel(
          largeSimple.html,
          largeSimple.root,
          {
            initialIntegrityFeatures: 0,
            previous: fastContent.model,
          },
        ),
      };
      recordDocumentImportMeasure(
        'a3s-office.document.structured-model',
        modelStartedAt,
        documentImportNow(),
        {
          logicalBlocks: largeSimple.logicalBlockCount,
          tableRows: largeSimple.tableRowCount,
          windowed: true,
        },
      );
      context?.controller.report('analyzing', 0);
      artifact.compatibility = largeSimpleDocxCompatibilityReport(
        file,
        largeSimple,
      );
      context?.controller.report('analyzing', 1);
      artifact.source = {
        name: file.name,
        contentType: file.type || DOCX_CONTENT_TYPE,
        size: file.size,
        updatedAt: file.lastModified || Date.now(),
        fingerprint: await workSourceFingerprint(arrayBuffer),
      };
      rememberWorkSourceBlob(artifact.id, file);
      recordDocumentImportMeasure(
        'a3s-office.document.import-total',
        importStartedAt,
        documentImportNow(),
        { path: 'large-simple' },
      );
      return artifact;
    }

    const mammoth = await import('mammoth');
    const { applyDocxSectionsToHtml, prepareDocxImport, readDocxLayout } =
      await import('./work-docx-import');
    const prepareStartedAt = documentImportNow();
    const prepared = await prepareDocxImport(
      arrayBuffer,
      sourcePackage ?? undefined,
    ).catch(() => null);
    recordDocumentImportMeasure(
      'a3s-office.document.prepare',
      prepareStartedAt,
      documentImportNow(),
    );
    context?.controller.report('parsing', 0.45);
    const conversionBuffer = prepared?.conversionBuffer ?? arrayBuffer;
    const input = {
      arrayBuffer: conversionBuffer,
      buffer: conversionBuffer,
    } as unknown as Parameters<typeof mammoth.convertToHtml>[0];
    const mammothStartedAt = documentImportNow();
    const result = await mammoth.convertToHtml(input, {
      styleMap: [
        "br[type='page'] => hr.work-page-break[data-page-break='true']:fresh",
      ],
    });
    recordDocumentImportMeasure(
      'a3s-office.document.mammoth',
      mammothStartedAt,
      documentImportNow(),
    );
    context?.controller.report('parsing', 0.8);
    const markersStartedAt = documentImportNow();
    html = prepared
      ? applyDocxSectionsToHtml(
          result.value,
          prepared.sections,
          prepared.captionMarkers,
          prepared.bookmarkMarkers,
          prepared.changeMarkers,
          prepared.commentMarkers,
          prepared.fieldMarkers,
          prepared.tableOfContentsMarkers,
          prepared.indexMarkers,
          prepared.equationMarkers,
          prepared.citationMarkers,
          prepared.listMarkers,
          prepared.imageLayoutMarkers,
          prepared.paragraphIdentityMarkers,
          prepared.paragraphFormattingChangeMarkers,
          prepared.paragraphAlignmentMarkers,
          prepared.runFormattingMarkers,
          prepared.paragraphDirectionMarkers,
          prepared.paragraphIndentMarkers,
          prepared.paragraphSpacingMarkers,
          prepared.paragraphBorderMarkers,
          prepared.paragraphShadingMarkers,
          prepared.paragraphPaginationMarkers,
          prepared.bibliography,
          prepared.tabStopMarkers,
          prepared.tableCellMarkers,
          prepared.tableRowMarkers,
          prepared.tableSizingMarkers,
        )
      : result.value;
    recordDocumentImportMeasure(
      'a3s-office.document.markers',
      markersStartedAt,
      documentImportNow(),
    );
    const layout = prepared
      ? {
          ...documentContentLayoutProperties(prepared.sections[0].layout),
          ...(prepared.pageColor ? { pageColor: prepared.pageColor } : {}),
        }
      : await readDocxLayout(arrayBuffer).catch(() => ({
          pageSize: 'a4' as const,
        }));
    const { analyzeDocxCompatibility } = await import(
      './work-office-diagnostics'
    );
    const artifact = createWorkArtifact('blank-document');
    artifact.title = fileNameWithoutExtension(file.name);
    const modelStartedAt = documentImportNow();
    artifact.content = await withStructuredDocumentModel({
      type: 'document',
      html,
      ...layout,
      ...(prepared?.trackChanges ? { trackChanges: true } : {}),
      ...(prepared?.commentMarkers.comments.length
        ? { comments: prepared.commentMarkers.comments }
        : {}),
      ...(prepared?.bibliography
        ? { bibliography: prepared.bibliography }
        : {}),
    });
    recordDocumentImportMeasure(
      'a3s-office.document.structured-model',
      modelStartedAt,
      documentImportNow(),
      { windowed: false },
    );
    context?.controller.report('analyzing', 0);
    const compatibilityStartedAt = documentImportNow();
    artifact.compatibility = await analyzeDocxCompatibility(
      file,
      result.messages,
      sourcePackage,
    );
    recordDocumentImportMeasure(
      'a3s-office.document.compatibility',
      compatibilityStartedAt,
      documentImportNow(),
    );
    context?.controller.report('analyzing', 1);
    artifact.source = {
      name: file.name,
      contentType: file.type || DOCX_CONTENT_TYPE,
      size: file.size,
      updatedAt: file.lastModified || Date.now(),
      fingerprint: await workSourceFingerprint(arrayBuffer),
    };
    rememberWorkSourceBlob(artifact.id, file);
    recordDocumentImportMeasure(
      'a3s-office.document.import-total',
      importStartedAt,
      documentImportNow(),
      { path: 'complete' },
    );
    return artifact;
  }

  const source = context
    ? new TextDecoder().decode(context.bytes)
    : await file.text();
  html =
    extension === 'html' || extension === 'htm' ? source : textToHtml(source);
  const artifact = createWorkArtifact('blank-document');
  artifact.title = fileNameWithoutExtension(file.name);
  artifact.content = await withStructuredDocumentModel({
    type: 'document',
    html,
    pageSize: 'a4',
  });
  context?.controller.report('parsing', 1);
  return artifact;
}

export async function exportWorkDocumentArtifact(
  artifact: WorkArtifact,
): Promise<void> {
  downloadBlob(
    await createWorkDocumentBlob(artifact),
    `${safeFileName(artifact.title)}.docx`,
  );
}

export async function createWorkDocumentBlob(
  artifact: WorkArtifact,
): Promise<Blob> {
  if (artifact.content.type !== 'document')
    throw new Error('当前文件不是文档。');
  const { materializeWorkDocumentContent } = await import(
    './work-document-model-codec'
  );
  const { createDocxBlob } = await import('./work-docx-export');
  const sourcePackage = artifact.source
    ? await readVerifiedWorkSourceBytes(artifact)
    : undefined;
  return createDocxBlob(
    materializeWorkDocumentContent(artifact.content),
    sourcePackage,
  );
}

function textToHtml(source: string): string {
  const blocks = source
    .replace(/\r\n?/g, '\n')
    .split(/\n{2,}/)
    .map((block) => block.trim())
    .filter(Boolean);
  return blocks
    .map((block) => {
      if (block.startsWith('# '))
        return `<h1>${escapeHtml(block.slice(2))}</h1>`;
      if (block.startsWith('## '))
        return `<h2>${escapeHtml(block.slice(3))}</h2>`;
      return `<p>${escapeHtml(block).replace(/\n/g, '<br>')}</p>`;
    })
    .join('');
}

async function withStructuredDocumentModel(
  content: WorkDocumentContent,
): Promise<WorkDocumentContent> {
  const { createWorkDocumentModelFromContent } = await import(
    './work-document-model-codec'
  );
  return createWorkDocumentModelFromContent(content);
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;');
}

function documentImportNow(): number {
  return globalThis.performance?.now?.() ?? Date.now();
}

function transferableDocumentXmlBuffer(bytes: Uint8Array): ArrayBuffer {
  if (
    bytes.byteOffset === 0 &&
    bytes.byteLength === bytes.buffer.byteLength &&
    bytes.buffer instanceof ArrayBuffer
  ) {
    return bytes.buffer;
  }
  return bytes.slice().buffer;
}

function recordDocumentImportMeasure(
  name: string,
  start: number,
  end: number,
  detail?: Record<string, unknown>,
): void {
  try {
    globalThis.performance?.measure(name, { detail, end, start });
  } catch {
    // User Timing diagnostics must never affect document import.
  }
}
