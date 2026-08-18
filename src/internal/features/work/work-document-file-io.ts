import { documentContentLayoutProperties } from './work-document-section';
import {
  downloadBlob,
  fileNameWithoutExtension,
  safeFileName,
} from './work-file-download';
import {
  readVerifiedWorkSourceBytes,
  rememberWorkSourceBlob,
  workSourceFingerprint,
} from './work-repository';
import type { WorkFileImportContext } from './work-file-import';
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
    const mammoth = await import('mammoth');
    const arrayBuffer = context?.bytes ?? (await file.arrayBuffer());
    context?.controller.report('parsing', 0.1);
    const { applyDocxSectionsToHtml, prepareDocxImport, readDocxLayout } =
      await import('./work-docx-import');
    const { OoxmlPackage } = await import('./work-ooxml-package');
    const sourcePackage = await OoxmlPackage.load(arrayBuffer).catch(
      () => null,
    );
    const prepared = await prepareDocxImport(
      arrayBuffer,
      sourcePackage ?? undefined,
    ).catch(() => null);
    context?.controller.report('parsing', 0.45);
    const conversionBuffer = prepared?.conversionBuffer ?? arrayBuffer;
    const input = {
      arrayBuffer: conversionBuffer,
      buffer: conversionBuffer,
    } as unknown as Parameters<typeof mammoth.convertToHtml>[0];
    const result = await mammoth.convertToHtml(input, {
      styleMap: [
        "br[type='page'] => hr.work-page-break[data-page-break='true']:fresh",
      ],
    });
    context?.controller.report('parsing', 0.8);
    html = prepared
      ? applyDocxSectionsToHtml(
          result.value,
          prepared.sections,
          prepared.captionMarkers,
          prepared.bookmarkMarkers,
          prepared.changeMarkers,
          prepared.commentMarkers,
          prepared.fieldMarkers,
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
    context?.controller.report('analyzing', 0);
    artifact.compatibility = await analyzeDocxCompatibility(
      file,
      result.messages,
      sourcePackage,
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
