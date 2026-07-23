import { fileNameWithoutExtension } from './work-file-download';
import { createWorkArtifact } from './work-templates';
import type { WorkArtifact } from './work-types';

const MARKDOWN_CONTENT_TYPE = 'text/markdown;charset=utf-8';

export async function importWorkMarkdownFile(
  file: File,
): Promise<WorkArtifact> {
  const artifact = createWorkArtifact('blank-markdown');
  artifact.title = fileNameWithoutExtension(file.name);
  artifact.content = {
    type: 'markdown',
    markdown: await file.text(),
  };
  return artifact;
}

export function createWorkMarkdownBlob(artifact: WorkArtifact): Blob {
  if (artifact.content.type !== 'markdown')
    throw new Error('当前文件不是 Markdown。');
  return new Blob([artifact.content.markdown], {
    type: MARKDOWN_CONTENT_TYPE,
  });
}
