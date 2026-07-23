import type { WorkArtifact } from './work-types';

const sourceBlobs = new Map<string, Blob>();

export function rememberWorkSourceBlob(artifactId: string, source: Blob): void {
  sourceBlobs.set(artifactId, source);
}

export function forgetWorkSourceBlob(artifactId: string): void {
  sourceBlobs.delete(artifactId);
}

export async function readWorkSourceBlob(
  artifact: WorkArtifact,
): Promise<Blob> {
  const source = sourceBlobs.get(artifact.id);
  if (source) return source;
  throw new Error(
    artifact.source
      ? 'The original file is not available in this browser session. Register it again before opening the PDF.'
      : 'This document has no original source file.',
  );
}
