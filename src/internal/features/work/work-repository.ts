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
      ? 'The original file is not available in this browser session. Register it again before opening or exporting this source-backed artifact.'
      : 'This document has no original source file.',
  );
}

export async function readVerifiedWorkSourceBytes(
  artifact: WorkArtifact,
): Promise<ArrayBuffer> {
  const source = await readWorkSourceBlob(artifact);
  const bytes = await source.arrayBuffer();
  const expected = artifact.source?.fingerprint;
  if (expected && (await workSourceFingerprint(bytes)) !== expected) {
    throw new Error(
      'The registered source file does not match the imported source fingerprint.',
    );
  }
  return bytes;
}

export async function workSourceFingerprint(
  bytes: ArrayBuffer,
): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) throw new Error('SHA-256 source fingerprinting is unavailable.');
  const digest = new Uint8Array(await subtle.digest('SHA-256', bytes));
  return `sha256:${Array.from(digest, (value) =>
    value.toString(16).padStart(2, '0'),
  ).join('')}`;
}
