import { readdir, readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import PptxGenJS from 'pptxgenjs';
import {
  createArtifactBlob,
  forgetSourceBlob,
  importOfficeFile,
} from '../../../src/core';

const fixturesDir = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  'fixtures',
);

/**
 * Browser hosts load pptxgen via a classic script tag that assigns
 * `window.PptxGenJS`. Unit/corpus tests under happy-dom cannot execute that
 * script file; install the same constructor from the package so export uses
 * the real runtime rather than a stub.
 */
function ensurePresentationExportRuntime(): void {
  if (typeof window === 'undefined') return;
  if (!window.PptxGenJS) {
    window.PptxGenJS = PptxGenJS;
  }
}

export type RealCorpusKind = 'document' | 'spreadsheet' | 'presentation' | 'pdf';

export type RealCorpusRoundTrip = {
  fileName: string;
  kind: RealCorpusKind;
  firstPassIssueCodes: string[];
  secondPassIssueCodes: string[];
  firstPassParts: string[];
  secondPassParts: string[];
};

const kindByExtension: Record<string, RealCorpusKind> = {
  '.docx': 'document',
  '.xlsx': 'spreadsheet',
  '.pptx': 'presentation',
  '.pdf': 'pdf',
};

const mimeByKind: Record<RealCorpusKind, string> = {
  document:
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  spreadsheet:
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  presentation:
    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  pdf: 'application/pdf',
};

export async function listRealCorpusFixtures(): Promise<string[]> {
  const entries = await readdir(fixturesDir);
  return entries
    .filter((name) => kindByExtension[path.extname(name).toLowerCase()])
    .sort();
}

export async function roundTripRealCorpusFile(
  fileName: string,
): Promise<RealCorpusRoundTrip> {
  ensurePresentationExportRuntime();
  const extension = path.extname(fileName).toLowerCase();
  const kind = kindByExtension[extension];
  if (!kind) {
    throw new Error(`Unsupported real-corpus fixture: ${fileName}`);
  }
  const bytes = await readFile(path.join(fixturesDir, fileName));
  const source = new File([bytes], fileName, { type: mimeByKind[kind] });
  const imported = await importOfficeFile(source);
  try {
    if (imported.content.type !== kind) {
      throw new Error(
        `Expected ${kind} for ${fileName}, got ${imported.content.type}.`,
      );
    }
    const firstPassIssueCodes = (imported.compatibility?.issues ?? []).map(
      (issue) => issue.code,
    );
    const firstPassParts = await packagePartNames(imported);
    const exported = await createArtifactBlob(imported);
    const reopened = await importOfficeFile(
      new File([exported], `reopened-${fileName}`, { type: exported.type }),
    );
    try {
      if (reopened.content.type !== kind) {
        throw new Error(
          `Reopen kind mismatch for ${fileName}: ${reopened.content.type}.`,
        );
      }
      return {
        fileName,
        kind,
        firstPassIssueCodes,
        secondPassIssueCodes: (reopened.compatibility?.issues ?? []).map(
          (issue) => issue.code,
        ),
        firstPassParts,
        secondPassParts: await packagePartNames(reopened),
      };
    } finally {
      forgetSourceBlob(reopened);
    }
  } finally {
    forgetSourceBlob(imported);
  }
}

async function packagePartNames(
  artifact: Awaited<ReturnType<typeof importOfficeFile>>,
): Promise<string[]> {
  const blob = await createArtifactBlob(artifact);
  if (artifact.content.type === 'pdf') {
    return [`pdf:pages=${await pdfPageCount(blob)}`];
  }
  const archive = await (await import('jszip')).default.loadAsync(
    await blob.arrayBuffer(),
  );
  return Object.keys(archive.files)
    .filter((part) => !archive.files[part]?.dir)
    .sort();
}

async function pdfPageCount(blob: Blob): Promise<number> {
  const { PDFDocument } = await import('pdf-lib');
  // Some redistributable samples (e.g. PDFBox fuzz fixtures) carry an
  // encryption dictionary with empty/open permissions. We only need the page
  // count for the corpus gate, so ignore the encryption wrapper.
  const document = await PDFDocument.load(await blob.arrayBuffer(), {
    ignoreEncryption: true,
  });
  return document.getPageCount();
}
