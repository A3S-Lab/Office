import {
  createArtifactBlob,
  forgetSourceBlob,
  importOfficeFile,
} from '../../../src/core';
import type { WorkCompatibilityIssue } from '../../../src/internal/features/work/work-types';

export type R0CommentIdentity = {
  author: string;
  text: string;
};

export type R0IdentitySnapshot = {
  bookmarkNames: string[];
  changeAuthors: string[];
  changeTexts: string[];
  commentAuthors: string[];
  commentTexts: string[];
  contentControlAliases: string[];
  contentControlTags: string[];
  contentControlTexts: string[];
  endnoteTexts: string[];
  footerTexts: string[];
  footnoteTexts: string[];
  headerTexts: string[];
  hrefs: string[];
  plainText: string;
  tableCellTexts: string[];
};

export type R0RoundTripResult = {
  firstPassIssues: WorkCompatibilityIssue[];
  identities: R0IdentitySnapshot;
  secondPassIssues: WorkCompatibilityIssue[];
  exportedParts: string[];
};

export async function roundTripDocx(
  bytes: Uint8Array,
  fileName: string,
): Promise<R0RoundTripResult> {
  const source = new File([bytes], fileName, {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  const imported = await importOfficeFile(source);
  try {
    if (imported.content.type !== 'document') {
      throw new Error(`Expected a document artifact for ${fileName}.`);
    }
    const firstPassIssues = imported.compatibility?.issues ?? [];
    const exported = await createArtifactBlob(imported);
    const archive = await (await import('jszip')).default.loadAsync(
      await exported.arrayBuffer(),
    );
    const exportedParts = Object.keys(archive.files)
      .filter((path) => !archive.files[path]?.dir)
      .sort();
    const reopened = await importOfficeFile(
      new File([exported], `reopened-${fileName}`, { type: exported.type }),
    );
    try {
      if (reopened.content.type !== 'document') {
        throw new Error(`Expected a reopened document for ${fileName}.`);
      }
      return {
        firstPassIssues,
        identities: extractDocumentIdentities(
          reopened.content.html,
          reopened.content.comments ?? [],
          {
            headerHtml: reopened.content.pageChrome?.default.headerHtml ?? '',
            footerHtml: reopened.content.pageChrome?.default.footerHtml ?? '',
          },
        ),
        secondPassIssues: reopened.compatibility?.issues ?? [],
        exportedParts,
      };
    } finally {
      forgetSourceBlob(reopened.id);
    }
  } finally {
    forgetSourceBlob(imported.id);
  }
}

export function extractDocumentIdentities(
  html: string,
  comments: readonly R0CommentIdentity[] = [],
  pageChrome: { headerHtml?: string; footerHtml?: string } = {},
): R0IdentitySnapshot {
  return {
    bookmarkNames: uniqueSorted(
      matchAll(html, /data-bookmark-name="([^"]+)"/g),
    ),
    changeAuthors: uniqueSorted(
      matchAll(html, /data-change-author="([^"]+)"/g),
    ),
    changeTexts: uniqueSorted(
      matchAll(
        html,
        /<(?:ins|del)\b[^>]*data-document-change="true"[^>]*>([\s\S]*?)<\/(?:ins|del)>/g,
      ).map((text) => text.replace(/<[^>]+>/g, '').trim()),
    ),
    commentAuthors: uniqueSorted(comments.map((comment) => comment.author)),
    commentTexts: uniqueSorted(comments.map((comment) => comment.text)),
    contentControlAliases: uniqueSorted(
      matchAll(html, /data-content-control-alias="([^"]+)"/g),
    ),
    contentControlTags: uniqueSorted(
      matchAll(html, /data-content-control-tag="([^"]+)"/g),
    ),
    contentControlTexts: uniqueSorted(
      matchAll(
        html,
        /<span\b[^>]*data-document-content-control="true"[^>]*>([\s\S]*?)<\/span>/gi,
      ).map((text) => text.replace(/<[^>]+>/g, '').trim()),
    ),
    endnoteTexts: uniqueSorted(
      matchAll(
        html,
        /<aside\b[^>]*data-document-note="true"[^>]*data-note-kind="endnote"[^>]*>([\s\S]*?)<\/aside>/gi,
      ).map((text) => text.replace(/<[^>]+>/g, '').trim()),
    ),
    footerTexts: uniqueSorted(
      [normalizePlainText(pageChrome.footerHtml ?? '')].filter(Boolean),
    ),
    footnoteTexts: uniqueSorted(
      matchAll(
        html,
        /<aside\b[^>]*data-document-note="true"[^>]*data-note-kind="footnote"[^>]*>([\s\S]*?)<\/aside>/gi,
      ).map((text) => text.replace(/<[^>]+>/g, '').trim()),
    ),
    headerTexts: uniqueSorted(
      [normalizePlainText(pageChrome.headerHtml ?? '')].filter(Boolean),
    ),
    hrefs: uniqueSorted(matchAll(html, /href="([^"]+)"/g)),
    plainText: normalizePlainText(html),
    tableCellTexts: uniqueSorted(
      matchAll(html, /<t[dh]\b[^>]*>([\s\S]*?)<\/t[dh]>/gi).map((text) =>
        text.replace(/<[^>]+>/g, '').trim(),
      ),
    ),
  };
}

export function issueCodes(
  issues: readonly WorkCompatibilityIssue[],
): string[] {
  return uniqueSorted(issues.map((issue) => issue.code));
}

export function expectIssueCodes(
  issues: readonly WorkCompatibilityIssue[],
  required: readonly string[],
): void {
  const codes = new Set(issueCodes(issues));
  for (const code of required) {
    if (!codes.has(code)) {
      throw new Error(
        `Missing compatibility issue ${code}. Observed: ${[...codes].join(', ') || '(none)'}`,
      );
    }
  }
}

function matchAll(source: string, pattern: RegExp): string[] {
  return [...source.matchAll(pattern)].map((match) => match[1] ?? '');
}

function uniqueSorted(values: readonly string[]): string[] {
  return [...new Set(values.filter(Boolean))].sort((left, right) =>
    left.localeCompare(right),
  );
}

function normalizePlainText(html: string): string {
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}
