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
  captionIds: string[];
  captionKinds: string[];
  changeAuthors: string[];
  changeTexts: string[];
  commentAuthors: string[];
  commentTexts: string[];
  contentControlAliases: string[];
  contentControlTags: string[];
  contentControlTexts: string[];
  contentControlTypes: string[];
  crossReferenceTargetIds: string[];
  endnoteTexts: string[];
  fieldInstructions: string[];
  fieldKinds: string[];
  fieldTargetNames: string[];
  footerTexts: string[];
  footnoteTexts: string[];
  headerTexts: string[];
  hrefs: string[];
  indexColumns: string[];
  indexMainEntries: string[];
  indexSubEntries: string[];
  plainText: string;
  tableCellTexts: string[];
  tocEntryTitles: string[];
  tocHyperlinks: string[];
  tocMaxLevels: string[];
  tocMinLevels: string[];
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
  const headerHtml = pageChrome.headerHtml ?? '';
  const footerHtml = pageChrome.footerHtml ?? '';
  const fieldHtml = `${html}\n${headerHtml}\n${footerHtml}`;
  return {
    bookmarkNames: uniqueSorted(
      matchAll(html, /data-bookmark-name="([^"]+)"/g),
    ),
    captionIds: uniqueSorted(matchAll(html, /data-caption-id="([^"]+)"/g)),
    captionKinds: uniqueSorted(matchAll(html, /data-caption-kind="([^"]+)"/g)),
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
    contentControlTypes: uniqueSorted(
      matchAll(html, /data-content-control-type="([^"]+)"/g),
    ),
    crossReferenceTargetIds: uniqueSorted(
      matchAll(html, /data-reference-target-id="([^"]+)"/g),
    ),
    endnoteTexts: uniqueSorted(
      matchAll(
        html,
        /<aside\b[^>]*data-document-note="true"[^>]*data-note-kind="endnote"[^>]*>([\s\S]*?)<\/aside>/gi,
      ).map((text) => text.replace(/<[^>]+>/g, '').trim()),
    ),
    fieldInstructions: uniqueSorted(
      matchAll(fieldHtml, /data-field-instruction="([^"]+)"/g).map(
        decodeHtmlAttr,
      ),
    ),
    fieldKinds: uniqueSorted(matchAll(fieldHtml, /data-field-kind="([^"]+)"/g)),
    fieldTargetNames: uniqueSorted(
      matchAll(fieldHtml, /data-field-target-name="([^"]+)"/g),
    ),
    footerTexts: uniqueSorted([normalizePlainText(footerHtml)].filter(Boolean)),
    footnoteTexts: uniqueSorted(
      matchAll(
        html,
        /<aside\b[^>]*data-document-note="true"[^>]*data-note-kind="footnote"[^>]*>([\s\S]*?)<\/aside>/gi,
      ).map((text) => text.replace(/<[^>]+>/g, '').trim()),
    ),
    headerTexts: uniqueSorted([normalizePlainText(headerHtml)].filter(Boolean)),
    hrefs: uniqueSorted(matchAll(html, /href="([^"]+)"/g)),
    indexColumns: uniqueSorted(matchAll(html, /data-index-columns="([^"]+)"/g)),
    indexMainEntries: uniqueSorted(
      matchAll(html, /data-index-main-entry="([^"]+)"/g).map(decodeHtmlAttr),
    ),
    indexSubEntries: uniqueSorted(
      matchAll(html, /data-index-sub-entry="([^"]+)"/g).map(decodeHtmlAttr),
    ),
    plainText: normalizePlainText(html),
    tableCellTexts: uniqueSorted(
      matchAll(html, /<t[dh]\b[^>]*>([\s\S]*?)<\/t[dh]>/gi).map((text) =>
        text.replace(/<[^>]+>/g, '').trim(),
      ),
    ),
    tocEntryTitles: uniqueSorted(
      matchAll(html, /data-toc-entries="([^"]+)"/g)
        .map(decodeHtmlAttr)
        .flatMap((raw) => {
          try {
            const entries = JSON.parse(raw) as Array<{ title?: string }>;
            return entries.map((entry) => entry.title?.trim() ?? '');
          } catch {
            return [];
          }
        }),
    ),
    tocHyperlinks: uniqueSorted(
      matchAll(html, /data-toc-hyperlinks="([^"]+)"/g),
    ),
    tocMaxLevels: uniqueSorted(matchAll(html, /data-toc-max-level="([^"]+)"/g)),
    tocMinLevels: uniqueSorted(matchAll(html, /data-toc-min-level="([^"]+)"/g)),
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

function decodeHtmlAttr(value: string): string {
  return value
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
}
