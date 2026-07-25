export const DOCUMENT_LINK_VALIDATION_MESSAGE =
  '请输入完整的 http、https、mailto 或 # 文档内地址。';

export function normalizeDocumentHref(value: string): string | null {
  const href = value.trim();
  if (/^#[^\s#]+$/u.test(href)) return href;
  if (/^mailto:[^\s@]+@[^\s@]+$/iu.test(href)) return href;
  if (!/^https?:/iu.test(href)) return null;

  try {
    const url = new URL(href);
    return url.hostname &&
      (url.protocol === 'http:' || url.protocol === 'https:')
      ? href
      : null;
  } catch {
    return null;
  }
}
