import DOMPurify from 'dompurify';

// Writer content is HTML plus the SVG (connectors) and MathML (equations) the
// schema renders. Everything else that can navigate, submit, restyle the host
// page, or rebase its URLs has no document meaning and is removed.
const FORBIDDEN_TAGS = [
  'base',
  'button',
  'form',
  'input',
  'link',
  'meta',
  'select',
  'style',
  'textarea',
];

const FORBIDDEN_ATTRIBUTES = ['action', 'formaction'];

/**
 * Returns an allow-listed copy of Writer HTML that is safe to render with
 * `innerHTML`. URL attributes keep only safe schemes; `data:` URIs survive
 * only on media elements.
 */
export function sanitizeDocumentHtml(source: string): string {
  return DOMPurify.sanitize(source, {
    USE_PROFILES: { html: true, svg: true, mathMl: true },
    FORBID_TAGS: FORBIDDEN_TAGS,
    FORBID_ATTR: FORBIDDEN_ATTRIBUTES,
  });
}
