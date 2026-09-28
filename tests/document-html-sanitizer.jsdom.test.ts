import { describe, expect, test } from '@rstest/core';
import { importWorkDocumentFile } from '../src/internal/features/work/work-document-file-io';
import { sanitizeDocumentHtml } from '../src/internal/features/work/work-document-html-sanitizer';
import { documentPageDescriptors } from '../src/internal/features/work/work-document-pages';

const hostilePayloads: Record<string, string> = {
  svgAnchor:
    '<p>x</p><svg><a href="javascript:alert(1)"><text>click</text></a></svg>',
  svgXlink:
    '<p>x</p><svg xmlns:xlink="http://www.w3.org/1999/xlink"><a xlink:href="javascript:alert(1)"><text>c</text></a></svg>',
  formAction:
    '<p>x</p><form action="javascript:alert(1)"><button>go</button></form>',
  buttonFormaction:
    '<p>x</p><button formaction="javascript:alert(1)">go</button>',
  style: '<p>x</p><style>body{background:url(https://evil.example/t)}</style>',
  base: '<p>x</p><base href="https://evil.example/">',
  areaHref: '<p>x</p><map><area href="javascript:alert(1)"></map>',
  mathHref: '<p>x</p><math><mi href="javascript:alert(1)">m</mi></math>',
  iframeSrcdoc: '<p>x</p><iframe srcdoc="<script>alert(1)</script>"></iframe>',
  imgOnError: '<p>x<img src="x" onerror="alert(1)"></p>',
  anchorData: '<p><a href="data:text/html,<script>alert(1)</script>">d</a></p>',
};

const UNSAFE_MARKERS =
  /javascript:|<style|<base|<form|<button|<iframe|<script|onerror=|data:text\/html|evil\.example/i;

function renderedPageHtml(html: string): string {
  return documentPageDescriptors({ html } as never)
    .flatMap((page) => page.segments.map((segment) => segment.html))
    .join('');
}

describe('document HTML sanitizer', () => {
  for (const [name, payload] of Object.entries(hostilePayloads)) {
    test(`removes executable or page-hijacking markup: ${name}`, () => {
      expect(sanitizeDocumentHtml(payload)).not.toMatch(UNSAFE_MARKERS);
      expect(renderedPageHtml(payload)).not.toMatch(UNSAFE_MARKERS);
    });
  }

  test('sanitizes an imported HTML file before it becomes document content', async () => {
    const source = `<!doctype html><html><head><base href="https://evil.example/"><style>p{color:red}</style></head><body><p>Kept paragraph</p>${hostilePayloads.svgAnchor}${hostilePayloads.formAction}</body></html>`;
    const artifact = await importWorkDocumentFile(
      new File([source], 'hostile.html', { type: 'text/html' }),
      'html',
    );
    const content = artifact.content as { html: string };
    expect(content.html).toContain('Kept paragraph');
    expect(content.html).not.toMatch(UNSAFE_MARKERS);
  });

  test('removes hostile markup from rendered footnote bodies', () => {
    const html =
      '<p>Body<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="n1">1</sup></p>' +
      '<aside data-document-note="true" data-note-kind="footnote" data-note-id="n1"><p>Note text</p><svg><a href="javascript:alert(1)"><text>t</text></a></svg></aside>';
    const notes = documentPageDescriptors({ html } as never).flatMap((page) =>
      page.footnotes.map((note) => note.html),
    );
    expect(notes.join('')).toContain('Note text');
    expect(notes.join('')).not.toMatch(UNSAFE_MARKERS);
  });

  test('keeps the markup the Writer renderer emits', () => {
    const connector =
      '<div data-document-connector="true" data-connector-kind="straight" class="work-document-connector" role="img" aria-label="直线连接符" style="width: 120px; height: 40px">' +
      '<svg class="work-document-connector-svg" viewBox="0 0 100 100" preserveAspectRatio="none" focusable="false" aria-hidden="true">' +
      '<defs><marker id="c1-end" viewBox="0 0 10 10" refX="10" refY="5" markerWidth="6" markerHeight="6" orient="auto"><path d="M0 0 L10 5 L0 10 z" fill="#1f2937"></path></marker></defs>' +
      '<path d="M0 50 L100 50" stroke="#1f2937" stroke-width="2" marker-end="url(#c1-end)"></path></svg></div>';
    const equation =
      '<span data-document-equation="true"><math xmlns="http://www.w3.org/1998/Math/MathML"><mfrac><mi>a</mi><mi>b</mi></mfrac></math></span>';
    const body =
      '<p style="text-align: center" data-paragraph-id="p-1"><strong>Bold</strong> <a href="https://a3s.dev/office">link</a> <a href="#Architecture">anchor</a></p>' +
      '<table><tbody><tr><td colspan="2" style="border: 1px solid #000">cell</td></tr></tbody></table>' +
      '<p><img src="data:image/png;base64,iVBORw0KGgo=" alt="figure" width="10" height="10"></p>';
    const sanitized = sanitizeDocumentHtml(`${body}${connector}${equation}`);

    expect(sanitized).toContain('data-paragraph-id="p-1"');
    expect(sanitized).toContain('style="text-align: center"');
    expect(sanitized).toContain('href="https://a3s.dev/office"');
    expect(sanitized).toContain('href="#Architecture"');
    expect(sanitized).toContain('colspan="2"');
    expect(sanitized).toContain('src="data:image/png;base64,iVBORw0KGgo="');
    expect(sanitized).toContain('data-document-connector="true"');
    expect(sanitized).toContain('marker-end="url(#c1-end)"');
    expect(sanitized).toContain('<marker');
    expect(sanitized).toContain('<mfrac>');
    expect(sanitized).toContain('data-document-equation="true"');
  });
});
