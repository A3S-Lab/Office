import bash from '@shikijs/langs/bash';
import css from '@shikijs/langs/css';
import diff from '@shikijs/langs/diff';
import go from '@shikijs/langs/go';
import html from '@shikijs/langs/html';
import javascript from '@shikijs/langs/javascript';
import json from '@shikijs/langs/json';
import jsx from '@shikijs/langs/jsx';
import markdown from '@shikijs/langs/markdown';
import python from '@shikijs/langs/python';
import rust from '@shikijs/langs/rust';
import sql from '@shikijs/langs/sql';
import toml from '@shikijs/langs/toml';
import tsx from '@shikijs/langs/tsx';
import typescript from '@shikijs/langs/typescript';
import yaml from '@shikijs/langs/yaml';
import { renderMermaidSVG } from 'beautiful-mermaid';
import {
  type BundledLanguage,
  bundledLanguagesInfo,
  createHighlighter,
  type Highlighter,
  type ThemedToken,
} from 'shiki';
import { createJavaScriptRegexEngine } from 'shiki/engine/javascript';

const LIGHT = 'github-light';
const DARK = 'github-dark';
const DIAGRAM_FONT =
  '"SFMono-Regular", Consolas, "Liberation Mono", Menlo, monospace';

const READY_LANGUAGES = [
  bash,
  css,
  diff,
  go,
  html,
  javascript,
  json,
  jsx,
  markdown,
  python,
  rust,
  sql,
  toml,
  tsx,
  typescript,
  yaml,
];

const LANGUAGE_ID = new Map<string, BundledLanguage>();
for (const info of bundledLanguagesInfo) {
  LANGUAGE_ID.set(info.id, info.id);
  for (const alias of info.aliases ?? []) LANGUAGE_ID.set(alias, info.id);
}

const PLAIN = new Set(['', 'text', 'plaintext', 'plain', 'txt']);

export interface MarkdownCodeToken {
  from: number;
  to: number;
  light: string;
  dark: string;
}

export function resolveMarkdownCodeLanguage(
  language: string | null | undefined,
): BundledLanguage | 'text' | 'mermaid' {
  const key = language?.trim().toLowerCase() ?? '';
  if (key === 'mermaid') return 'mermaid';
  if (PLAIN.has(key)) return 'text';
  return LANGUAGE_ID.get(key) ?? 'text';
}

let highlighter: Promise<Highlighter> | undefined;

function highlighterPromise(): Promise<Highlighter> {
  highlighter ??= createHighlighter({
    engine: createJavaScriptRegexEngine({ forgiving: true }),
    langs: READY_LANGUAGES,
    themes: [LIGHT, DARK],
  });
  return highlighter;
}

function tokenColors(
  token: ThemedToken,
): { light: string; dark: string } | undefined {
  const style = token.htmlStyle ?? {};
  const light = style['--shiki-light'] || style.color || token.color;
  const dark = style['--shiki-dark'] || light;
  if (!light || !dark) return undefined;
  return { dark, light };
}

/** Shiki token ranges for a fenced block. Plaintext returns no tokens. */
export async function highlightMarkdownCode(
  source: string,
  language: string | null | undefined,
): Promise<MarkdownCodeToken[]> {
  const lang = resolveMarkdownCodeLanguage(language);
  if (lang === 'text' || lang === 'mermaid' || !source) return [];
  try {
    const hl = await highlighterPromise();
    if (!hl.getLoadedLanguages().includes(lang)) {
      await hl.loadLanguage(lang);
    }
    const result = hl.codeToTokens(source, {
      defaultColor: false,
      lang,
      themes: { dark: DARK, light: LIGHT },
    });
    const tokens: MarkdownCodeToken[] = [];
    for (const line of result.tokens) {
      for (const token of line) {
        if (!token.content) continue;
        const colors = tokenColors(token);
        if (!colors || token.offset === undefined) continue;
        tokens.push({
          dark: colors.dark,
          from: token.offset,
          light: colors.light,
          to: token.offset + token.content.length,
        });
      }
    }
    return tokens;
  } catch {
    return [];
  }
}

function drawnDiagram(svg: string): boolean {
  return /<(?:path|rect|circle|ellipse|polygon|polyline|text|line)\b/i.test(
    svg,
  );
}

/** Keep designed size and product theme variables. Drop the Google Fonts import. */
export function presentMarkdownDiagram(svg: string): string {
  const withoutImports = svg.replace(/@import\s+url\([^)]*\)\s*;?/gi, '');
  const style = /\sstyle="([^"]*)"/i.exec(withoutImports)?.[1] ?? '';
  const keepTheme = /--(?:bg|fg)\s*:/.test(style);
  const viewBox = /viewBox="([\d.\s-]+)"/i.exec(withoutImports);
  const parts = viewBox?.[1].trim().split(/\s+/u).map(Number) ?? [];
  const width = parts.length === 4 ? parts[2] : undefined;
  const height = parts.length === 4 ? parts[3] : undefined;
  let next = keepTheme
    ? withoutImports
    : withoutImports.replace(/\sstyle="[^"]*"/i, '');
  next = next.replace(/\swidth="[^"]*"/i, '').replace(/\sheight="[^"]*"/i, '');
  if (width && height && Number.isFinite(width) && Number.isFinite(height)) {
    next = next.replace(/<svg\b/i, `<svg width="${width}" height="${height}"`);
  }
  return next;
}

/** beautiful-mermaid SVG, or undefined when the diagram type is unsupported. */
export function renderMarkdownDiagram(source: string): string | undefined {
  const body = source.trim();
  if (!body) return undefined;
  try {
    const svg = renderMermaidSVG(body, {
      accent: 'var(--a3s-ink)',
      bg: 'transparent',
      border: 'var(--a3s-line-strong)',
      fg: 'var(--a3s-ink)',
      font: DIAGRAM_FONT,
      line: 'var(--a3s-line-strong)',
      muted: 'var(--a3s-muted)',
      surface: 'var(--a3s-panel)',
      transparent: true,
    });
    const presented = presentMarkdownDiagram(svg);
    return drawnDiagram(presented) ? presented : undefined;
  } catch {
    return undefined;
  }
}
