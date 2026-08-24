import { afterEach, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { DocumentProofingDialog } from '../src/internal/features/work/editors/document-proofing-dialog';
import {
  applyDocumentProofingDialogPatch,
  documentProofingDialogSource,
} from '../src/internal/features/work/editors/document-proofing-dialog-model';
import {
  documentProofingDomAttributes,
  parseDocumentProofingLanguages,
} from '../src/internal/features/work/work-document-proofing';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  createWorkArtifact,
  WORK_TEMPLATES,
} from '../src/internal/features/work/work-templates';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

test('reports independently mixed proofing language slots and proofing state', () => {
  editor = createEditor(
    `<p>${proofingSpan(
      { latin: 'en-US', eastAsia: 'zh-CN', bidi: 'ar-SA' },
      false,
      'first',
    )}${proofingSpan(
      { latin: 'fr-FR', eastAsia: 'zh-CN', bidi: 'ar-SA' },
      true,
      'second',
    )}</p>`,
  );
  editor.commands.selectAll();

  expect(documentProofingDialogSource(editor)).toMatchObject({
    latin: { mixed: true, value: null },
    eastAsia: { mixed: false, value: 'zh-CN' },
    bidi: { mixed: false, value: 'ar-SA' },
    noProof: { mixed: true, value: null },
  });
});

test('patches one language slot across a mixed selection without clobbering other text styles', () => {
  editor = createEditor(
    `<p>${proofingSpan(
      { latin: 'en-US', eastAsia: 'zh-CN', bidi: 'ar-SA' },
      true,
      'first',
      'color: #c00000; font-size: 12pt',
    )}${proofingSpan(
      { latin: 'fr-FR', eastAsia: 'ko-KR', bidi: 'he-IL' },
      false,
      'second',
      'color: #0070c0; font-size: 14pt',
    )}</p>`,
  );
  editor.commands.selectAll();
  expect(
    applyDocumentProofingDialogPatch(
      editor,
      { from: editor.state.selection.from, to: editor.state.selection.to },
      {
        languages: { eastAsia: 'ja-JP' },
        noProof: false,
      },
    ),
  ).toBe(true);

  const spans = proofingSpans(editor);
  expect(spans.map(({ languages }) => languages)).toEqual([
    { latin: 'en-US', eastAsia: 'ja-JP', bidi: 'ar-SA' },
    { latin: 'fr-FR', eastAsia: 'ja-JP', bidi: 'he-IL' },
  ]);
  expect(spans.map(({ noProof }) => noProof)).toEqual(['false', 'false']);
  expect(spans.map(({ color }) => color)).toEqual(['#c00000', '#0070c0']);
  expect(spans.map(({ fontSize }) => fontSize)).toEqual(['12pt', '14pt']);

  expect(editor.commands.undo()).toBe(true);
  expect(proofingSpans(editor).map(({ languages }) => languages)).toEqual([
    { latin: 'en-US', eastAsia: 'zh-CN', bidi: 'ar-SA' },
    { latin: 'fr-FR', eastAsia: 'ko-KR', bidi: 'he-IL' },
  ]);
  expect(editor.commands.undo()).toBe(false);
});

test('keeps untouched mixed values and applies custom tags through one accessible dialog', () => {
  const source = {
    latin: { mixed: true, value: null },
    eastAsia: { mixed: false, value: 'zh-CN' },
    bidi: { mixed: false, value: null },
    noProof: { mixed: true, value: null },
    selectedCharacters: 12,
  } as const;
  const patches: unknown[] = [];

  render(
    <DocumentProofingDialog
      source={source}
      restoreFocusTarget={() => null}
      onApply={(patch) => {
        patches.push(patch);
        return true;
      }}
      onClose={() => undefined}
    />,
  );

  expect(screen.getByText(/不同的拉丁文字校对语言/)).toBeVisible();
  expect(screen.getByRole('button', { name: '应用' })).toBeDisabled();
  fireEvent.change(screen.getByRole('combobox', { name: '东亚文字校对语言' }), {
    target: { value: 'ja-JP' },
  });
  fireEvent.click(screen.getByRole('combobox', { name: '校对行为' }));
  fireEvent.click(screen.getByRole('option', { name: '不检查拼写或语法' }));
  fireEvent.click(screen.getByRole('button', { name: '应用' }));

  expect(patches).toEqual([
    {
      languages: { eastAsia: 'ja-JP' },
      noProof: true,
    },
  ]);
});

test('clears a direct language slot through Follow Style', () => {
  const patches: unknown[] = [];
  render(
    <DocumentProofingDialog
      source={{
        latin: { mixed: false, value: 'en-US' },
        eastAsia: { mixed: false, value: 'zh-CN' },
        bidi: { mixed: false, value: 'ar-SA' },
        noProof: { mixed: false, value: false },
        selectedCharacters: 4,
      }}
      restoreFocusTarget={() => null}
      onApply={(patch) => {
        patches.push(patch);
        return true;
      }}
      onClose={() => undefined}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '东亚文字跟随样式' }));
  fireEvent.click(screen.getByRole('button', { name: '应用' }));

  expect(patches).toEqual([{ languages: { eastAsia: null } }]);
});

test('publishes a Playground template with Latin, East Asian, bidi, and excluded proofing examples', () => {
  expect(WORK_TEMPLATES).toContainEqual(
    expect.objectContaining({
      id: 'proofing-languages',
      kind: 'document',
      name: '校对语言',
    }),
  );
  const artifact = createWorkArtifact('proofing-languages');
  expect(artifact.title).toBe('校对语言示例');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected the proofing template to create a document.');
  }
  expect(artifact.content.html).toContain('"latin":"en-US"');
  expect(artifact.content.html).toContain('"eastAsia":"zh-CN"');
  expect(artifact.content.html).toContain('"bidi":"ar-SA"');
  expect(artifact.content.html).toContain("data-office-no-proof='true'");
  expect(artifact.content.html).toContain(
    "data-office-script-font-slot='eastAsia'",
  );
  expect(artifact.content.html).toContain(
    "data-office-script-font-slot='complexScript'",
  );
  expect(artifact.content.html).toContain('dir="rtl"');

  editor = createEditor(artifact.content.html);
  editor.commands.setContent(editor.getHTML());
  const reopened = new DOMParser().parseFromString(
    editor.getHTML(),
    'text/html',
  );
  const languageFor = (text: string) =>
    Array.from(reopened.querySelectorAll<HTMLElement>('span')).find(
      (span) => span.textContent === text,
    )?.lang;
  expect(languageFor('English proofing language')).toBe('en-US');
  expect(languageFor('简体中文校对语言')).toBe('zh-CN');
  expect(languageFor('لغة التدقيق العربية')).toBe('ar-SA');
});

function createEditor(content: string): Editor {
  return new Editor({
    extensions: createWorkDocumentExtensions(),
    content,
  });
}

function proofingSpan(
  languages: Record<string, string>,
  noProof: boolean,
  text: string,
  style?: string,
): string {
  const attributes = documentProofingDomAttributes(languages, noProof);
  return `<span ${Object.entries(attributes)
    .map(([name, value]) => `${name}='${value}'`)
    .join(' ')}${style ? ` style="${style}"` : ''}>${text}</span>`;
}

function proofingSpans(currentEditor: Editor): Array<{
  color: string;
  fontSize: string;
  languages: ReturnType<typeof parseDocumentProofingLanguages>;
  noProof: string | undefined;
}> {
  const document = new DOMParser().parseFromString(
    currentEditor.getHTML(),
    'text/html',
  );
  return Array.from(
    document.querySelectorAll<HTMLElement>('[data-office-proofing-languages]'),
    (span) => ({
      color: span.style.color,
      fontSize: span.style.fontSize,
      languages: parseDocumentProofingLanguages(
        span.dataset.officeProofingLanguages,
      ),
      noProof: span.dataset.officeNoProof,
    }),
  );
}
