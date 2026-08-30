import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { Editor } from '@tiptap/core';
import {
  newPresentationElement,
  newSlide,
} from '../src/internal/features/work/editors/presentation-editor-operations';
import { SlideElementTextPreview } from '../src/internal/features/work/editors/presentation-slide-canvas';
import {
  applyPresentationTextFormatting,
  createPresentationTextEditorExtensions,
  PresentationTextEditor,
  presentationTextElementHtml,
  presentationTextValue,
} from '../src/internal/features/work/editors/presentation-text-editor';
import type { WorkSlideElement } from '../src/internal/features/work/work-types';

test('round-trips imported presentation text runs through TipTap', () => {
  const element = textElement({
    bold: true,
    fontFamily: 'Arial',
    text: 'A3S Office',
    textRuns: [
      {
        text: 'A3S ',
        bold: true,
        color: '#175cd3',
        fontFamily: 'Arial',
        fontSize: 28,
      },
      {
        text: 'Office',
        bold: false,
        color: '#172033',
        fontFamily: 'Times New Roman',
        fontSize: 22,
        href: 'https://a3s.dev',
        italic: true,
        underline: true,
      },
    ],
  });
  const editor = new Editor({
    extensions: createPresentationTextEditorExtensions(),
    content: presentationTextElementHtml(element),
  });

  expect(presentationTextValue(editor, element)).toEqual({
    text: 'A3S Office',
    textRuns: [
      {
        text: 'A3S ',
        bold: true,
        color: '#175cd3',
        fontFamily: 'Arial',
        fontSize: 28,
        italic: false,
        underline: false,
      },
      {
        text: 'Office',
        bold: false,
        color: '#172033',
        fontFamily: 'Times New Roman',
        fontSize: 22,
        href: 'https://a3s.dev',
        italic: true,
        underline: true,
      },
    ],
  });

  editor.destroy();
});

test('applies presentation formatting to the active TipTap text selection', () => {
  const element = textElement({ text: 'A3S Office' });
  const editor = new Editor({
    extensions: createPresentationTextEditorExtensions(),
    content: presentationTextElementHtml(element),
  });

  editor.commands.setTextSelection({ from: 1, to: 4 });
  expect(
    applyPresentationTextFormatting(editor, {
      bold: true,
      color: '#d92d20',
      fontSize: 30,
      href: 'https://a3s.dev',
    }),
  ).toBe(true);

  expect(presentationTextValue(editor, element)).toEqual({
    text: 'A3S Office',
    textRuns: [
      {
        text: 'A3S',
        bold: true,
        color: '#d92d20',
        fontFamily: 'Aptos',
        fontSize: 30,
        href: 'https://a3s.dev',
        italic: false,
        underline: false,
      },
      {
        text: ' Office',
        bold: false,
        color: '#172033',
        fontFamily: 'Aptos',
        fontSize: 24,
        italic: false,
        underline: false,
      },
    ],
  });
  expect(applyPresentationTextFormatting(editor, { align: 'right' })).toBe(
    false,
  );

  editor.destroy();
});

test('mounts one TipTap-backed presentation text surface', async () => {
  render(
    <PresentationTextEditor
      element={textElement({ text: 'Editable slide text' })}
      onChange={() => undefined}
    />,
  );

  expect(
    await screen.findByRole('textbox', { name: '幻灯片文本' }),
  ).toHaveAttribute('data-presentation-text-engine', 'tiptap');
});

test('publishes only the committed Chinese IME text', async () => {
  let editor: Editor | null = null;
  const published: string[] = [];
  render(
    <PresentationTextEditor
      element={textElement()}
      onChange={(value) => published.push(value.text)}
      onEditorChange={(current) => {
        editor = current;
      }}
    />,
  );

  const textbox = await screen.findByRole('textbox', {
    name: '幻灯片文本',
  });
  await waitFor(() => expect(editor).not.toBeNull());
  const currentEditor = editor as Editor;

  fireEvent.compositionStart(textbox, { data: 'ni' });
  expect(currentEditor.view.composing).toBe(true);
  currentEditor.commands.insertContent('ni');
  expect(published).toEqual([]);

  currentEditor.chain().selectAll().insertContent('你').run();
  fireEvent.compositionEnd(textbox, { data: '你' });
  await waitFor(() => expect(published).toEqual(['你']));
});

test('returns Escape from text editing to its slide object', async () => {
  let exits = 0;
  render(
    <fieldset
      data-slide-element-origin="slide"
      data-testid="slide-object"
      // biome-ignore lint/a11y/noNoninteractiveTabindex: Mirrors the keyboard-selectable slide object wrapper.
      tabIndex={0}
    >
      <legend>Slide object</legend>
      <PresentationTextEditor
        element={textElement({ text: 'Editable slide text' })}
        onChange={() => undefined}
        onExitEditing={() => {
          exits += 1;
        }}
      />
    </fieldset>,
  );

  const textbox = await screen.findByRole('textbox', {
    name: '幻灯片文本',
  });
  fireEvent.keyDown(textbox, { key: 'Escape' });
  expect(exits).toBe(1);
  await waitFor(() => expect(screen.getByTestId('slide-object')).toHaveFocus());
});

test('keeps form control focus while formatting a text selection', async () => {
  let editor: Editor | null = null;
  render(
    <>
      <input aria-label="演示字号输入" />
      <PresentationTextEditor
        element={textElement({ text: 'Editable slide text' })}
        onChange={() => undefined}
        onEditorChange={(current) => {
          editor = current;
        }}
      />
    </>,
  );

  await screen.findByRole('textbox', { name: '幻灯片文本' });
  const sizeInput = screen.getByRole('textbox', { name: '演示字号输入' });
  sizeInput.focus();
  expect(sizeInput).toHaveFocus();
  expect(editor).not.toBeNull();

  const currentEditor = editor as Editor;
  currentEditor.commands.setTextSelection({ from: 1, to: 9 });
  expect(
    applyPresentationTextFormatting(
      currentEditor,
      { fontSize: 32 },
      { restoreFocus: false },
    ),
  ).toBe(true);

  expect(sizeInput).toHaveFocus();
  expect(
    presentationTextValue(currentEditor, textElement()).textRuns?.[0],
  ).toMatchObject({
    fontSize: 32,
    text: 'Editable',
  });
});

test('keeps presentation prompts out of stored slide text', async () => {
  const placeholder = textElement({
    text: '',
    textRuns: [{ text: '' }],
    placeholder: {
      key: 'title',
      type: 'title',
      prompt: '单击添加标题',
    },
  });

  const view = render(
    <PresentationTextEditor element={placeholder} onChange={() => undefined} />,
  );
  const textbox = await screen.findByRole('textbox', { name: '幻灯片文本' });
  expect(textbox).toHaveTextContent('');
  expect(textbox.querySelector('p')).toHaveAttribute(
    'data-placeholder',
    '单击添加标题',
  );

  view.rerender(
    <SlideElementTextPreview element={placeholder} showPlaceholder />,
  );
  expect(screen.getByText('单击添加标题')).toBeInTheDocument();

  expect(newSlide(2).elements[0]).toMatchObject({
    text: '',
    placeholder: { type: 'title', prompt: '单击添加标题' },
  });
  expect(newPresentationElement('text')).toMatchObject({ text: '' });
});

function textElement(patch: Partial<WorkSlideElement> = {}): WorkSlideElement {
  return {
    id: 'text-1',
    type: 'text',
    x: 10,
    y: 10,
    width: 60,
    height: 20,
    text: '',
    fontSize: 24,
    color: '#172033',
    fill: 'transparent',
    bold: false,
    align: 'left',
    fontFamily: 'Aptos',
    ...patch,
  };
}
