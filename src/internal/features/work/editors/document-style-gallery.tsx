import type { Editor } from '@tiptap/core';
import { type KeyboardEvent, useId, useRef } from 'react';
import { OfficeSelect } from './office-controls';

const documentParagraphStyles = [
  { value: 'paragraph', label: '正文', level: null },
  { value: 'h1', label: '标题 1', level: 1 },
  { value: 'h2', label: '标题 2', level: 2 },
  { value: 'h3', label: '标题 3', level: 3 },
] as const;

type DocumentParagraphStyle = (typeof documentParagraphStyles)[number];
type DocumentParagraphStyleValue = DocumentParagraphStyle['value'];

export function DocumentStyleGallery({ editor }: { editor: Editor }) {
  const groupName = useId();
  const inputsRef = useRef<Array<HTMLInputElement | null>>([]);
  const activeStyle = documentParagraphStyleValue(editor);

  const moveSelection = (
    event: KeyboardEvent<HTMLInputElement>,
    nextIndex: number,
  ) => {
    event.preventDefault();
    const normalizedIndex =
      (nextIndex + documentParagraphStyles.length) %
      documentParagraphStyles.length;
    const style = documentParagraphStyles[normalizedIndex];
    if (!style) return;
    applyDocumentParagraphStyle(editor, style);
    inputsRef.current[normalizedIndex]?.focus({ preventScroll: true });
  };

  return (
    <div className="work-document-style-tools">
      <div
        className="work-document-style-gallery"
        role="radiogroup"
        aria-label="段落样式库"
      >
        {documentParagraphStyles.map((style, index) => {
          const active = style.value === activeStyle;
          return (
            <label
              key={style.value}
              className={active ? 'active' : ''}
              data-document-style={style.value}
              title={style.label}
            >
              <input
                ref={(input) => {
                  inputsRef.current[index] = input;
                }}
                type="radio"
                name={groupName}
                aria-label={`应用样式：${style.label}`}
                checked={active}
                tabIndex={active ? 0 : -1}
                onChange={() => undefined}
                onClick={() => applyDocumentParagraphStyle(editor, style)}
                onKeyDown={(event) => {
                  if (event.key === 'ArrowRight' || event.key === 'ArrowDown') {
                    moveSelection(event, index + 1);
                  } else if (
                    event.key === 'ArrowLeft' ||
                    event.key === 'ArrowUp'
                  ) {
                    moveSelection(event, index - 1);
                  } else if (event.key === 'Home') {
                    moveSelection(event, 0);
                  } else if (event.key === 'End') {
                    moveSelection(event, documentParagraphStyles.length - 1);
                  }
                }}
              />
              <span aria-hidden="true">{style.label}</span>
            </label>
          );
        })}
      </div>
      <OfficeSelect
        ariaLabel="段落样式"
        className="work-document-style-select"
        value={activeStyle}
        options={documentParagraphStyles}
        onValueChange={(value) => {
          const style = documentParagraphStyles.find(
            (candidate) => candidate.value === value,
          );
          if (style) applyDocumentParagraphStyle(editor, style);
        }}
      />
    </div>
  );
}

function documentParagraphStyleValue(
  editor: Editor,
): DocumentParagraphStyleValue {
  if (editor.isActive('heading', { level: 1 })) return 'h1';
  if (editor.isActive('heading', { level: 2 })) return 'h2';
  if (editor.isActive('heading', { level: 3 })) return 'h3';
  return 'paragraph';
}

function applyDocumentParagraphStyle(
  editor: Editor,
  style: DocumentParagraphStyle,
): boolean {
  if (style.level === null) {
    return editor.chain().focus().setParagraph().run();
  }
  return editor.chain().focus().setHeading({ level: style.level }).run();
}
