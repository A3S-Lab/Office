import type { Editor } from '@tiptap/core';
import { type KeyboardEvent, useId, useRef } from 'react';
import { getDocumentCommandDefinition } from './document-command-catalog';
import { OfficeSelect } from './office-controls';

const documentParagraphStyles = [
  {
    value: 'paragraph',
    label: '正文',
    level: null,
    shortcut: getDocumentCommandDefinition('normalStyle').shortcut,
  },
  {
    value: 'h1',
    label: '标题 1',
    level: 1,
    shortcut: getDocumentCommandDefinition('heading1').shortcut,
  },
  {
    value: 'h2',
    label: '标题 2',
    level: 2,
    shortcut: getDocumentCommandDefinition('heading2').shortcut,
  },
  {
    value: 'h3',
    label: '标题 3',
    level: 3,
    shortcut: getDocumentCommandDefinition('heading3').shortcut,
  },
  {
    value: 'h4',
    label: '标题 4',
    level: 4,
    shortcut: undefined,
  },
  {
    value: 'h5',
    label: '标题 5',
    level: 5,
    shortcut: undefined,
  },
  {
    value: 'h6',
    label: '标题 6',
    level: 6,
    shortcut: undefined,
  },
] as const;

const documentParagraphStyleAriaKeyShortcuts = documentParagraphStyles
  .map((style) => style.shortcut?.aria)
  .filter((value): value is string => Boolean(value))
  .join(' ');

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
              title={
                style.shortcut
                  ? `${style.label}（${style.shortcut.label}）`
                  : style.label
              }
            >
              <input
                ref={(input) => {
                  inputsRef.current[index] = input;
                }}
                type="radio"
                name={groupName}
                aria-label={`应用样式：${style.label}`}
                aria-keyshortcuts={style.shortcut?.aria}
                checked={active}
                tabIndex={active ? 0 : -1}
                onChange={() => applyDocumentParagraphStyle(editor, style)}
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
        ariaKeyShortcuts={documentParagraphStyleAriaKeyShortcuts}
        className="work-document-style-select"
        value={activeStyle}
        options={documentParagraphStyles.map((style) => ({
          ...style,
          meta: style.shortcut?.label,
        }))}
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
  for (const level of [1, 2, 3, 4, 5, 6] as const) {
    if (editor.isActive('heading', { level })) {
      return `h${level}` as DocumentParagraphStyleValue;
    }
  }
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
