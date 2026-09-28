import type { Editor } from '@tiptap/core';
import { type KeyboardEvent, useId, useMemo, useRef } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import { getDocumentCommandDefinition } from './document-command-catalog';
import { documentCommandLabel } from './document-command-i18n';
import { OfficeSelect } from './office-controls';
import { useOfficeMessages } from './office-messages-context';

type DocumentParagraphStyleValue =
  | 'paragraph'
  | 'h1'
  | 'h2'
  | 'h3'
  | 'h4'
  | 'h5'
  | 'h6';

type DocumentParagraphStyle = {
  value: DocumentParagraphStyleValue;
  label: string;
  level: 1 | 2 | 3 | 4 | 5 | 6 | null;
  shortcut: ReturnType<typeof getDocumentCommandDefinition>['shortcut'];
};

function buildDocumentParagraphStyles(
  messages: OfficeMessageCatalog,
): DocumentParagraphStyle[] {
  return [
    {
      value: 'paragraph',
      label: documentCommandLabel('normalStyle', messages),
      level: null,
      shortcut: getDocumentCommandDefinition('normalStyle').shortcut,
    },
    {
      value: 'h1',
      label: documentCommandLabel('heading1', messages),
      level: 1,
      shortcut: getDocumentCommandDefinition('heading1').shortcut,
    },
    {
      value: 'h2',
      label: documentCommandLabel('heading2', messages),
      level: 2,
      shortcut: getDocumentCommandDefinition('heading2').shortcut,
    },
    {
      value: 'h3',
      label: documentCommandLabel('heading3', messages),
      level: 3,
      shortcut: getDocumentCommandDefinition('heading3').shortcut,
    },
    {
      value: 'h4',
      label: messages['document.style.heading4'],
      level: 4,
      shortcut: undefined,
    },
    {
      value: 'h5',
      label: messages['document.style.heading5'],
      level: 5,
      shortcut: undefined,
    },
    {
      value: 'h6',
      label: messages['document.style.heading6'],
      level: 6,
      shortcut: undefined,
    },
  ];
}

export function DocumentStyleGallery({ editor }: { editor: Editor }) {
  const messages = useOfficeMessages();
  const styles = useMemo(
    () => buildDocumentParagraphStyles(messages),
    [messages],
  );
  const styleAriaKeyShortcuts = useMemo(
    () =>
      styles
        .map((style) => style.shortcut?.aria)
        .filter((value): value is string => Boolean(value))
        .join(' '),
    [styles],
  );
  const groupName = useId();
  const inputsRef = useRef<Array<HTMLInputElement | null>>([]);
  const activeStyle = documentParagraphStyleValue(editor);

  const moveSelection = (
    event: KeyboardEvent<HTMLInputElement>,
    nextIndex: number,
  ) => {
    event.preventDefault();
    const normalizedIndex = (nextIndex + styles.length) % styles.length;
    const style = styles[normalizedIndex];
    if (!style) return;
    applyDocumentParagraphStyle(editor, style);
    inputsRef.current[normalizedIndex]?.focus({ preventScroll: true });
  };

  return (
    <div className="work-document-style-tools">
      <div
        className="work-document-style-gallery"
        role="radiogroup"
        aria-label={messages['document.style.gallery']}
      >
        {styles.map((style, index) => {
          const active = style.value === activeStyle;
          return (
            <label
              key={style.value}
              className={active ? 'active' : ''}
              data-document-style={style.value}
              title={
                style.shortcut
                  ? officeMessage(messages, 'document.command.withShortcut', {
                      label: style.label,
                      shortcut: style.shortcut.label,
                    })
                  : style.label
              }
            >
              <input
                ref={(input) => {
                  inputsRef.current[index] = input;
                }}
                type="radio"
                name={groupName}
                aria-label={officeMessage(messages, 'document.style.apply', {
                  label: style.label,
                })}
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
                    moveSelection(event, styles.length - 1);
                  }
                }}
              />
              <span aria-hidden="true">{style.label}</span>
            </label>
          );
        })}
      </div>
      <OfficeSelect
        ariaLabel={messages['document.style.select']}
        ariaKeyShortcuts={styleAriaKeyShortcuts}
        className="work-document-style-select"
        value={activeStyle}
        options={styles.map((style) => ({
          ...style,
          meta: style.shortcut?.label,
        }))}
        onValueChange={(value) => {
          const style = styles.find((candidate) => candidate.value === value);
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
  // Keep ribbon focus on the style combobox / radiogroup. chain().focus()
  // schedules into the editor and breaks L2 loops after a Popover pick.
  if (style.level === null) {
    return editor.commands.setParagraph();
  }
  return editor.commands.setHeading({ level: style.level });
}
