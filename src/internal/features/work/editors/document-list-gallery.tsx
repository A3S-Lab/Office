import type { Editor } from '@tiptap/core';
import {
  ChevronDown,
  List,
  ListOrdered,
  RotateCcw,
  Unlink,
} from 'lucide-react';
import {
  type KeyboardEvent,
  type MutableRefObject,
  useCallback,
  useRef,
  useState,
  useSyncExternalStore,
} from 'react';
import { Popover } from '../../../design-system/primitives';
import {
  type DocumentBulletListStyle,
  type DocumentOrderedListStyle,
  documentBulletListStyle,
  documentOrderedListState,
  MAX_DOCUMENT_NUMBERING_START,
} from '../work-document-lists';
import { getDocumentCommandDefinition } from './document-command-catalog';
import { OfficeNumberField } from './office-controls';
import { WorkOfficeRibbonButton } from './work-office-chrome';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import { useOfficeMessages } from './office-messages-context';

function bulletStyles(messages: OfficeMessageCatalog): ReadonlyArray<{
  value: DocumentBulletListStyle;
  label: string;
  marker: string;
}> {
  return [
    {
      value: 'disc',
      label: officeMessage(messages, 'document.list.bullet.disc'),
      marker: '●',
    },
    {
      value: 'circle',
      label: officeMessage(messages, 'document.list.bullet.circle'),
      marker: '○',
    },
    {
      value: 'square',
      label: officeMessage(messages, 'document.list.bullet.square'),
      marker: '■',
    },
  ];
}

function orderedStyles(messages: OfficeMessageCatalog): ReadonlyArray<{
  value: DocumentOrderedListStyle;
  label: string;
  markers: readonly [string, string, string];
}> {
  return [
    {
      value: 'decimal',
      label: officeMessage(messages, 'document.list.number.decimal'),
      markers: ['1.', '2.', '3.'],
    },
    {
      value: 'lower-alpha',
      label: officeMessage(messages, 'document.list.number.lowerAlpha'),
      markers: ['a.', 'b.', 'c.'],
    },
    {
      value: 'upper-alpha',
      label: officeMessage(messages, 'document.list.number.upperAlpha'),
      markers: ['A.', 'B.', 'C.'],
    },
    {
      value: 'lower-roman',
      label: officeMessage(messages, 'document.list.number.lowerRoman'),
      markers: ['i.', 'ii.', 'iii.'],
    },
    {
      value: 'upper-roman',
      label: officeMessage(messages, 'document.list.number.upperRoman'),
      markers: ['I.', 'II.', 'III.'],
    },
  ];
}

export function DocumentListGallery({ editor }: { editor: Editor }) {
  const subscribe = useCallback(
    (notify: () => void) => {
      editor.on('transaction', notify);
      return () => editor.off('transaction', notify);
    },
    [editor],
  );
  useSyncExternalStore(
    subscribe,
    () => documentListSnapshot(editor),
    () => documentListSnapshot(editor),
  );
  const bulletStyle = documentBulletListStyle(editor);
  const orderedState = documentOrderedListState(editor);
  return (
    <div className="work-document-list-tools">
      <BulletListControl editor={editor} activeStyle={bulletStyle} />
      <OrderedListControl editor={editor} activeState={orderedState} />
    </div>
  );
}

function BulletListControl({
  editor,
  activeStyle,
}: {
  editor: Editor;
  activeStyle: DocumentBulletListStyle | null;
}) {
  const messages = useOfficeMessages();
  const [open, setOpen] = useState(false);
  const optionRefs = useRef<Array<HTMLButtonElement | null>>([]);
  const activeIndex = Math.max(
    0,
    bulletStyles(messages).findIndex((style) => style.value === activeStyle),
  );
  const [focusIndex, setFocusIndex] = useState(activeIndex);
  const focusActiveOption = () => {
    setFocusIndex(activeIndex);
    requestAnimationFrame(() =>
      optionRefs.current[activeIndex]?.focus({ preventScroll: true }),
    );
  };

  const bulletShortcut = getDocumentCommandDefinition('bulletList').shortcut;
  return (
    <div
      className="work-document-list-split"
      data-active={Boolean(activeStyle)}
    >
      <WorkOfficeRibbonButton
        label={officeMessage(messages, 'document.list.bullet')}
        title={
          bulletShortcut
          ? officeMessage(messages, 'document.list.bulletWithShortcut', {
              shortcut: bulletShortcut.label,
            })
          : officeMessage(messages, 'document.list.bullet')
        }
        aria-keyshortcuts={bulletShortcut?.aria}
        displayLabel={false}
        active={Boolean(activeStyle)}
        className="work-document-list-primary"
        onMouseDown={(event) => event.preventDefault()}
        onClick={() => {
          // Keep ribbon focus — chain().focus() steals into the editor.
          if (activeStyle) editor.commands.clearDocumentList();
          else editor.commands.applyDocumentBulletList('disc');
        }}
      >
        <List size={16} />
      </WorkOfficeRibbonButton>
      <Popover
        label={officeMessage(messages, 'document.list.bulletGallery')}
        panelLabel={officeMessage(messages, 'document.list.bulletGallery')}
        panelRole="dialog"
        portal
        open={open}
        onOpenChange={(nextOpen) => {
          setOpen(nextOpen);
          if (nextOpen) focusActiveOption();
        }}
        className="work-document-list-popover"
        panelClassName="work-document-list-panel work-document-bullet-panel"
        trigger={(triggerProps, { open: popoverOpen }) => (
          <button
            {...triggerProps}
            className={`work-document-list-gallery-trigger${popoverOpen ? ' active' : ''}`}
            title={officeMessage(messages, 'document.list.bulletGallery')}
            onKeyDown={(event) => {
              if (event.key !== 'ArrowDown') return;
              event.preventDefault();
              if (!popoverOpen) event.currentTarget.click();
              else focusActiveOption();
            }}
          >
            <ChevronDown size={10} aria-hidden="true" />
          </button>
        )}
      >
        {(close) => (
          <>
            <strong className="work-document-list-panel-title">{officeMessage(messages, 'document.list.bullet')}</strong>
            <div
              className="work-document-list-options bullet-options"
              role="radiogroup"
              aria-label={officeMessage(messages, 'document.list.bulletStylesAria')}
            >
              {bulletStyles(messages).map((style, index) => (
                // biome-ignore lint/a11y/useSemanticElements: styled gallery radios; native input radios can't host glyph + keyboard grid.
                <button
                  ref={(element) => {
                    optionRefs.current[index] = element;
                  }}
                  key={style.value}
                  type="button"
                  role="radio"
                  aria-label={style.label}
                  aria-checked={style.value === activeStyle}
                  tabIndex={index === focusIndex ? 0 : -1}
                  data-list-style={style.value}
                  onFocus={() => setFocusIndex(index)}
                  onClick={() => {
                    runDocumentListMenuCommand(close, () =>
                      editor.commands.applyDocumentBulletList(style.value),
                    );
                  }}
                  onKeyDown={(event) =>
                    handleGalleryKeyDown(
                      event,
                      index,
                      bulletStyles.length,
                      3,
                      optionRefs,
                      setFocusIndex,
                      () =>
                        runDocumentListMenuCommand(close, () =>
                          editor.commands.applyDocumentBulletList(style.value),
                        ),
                    )
                  }
                >
                  <span aria-hidden="true">{style.marker}</span>
                  <small>{style.label}</small>
                </button>
              ))}
            </div>
            <button
              type="button"
              className="work-document-list-clear"
              disabled={!activeStyle}
              onClick={() =>
                runDocumentListMenuCommand(close, () =>
                  editor.commands.clearDocumentList(),
                )
              }
            >
              <Unlink size={13} aria-hidden="true" />
              {officeMessage(messages, 'document.list.clearBullet')}
            </button>
          </>
        )}
      </Popover>
    </div>
  );
}

function OrderedListControl({
  editor,
  activeState,
}: {
  editor: Editor;
  activeState: ReturnType<typeof documentOrderedListState>;
}) {
  const messages = useOfficeMessages();
  const [open, setOpen] = useState(false);
  const [startValue, setStartValue] = useState('1');
  const optionRefs = useRef<Array<HTMLButtonElement | null>>([]);
  const activeIndex = Math.max(
    0,
    orderedStyles(messages).findIndex((style) => style.value === activeState?.style),
  );
  const [focusIndex, setFocusIndex] = useState(activeIndex);
  const validStart = validStartValue(startValue);
  const committedStart = String(activeState?.start ?? 1);
  const startDirty = startValue !== committedStart;
  const focusActiveOption = () => {
    setFocusIndex(activeIndex);
    requestAnimationFrame(() =>
      optionRefs.current[activeIndex]?.focus({ preventScroll: true }),
    );
  };
  const commitStart = (raw: string) => {
    const next = validStartValue(raw);
    if (next === null) {
      setStartValue(committedStart);
      return;
    }
    setStartValue(String(next));
    if (next === (activeState?.start ?? 1)) return;
    editor.commands.setDocumentNumberingStart(next);
  };
  const applyStart = (close: () => void) => {
    if (validStart === null) return;
    if (validStart !== (activeState?.start ?? 1)) {
      if (!editor.commands.setDocumentNumberingStart(validStart)) return;
    }
    // Blur/Enter may already have committed; Apply still confirms and closes.
    close();
  };

  return (
    <div
      className="work-document-list-split"
      data-active={Boolean(activeState)}
    >
      <WorkOfficeRibbonButton
        label={officeMessage(messages, 'document.list.number')}
        title={officeMessage(messages, 'document.list.number')}
        displayLabel={false}
        active={Boolean(activeState)}
        className="work-document-list-primary"
        onMouseDown={(event) => event.preventDefault()}
        onClick={() => {
          // Keep ribbon focus — chain().focus() steals into the editor.
          if (activeState) editor.commands.clearDocumentList();
          else editor.commands.applyDocumentOrderedList('decimal');
        }}
      >
        <ListOrdered size={16} />
      </WorkOfficeRibbonButton>
      <Popover
        label={officeMessage(messages, 'document.list.numberGallery')}
        panelLabel={officeMessage(messages, 'document.list.numberGallery')}
        panelRole="dialog"
        portal
        open={open}
        onOpenChange={(nextOpen) => {
          setOpen(nextOpen);
          if (nextOpen) {
            setStartValue(String(activeState?.start ?? 1));
            focusActiveOption();
          }
        }}
        className="work-document-list-popover"
        panelClassName="work-document-list-panel work-document-numbering-panel"
        trigger={(triggerProps, { open: popoverOpen }) => (
          <button
            {...triggerProps}
            className={`work-document-list-gallery-trigger${popoverOpen ? ' active' : ''}`}
            title={officeMessage(messages, 'document.list.numberGallery')}
            onKeyDown={(event) => {
              if (event.key !== 'ArrowDown') return;
              event.preventDefault();
              if (!popoverOpen) event.currentTarget.click();
              else focusActiveOption();
            }}
          >
            <ChevronDown size={10} aria-hidden="true" />
          </button>
        )}
      >
        {(close) => (
          <>
            <strong className="work-document-list-panel-title">{officeMessage(messages, 'document.list.number')}</strong>
            <div
              className="work-document-list-options ordered-options"
              role="radiogroup"
              aria-label={officeMessage(messages, 'document.list.numberStylesAria')}
            >
              {orderedStyles(messages).map((style, index) => (
                // biome-ignore lint/a11y/useSemanticElements: styled gallery radios; native input radios can't host glyph + keyboard grid.
                <button
                  ref={(element) => {
                    optionRefs.current[index] = element;
                  }}
                  key={style.value}
                  type="button"
                  role="radio"
                  aria-label={style.label}
                  aria-checked={style.value === activeState?.style}
                  tabIndex={index === focusIndex ? 0 : -1}
                  data-list-style={style.value}
                  onFocus={() => setFocusIndex(index)}
                  onClick={() => {
                    runDocumentListMenuCommand(close, () =>
                      editor.commands.applyDocumentOrderedList(style.value),
                    );
                  }}
                  onKeyDown={(event) =>
                    handleGalleryKeyDown(
                      event,
                      index,
                      orderedStyles.length,
                      3,
                      optionRefs,
                      setFocusIndex,
                      () =>
                        runDocumentListMenuCommand(close, () =>
                          editor.commands.applyDocumentOrderedList(style.value),
                        ),
                    )
                  }
                >
                  <span className="work-document-numbering-preview">
                    {style.markers.map((marker) => (
                      <i key={marker}>{marker}</i>
                    ))}
                  </span>
                  <small>{style.label}</small>
                </button>
              ))}
            </div>
            {activeState && (
              <fieldset
                className="work-document-numbering-settings"
                data-office-escape-consumer={startDirty || undefined}
                onKeyDown={(event) => {
                  if (event.key !== 'Escape' || !startDirty) return;
                  event.preventDefault();
                  event.stopPropagation();
                  setStartValue(committedStart);
                }}
              >
                <legend className="sr-only">{officeMessage(messages, 'document.list.numberStartLegend')}</legend>
                <div className="work-document-numbering-actions">
                  <button
                    type="button"
                    disabled={activeState.start === 1}
                    onClick={() =>
                      runDocumentListMenuCommand(close, () =>
                        editor.commands.restartDocumentNumbering(),
                      )
                    }
                  >
                    <RotateCcw size={13} aria-hidden="true" />
                    {officeMessage(messages, 'document.list.restart')}
                  </button>
                  <button
                    type="button"
                    disabled={!editor.can().continueDocumentNumbering()}
                    onClick={() =>
                      runDocumentListMenuCommand(close, () =>
                        editor.commands.continueDocumentNumbering(),
                      )
                    }
                  >
                    {officeMessage(messages, 'document.list.continue')}
                  </button>
                </div>
                <form
                  className="work-document-numbering-start"
                  aria-label={officeMessage(messages, 'document.list.startSettingsAria')}
                  onSubmit={(event) => {
                    event.preventDefault();
                    applyStart(close);
                  }}
                >
                  <span>{officeMessage(messages, 'document.list.startNumber')}</span>
                  <OfficeNumberField
                    ariaLabel={officeMessage(messages, 'document.list.startNumberAria')}
                    value={startValue}
                    min={1}
                    max={MAX_DOCUMENT_NUMBERING_START}
                    step={1}
                    escapeConsumer={startDirty}
                    onValueChange={setStartValue}
                    onCommit={commitStart}
                    onCancel={
                      startDirty
                        ? () => setStartValue(committedStart)
                        : undefined
                    }
                  />
                  <button type="submit" disabled={validStart === null}>
                    {officeMessage(messages, 'document.list.applyStart')}
                  </button>
                </form>
              </fieldset>
            )}
            <button
              type="button"
              className="work-document-list-clear"
              disabled={!activeState}
              onClick={() =>
                runDocumentListMenuCommand(close, () =>
                  editor.commands.clearDocumentList(),
                )
              }
            >
              <Unlink size={13} aria-hidden="true" />
              {officeMessage(messages, 'document.list.clearNumber')}
            </button>
          </>
        )}
      </Popover>
    </div>
  );
}

function runDocumentListMenuCommand(
  close: () => void,
  command: () => boolean,
): boolean {
  const handled = command();
  if (!handled) return false;
  // Keep ribbon keyboard focus on the gallery trigger via Popover restore.
  // Forcing editor focus after choose breaks L2 disclosure loops.
  close();
  return true;
}

function handleGalleryKeyDown(
  event: KeyboardEvent<HTMLButtonElement>,
  index: number,
  count: number,
  columns: number,
  refs: MutableRefObject<Array<HTMLButtonElement | null>>,
  setFocusIndex: (index: number) => void,
  select: () => void,
): void {
  if (event.key === 'Enter' || event.key === ' ') {
    event.preventDefault();
    select();
    return;
  }
  let next = index;
  if (event.key === 'ArrowRight') next = (index + 1) % count;
  else if (event.key === 'ArrowLeft') next = (index - 1 + count) % count;
  else if (event.key === 'ArrowDown')
    next = Math.min(count - 1, index + columns);
  else if (event.key === 'ArrowUp') next = Math.max(0, index - columns);
  else if (event.key === 'Home') next = 0;
  else if (event.key === 'End') next = count - 1;
  else return;
  event.preventDefault();
  setFocusIndex(next);
  refs.current[next]?.focus({ preventScroll: true });
}

function validStartValue(value: string): number | null {
  const number = Number(value);
  return Number.isSafeInteger(number) &&
    number >= 1 &&
    number <= MAX_DOCUMENT_NUMBERING_START
    ? number
    : null;
}

function documentListSnapshot(editor: Editor): string {
  const bullet = documentBulletListStyle(editor);
  const ordered = documentOrderedListState(editor);
  return `${bullet ?? ''}:${ordered?.style ?? ''}:${ordered?.start ?? ''}`;
}
