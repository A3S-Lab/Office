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
  MAX_DOCUMENT_NUMBERING_START,
  documentBulletListStyle,
  documentOrderedListState,
} from '../work-document-lists';
import { OfficeNumberField } from './office-controls';
import { WorkOfficeRibbonButton } from './work-office-chrome';

const bulletStyles: ReadonlyArray<{
  value: DocumentBulletListStyle;
  label: string;
  marker: string;
}> = [
  { value: 'disc', label: '实心圆点', marker: '●' },
  { value: 'circle', label: '空心圆点', marker: '○' },
  { value: 'square', label: '方块', marker: '■' },
];

const orderedStyles: ReadonlyArray<{
  value: DocumentOrderedListStyle;
  label: string;
  markers: readonly [string, string, string];
}> = [
  { value: 'decimal', label: '数字', markers: ['1.', '2.', '3.'] },
  { value: 'lower-alpha', label: '小写字母', markers: ['a.', 'b.', 'c.'] },
  { value: 'upper-alpha', label: '大写字母', markers: ['A.', 'B.', 'C.'] },
  {
    value: 'lower-roman',
    label: '小写罗马数字',
    markers: ['i.', 'ii.', 'iii.'],
  },
  {
    value: 'upper-roman',
    label: '大写罗马数字',
    markers: ['I.', 'II.', 'III.'],
  },
];

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
  const [open, setOpen] = useState(false);
  const optionRefs = useRef<Array<HTMLButtonElement | null>>([]);
  const activeIndex = Math.max(
    0,
    bulletStyles.findIndex((style) => style.value === activeStyle),
  );
  const [focusIndex, setFocusIndex] = useState(activeIndex);
  const focusActiveOption = () => {
    setFocusIndex(activeIndex);
    requestAnimationFrame(() =>
      optionRefs.current[activeIndex]?.focus({ preventScroll: true }),
    );
  };

  return (
    <div
      className="work-document-list-split"
      data-active={Boolean(activeStyle)}
    >
      <WorkOfficeRibbonButton
        label="项目符号"
        title="项目符号"
        displayLabel={false}
        active={Boolean(activeStyle)}
        className="work-document-list-primary"
        onClick={() => {
          if (activeStyle) editor.chain().focus().clearDocumentList().run();
          else editor.chain().focus().applyDocumentBulletList('disc').run();
        }}
      >
        <List size={16} />
      </WorkOfficeRibbonButton>
      <Popover
        label="项目符号库"
        panelLabel="项目符号库"
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
            title="项目符号库"
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
            <strong className="work-document-list-panel-title">项目符号</strong>
            <div
              className="work-document-list-options bullet-options"
              role="menu"
              aria-label="项目符号样式"
            >
              {bulletStyles.map((style, index) => (
                <button
                  ref={(element) => {
                    optionRefs.current[index] = element;
                  }}
                  key={style.value}
                  type="button"
                  role="menuitemradio"
                  aria-label={style.label}
                  aria-checked={style.value === activeStyle}
                  tabIndex={index === focusIndex ? 0 : -1}
                  data-list-style={style.value}
                  onFocus={() => setFocusIndex(index)}
                  onClick={() => {
                    runDocumentListMenuCommand(editor, close, () =>
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
                        runDocumentListMenuCommand(editor, close, () =>
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
                runDocumentListMenuCommand(editor, close, () =>
                  editor.commands.clearDocumentList(),
                )
              }
            >
              <Unlink size={13} aria-hidden="true" />
              清除项目符号
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
  const [open, setOpen] = useState(false);
  const [startValue, setStartValue] = useState('1');
  const optionRefs = useRef<Array<HTMLButtonElement | null>>([]);
  const activeIndex = Math.max(
    0,
    orderedStyles.findIndex((style) => style.value === activeState?.style),
  );
  const [focusIndex, setFocusIndex] = useState(activeIndex);
  const validStart = validStartValue(startValue);
  const focusActiveOption = () => {
    setFocusIndex(activeIndex);
    requestAnimationFrame(() =>
      optionRefs.current[activeIndex]?.focus({ preventScroll: true }),
    );
  };
  const applyStart = (close: () => void) => {
    if (validStart === null) return;
    runDocumentListMenuCommand(editor, close, () =>
      editor.commands.setDocumentNumberingStart(validStart),
    );
  };

  return (
    <div
      className="work-document-list-split"
      data-active={Boolean(activeState)}
    >
      <WorkOfficeRibbonButton
        label="编号"
        title="编号"
        displayLabel={false}
        active={Boolean(activeState)}
        className="work-document-list-primary"
        onClick={() => {
          if (activeState) editor.chain().focus().clearDocumentList().run();
          else editor.chain().focus().applyDocumentOrderedList('decimal').run();
        }}
      >
        <ListOrdered size={16} />
      </WorkOfficeRibbonButton>
      <Popover
        label="编号库"
        panelLabel="编号库"
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
            title="编号库"
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
            <strong className="work-document-list-panel-title">编号</strong>
            <div
              className="work-document-list-options ordered-options"
              role="menu"
              aria-label="编号样式"
            >
              {orderedStyles.map((style, index) => (
                <button
                  ref={(element) => {
                    optionRefs.current[index] = element;
                  }}
                  key={style.value}
                  type="button"
                  role="menuitemradio"
                  aria-label={style.label}
                  aria-checked={style.value === activeState?.style}
                  tabIndex={index === focusIndex ? 0 : -1}
                  data-list-style={style.value}
                  onFocus={() => setFocusIndex(index)}
                  onClick={() => {
                    runDocumentListMenuCommand(editor, close, () =>
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
                        runDocumentListMenuCommand(editor, close, () =>
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
              <div className="work-document-numbering-settings">
                <div className="work-document-numbering-actions">
                  <button
                    type="button"
                    disabled={activeState.start === 1}
                    onClick={() =>
                      runDocumentListMenuCommand(editor, close, () =>
                        editor.commands.restartDocumentNumbering(),
                      )
                    }
                  >
                    <RotateCcw size={13} aria-hidden="true" />
                    重新从 1 开始
                  </button>
                  <button
                    type="button"
                    disabled={!editor.can().continueDocumentNumbering()}
                    onClick={() =>
                      runDocumentListMenuCommand(editor, close, () =>
                        editor.commands.continueDocumentNumbering(),
                      )
                    }
                  >
                    继续前一列表
                  </button>
                </div>
                <form
                  className="work-document-numbering-start"
                  aria-label="起始编号设置"
                  onSubmit={(event) => {
                    event.preventDefault();
                    applyStart(close);
                  }}
                >
                  <span>起始编号</span>
                  <OfficeNumberField
                    ariaLabel="起始编号"
                    value={startValue}
                    min={1}
                    max={MAX_DOCUMENT_NUMBERING_START}
                    step={1}
                    onValueChange={setStartValue}
                  />
                  <button
                    type="submit"
                    disabled={
                      validStart === null || validStart === activeState.start
                    }
                  >
                    应用起始值
                  </button>
                </form>
              </div>
            )}
            <button
              type="button"
              className="work-document-list-clear"
              disabled={!activeState}
              onClick={() =>
                runDocumentListMenuCommand(editor, close, () =>
                  editor.commands.clearDocumentList(),
                )
              }
            >
              <Unlink size={13} aria-hidden="true" />
              清除编号
            </button>
          </>
        )}
      </Popover>
    </div>
  );
}

function runDocumentListMenuCommand(
  editor: Editor,
  close: () => void,
  command: () => boolean,
): boolean {
  const handled = command();
  if (!handled) return false;
  close();
  editor.commands.focus();
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
