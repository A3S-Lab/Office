import {
  Check,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  ChevronUp,
  Ellipsis,
  GalleryVerticalEnd,
  Highlighter,
  Loader2,
  Minus,
  MousePointer2,
  MoveHorizontal,
  Pencil,
  Plus,
  Redo2,
  Save,
  Scan,
  Search,
  SlidersHorizontal,
  Strikethrough,
  Trash2,
  Type,
  Underline,
  Undo2,
  X,
} from 'lucide-react';
import {
  type FormEvent,
  type KeyboardEvent,
  type ReactNode,
  type RefObject,
  useEffect,
  useId,
  useRef,
  useState,
} from 'react';
import {
  Button,
  IconButton,
  Popover,
  StatusBadge,
} from '../../../design-system/primitives';
import { OfficeColorPicker } from './office-color-picker';
import { OfficeTextField } from './office-controls';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import type { PdfAnnotationControllerState } from './pdf-annotation-controller';
import type {
  PdfEditorCanCommands,
  PdfEditorCommands,
} from './pdf-editor-extensions';
import type { PdfViewerControllerState } from './pdf-viewer-controller';

export type PdfSaveState = 'idle' | 'saving' | 'saved' | 'error';

interface PdfPageNavigationControl {
  controlsId: string;
  expanded: boolean;
  onOpen: () => void;
  toggleRef: RefObject<HTMLButtonElement | null>;
}

const pdfKeyboardShortcuts = {
  deleteAnnotation: 'Delete Backspace',
  fitPage: 'Control+0 Meta+0',
  redo: 'Control+Shift+Z Meta+Shift+Z Control+Y Meta+Y',
  save: 'Control+S Meta+S',
  search: 'Control+F Meta+F',
  undo: 'Control+Z Meta+Z',
  zoomIn: 'Control+= Meta+= Control+Shift++ Meta+Shift++',
  zoomOut: 'Control+- Meta+-',
} as const;

export function PdfToolbar({
  annotationState,
  can,
  commands,
  editable,
  pageNavigation,
  saveAvailable,
  saveLabel,
  saveState,
  searchInputRef,
  state,
}: {
  annotationState: PdfAnnotationControllerState;
  can: PdfEditorCanCommands;
  commands: PdfEditorCommands;
  editable: boolean;
  pageNavigation?: PdfPageNavigationControl;
  saveAvailable?: boolean;
  saveLabel: string;
  saveState: PdfSaveState;
  searchInputRef: RefObject<HTMLInputElement | null>;
  state: PdfViewerControllerState;
}) {
  const showSave = saveAvailable ?? editable;
  const [pageValue, setPageValue] = useState('');
  const [searchValue, setSearchValue] = useState('');
  const cancelPageBlurCommitRef = useRef(false);

  useEffect(() => {
    setPageValue(state.currentPage > 0 ? String(state.currentPage) : '');
  }, [state.currentPage]);

  useEffect(() => {
    setSearchValue(state.search.query);
  }, [state.search.query]);

  const commitPage = () => {
    const page = Number(pageValue);
    if (can.goToPage(page)) {
      commands.goToPage(page);
      return;
    }
    setPageValue(state.currentPage > 0 ? String(state.currentPage) : '');
  };

  const submitSearch = (event: FormEvent) => {
    event.preventDefault();
    const query = searchValue.trim();
    if (query && query === state.search.query) {
      if (state.search.loading) return;
      if (state.search.error) {
        commands.search(query);
        return;
      }
      if (state.search.total > 0 && can.nextSearchResult()) {
        commands.nextSearchResult();
      }
      return;
    }
    commands.search(query);
  };

  return (
    <header className="work-pdf-toolbar" role="toolbar" aria-label="PDF 工具栏">
      {showSave && (
        <div className="work-pdf-toolbar-group work-pdf-save">
          <output aria-label="PDF 保存状态" aria-live="polite">
            {saveState === 'saving' && (
              <StatusBadge tone="info">
                <Loader2 className="spin" size={12} /> 保存中
              </StatusBadge>
            )}
            {saveState === 'saved' && (
              <StatusBadge tone="success">
                <Check size={12} /> 已保存
              </StatusBadge>
            )}
            {saveState === 'error' && (
              <StatusBadge tone="danger">保存失败</StatusBadge>
            )}
          </output>
          <Button
            tone="secondary"
            title={`${saveLabel}（Cmd/Ctrl+S）`}
            aria-keyshortcuts={pdfKeyboardShortcuts.save}
            disabled={!can.save()}
            onClick={() => void commands.save()}
          >
            <Save size={14} />
            {saveLabel}
          </Button>
        </div>
      )}

      {editable && (
        <div className="work-pdf-toolbar-group work-pdf-history">
          <IconButton
            label="撤销"
            title="撤销（Cmd/Ctrl+Z）"
            aria-keyshortcuts={pdfKeyboardShortcuts.undo}
            disabled={!can.undo()}
            onClick={commands.undo}
          >
            <Undo2 size={15} />
          </IconButton>
          <IconButton
            label="重做"
            title="重做（Cmd/Ctrl+Shift+Z 或 Cmd/Ctrl+Y）"
            aria-keyshortcuts={pdfKeyboardShortcuts.redo}
            disabled={!can.redo()}
            onClick={commands.redo}
          >
            <Redo2 size={15} />
          </IconButton>
        </div>
      )}

      {editable && (
        <fieldset className="work-pdf-toolbar-group work-pdf-annotation">
          <legend className="sr-only">PDF 批注工具</legend>
          <IconButton
            className="work-pdf-annotation-selection"
            label="选择"
            selected={annotationState.activeToolId === null}
            disabled={!can.selectAnnotationTool(null)}
            onClick={() => commands.selectAnnotationTool(null)}
          >
            <MousePointer2 size={14} />
          </IconButton>
          <IconButton
            label="高亮"
            selected={annotationState.activeToolId === 'highlight'}
            disabled={!can.selectAnnotationTool('highlight')}
            onClick={() => commands.selectAnnotationTool('highlight')}
          >
            <Highlighter size={14} />
          </IconButton>
          <IconButton
            className="work-pdf-annotation-optional"
            label="下划线批注"
            selected={annotationState.activeToolId === 'underline'}
            disabled={!can.selectAnnotationTool('underline')}
            onClick={() => commands.selectAnnotationTool('underline')}
          >
            <Underline size={14} />
          </IconButton>
          <IconButton
            className="work-pdf-annotation-optional"
            label="删除线批注"
            selected={annotationState.activeToolId === 'strikeout'}
            disabled={!can.selectAnnotationTool('strikeout')}
            onClick={() => commands.selectAnnotationTool('strikeout')}
          >
            <Strikethrough size={14} />
          </IconButton>
          <IconButton
            className="work-pdf-annotation-ink"
            label="画笔"
            selected={annotationState.activeToolId === 'ink'}
            disabled={!can.selectAnnotationTool('ink')}
            onClick={() => commands.selectAnnotationTool('ink')}
          >
            <Pencil size={14} />
          </IconButton>
          <IconButton
            className="work-pdf-annotation-optional"
            label="文字批注"
            selected={annotationState.activeToolId === 'freeText'}
            disabled={!can.selectAnnotationTool('freeText')}
            onClick={() => commands.selectAnnotationTool('freeText')}
          >
            <Type size={14} />
          </IconButton>
          <OfficeColorPicker
            ariaLabel="批注颜色"
            className="work-pdf-annotation-color"
            compact
            value={annotationState.annotationColor}
            disabled={!can.setAnnotationColor(annotationState.annotationColor)}
            onValueChange={commands.setAnnotationColor}
          />
          <PdfAnnotationStyleControl
            annotationState={annotationState}
            can={can}
            commands={commands}
          />
          <IconButton
            className="work-pdf-annotation-delete"
            label="删除所选批注"
            title="删除所选批注（Delete / Backspace）"
            aria-keyshortcuts={pdfKeyboardShortcuts.deleteAnnotation}
            disabled={!can.deleteAnnotationSelection()}
            onClick={commands.deleteAnnotationSelection}
          >
            <Trash2 size={14} />
          </IconButton>
        </fieldset>
      )}

      <PdfToolbarOverflow
        annotationState={annotationState}
        can={can}
        commands={commands}
        editable={editable}
        state={state}
      />

      <search className="work-pdf-search">
        <form onSubmit={submitSearch}>
          <Search size={14} aria-hidden="true" />
          <OfficeTextField
            ref={searchInputRef}
            type="search"
            aria-label="在 PDF 中搜索"
            aria-keyshortcuts={pdfKeyboardShortcuts.search}
            placeholder="搜索"
            value={searchValue}
            disabled={!can.search(searchValue)}
            onChange={(event) => setSearchValue(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Escape') {
                event.preventDefault();
                event.stopPropagation();
                setSearchValue('');
                commands.clearSearch();
              } else if (
                event.key === 'Enter' &&
                event.shiftKey &&
                searchValue.trim() === state.search.query
              ) {
                event.preventDefault();
                event.stopPropagation();
                if (
                  !state.search.loading &&
                  !state.search.error &&
                  state.search.total > 0 &&
                  can.previousSearchResult()
                ) {
                  commands.previousSearchResult();
                }
              }
            }}
          />
          {(searchValue || state.search.active) && (
            <IconButton
              className="work-pdf-search-clear"
              label="清除搜索"
              onClick={() => {
                setSearchValue('');
                commands.clearSearch();
                searchInputRef.current?.focus();
              }}
            >
              <X size={13} />
            </IconButton>
          )}
          <output className="work-pdf-search-state" aria-live="polite">
            {searchStatus(state)}
          </output>
          <IconButton
            label="上一个搜索结果"
            disabled={!can.previousSearchResult()}
            onClick={commands.previousSearchResult}
          >
            <ChevronUp size={14} />
          </IconButton>
          <IconButton
            label="下一个搜索结果"
            disabled={!can.nextSearchResult()}
            onClick={commands.nextSearchResult}
          >
            <ChevronDown size={14} />
          </IconButton>
        </form>
      </search>

      <div className="work-pdf-toolbar-group work-pdf-page-controls">
        {pageNavigation && (
          <IconButton
            ref={pageNavigation.toggleRef}
            className="work-pdf-page-navigation-toggle"
            label="打开 PDF 页面导航"
            tooltip="页面缩略图"
            aria-controls={pageNavigation.controlsId}
            aria-expanded={pageNavigation.expanded}
            onClick={pageNavigation.onOpen}
          >
            <GalleryVerticalEnd size={15} />
            <span className="sr-only">
              第 {Math.max(1, state.currentPage)} 页
            </span>
          </IconButton>
        )}
        <IconButton
          className="work-pdf-page-step"
          label="上一页"
          disabled={!can.previousPage()}
          onClick={commands.previousPage}
        >
          <ChevronLeft size={15} />
        </IconButton>
        <OfficeTextField
          className="work-pdf-page-field"
          aria-label="页码"
          inputMode="numeric"
          value={pageValue}
          disabled={!can.goToPage(state.currentPage || 1)}
          onBlur={() => {
            if (cancelPageBlurCommitRef.current) {
              cancelPageBlurCommitRef.current = false;
              return;
            }
            commitPage();
          }}
          onFocus={() => {
            cancelPageBlurCommitRef.current = false;
          }}
          onChange={(event) =>
            setPageValue(event.target.value.replace(/\D/g, ''))
          }
          onKeyDown={(event: KeyboardEvent<HTMLInputElement>) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              commitPage();
              event.currentTarget.select();
            } else if (event.key === 'Escape') {
              event.preventDefault();
              event.stopPropagation();
              cancelPageBlurCommitRef.current = true;
              setPageValue(String(state.currentPage));
              event.currentTarget.blur();
            }
          }}
        />
        <span className="work-pdf-page-total">/ {state.totalPages || '—'}</span>
        <IconButton
          className="work-pdf-page-step"
          label="下一页"
          disabled={!can.nextPage()}
          onClick={commands.nextPage}
        >
          <ChevronRight size={15} />
        </IconButton>
      </div>

      <div className="work-pdf-toolbar-group work-pdf-zoom-controls">
        <IconButton
          label="缩小"
          title="缩小（Cmd/Ctrl+-）"
          aria-keyshortcuts={pdfKeyboardShortcuts.zoomOut}
          disabled={!can.zoomOut()}
          onClick={commands.zoomOut}
        >
          <Minus size={14} />
        </IconButton>
        <output aria-label="PDF 缩放比例">{state.zoomPercent}%</output>
        <IconButton
          label="放大"
          title="放大（Cmd/Ctrl++）"
          aria-keyshortcuts={pdfKeyboardShortcuts.zoomIn}
          disabled={!can.zoomIn()}
          onClick={commands.zoomIn}
        >
          <Plus size={14} />
        </IconButton>
        <button
          type="button"
          className="work-pdf-fit-button"
          title="整页（Cmd/Ctrl+0）"
          aria-keyshortcuts={pdfKeyboardShortcuts.fitPage}
          aria-pressed={state.zoomMode === 'fit-page'}
          disabled={!can.fitPage()}
          onClick={commands.fitPage}
        >
          整页
        </button>
        <button
          type="button"
          className="work-pdf-fit-button"
          aria-pressed={state.zoomMode === 'fit-width'}
          disabled={!can.fitWidth()}
          onClick={commands.fitWidth}
        >
          页宽
        </button>
      </div>
    </header>
  );
}

const PDF_ANNOTATION_OPACITY_OPTIONS = [0.25, 0.5, 0.75, 1] as const;
const PDF_ANNOTATION_STROKE_WIDTH_OPTIONS = [1, 2, 4, 6, 10] as const;

function PdfAnnotationStyleControl({
  annotationState,
  can,
  commands,
}: {
  annotationState: PdfAnnotationControllerState;
  can: PdfEditorCanCommands;
  commands: PdfEditorCommands;
}) {
  const [open, setOpen] = useState(false);
  const panelRef = useRef<HTMLElement>(null);
  const opacityGroupName = useId();
  const strokeWidthGroupName = useId();
  const opacityPercent = Math.round(annotationState.annotationOpacity * 100);
  const disabled =
    !can.setAnnotationOpacity(annotationState.annotationOpacity) &&
    !can.setAnnotationStrokeWidth(annotationState.annotationStrokeWidth);
  const title = annotationState.supportsStrokeWidth
    ? `批注样式（${opacityPercent}%，线宽 ${annotationState.annotationStrokeWidth}）`
    : `批注样式（${opacityPercent}%）`;

  return (
    <Popover
      label="批注样式"
      panelLabel="批注样式"
      panelRole="dialog"
      placement="bottom-end"
      portal
      open={open}
      panelRef={panelRef}
      onOpenChange={(nextOpen) => {
        setOpen(nextOpen);
        if (!nextOpen) return;
        requestAnimationFrame(() => {
          const panel = panelRef.current;
          const target =
            panel?.querySelector<HTMLInputElement>(
              'input[type="radio"]:checked:not(:disabled)',
            ) ??
            panel?.querySelector<HTMLInputElement>(
              'input[type="radio"]:not(:disabled)',
            );
          target?.focus({ preventScroll: true });
        });
      }}
      disabled={disabled}
      className="work-pdf-annotation-style"
      panelClassName="work-pdf-annotation-style-panel"
      trigger={(triggerProps) => (
        <button {...triggerProps} className="ds-icon-button" title={title}>
          <SlidersHorizontal size={14} />
        </button>
      )}
    >
      <div className="work-pdf-annotation-style-content">
        {annotationState.supportsOpacity && (
          <fieldset className="work-pdf-annotation-style-row">
            <legend>透明度</legend>
            <div className="work-pdf-annotation-style-options">
              {PDF_ANNOTATION_OPACITY_OPTIONS.map((opacity) => {
                const label = `${Math.round(opacity * 100)}%`;
                return (
                  <label key={opacity}>
                    <input
                      type="radio"
                      name={opacityGroupName}
                      aria-label={`透明度 ${label}`}
                      checked={annotationState.annotationOpacity === opacity}
                      disabled={!can.setAnnotationOpacity(opacity)}
                      onChange={() => commands.setAnnotationOpacity(opacity)}
                      onKeyDown={movePdfAnnotationStyleOption}
                    />
                    <span>{label}</span>
                  </label>
                );
              })}
            </div>
          </fieldset>
        )}
        {annotationState.supportsStrokeWidth && (
          <fieldset className="work-pdf-annotation-style-row">
            <legend>线宽</legend>
            <div className="work-pdf-annotation-style-options">
              {PDF_ANNOTATION_STROKE_WIDTH_OPTIONS.map((strokeWidth) => (
                <label key={strokeWidth}>
                  <input
                    type="radio"
                    name={strokeWidthGroupName}
                    aria-label={`线宽 ${strokeWidth}`}
                    checked={
                      annotationState.annotationStrokeWidth === strokeWidth
                    }
                    disabled={!can.setAnnotationStrokeWidth(strokeWidth)}
                    onChange={() =>
                      commands.setAnnotationStrokeWidth(strokeWidth)
                    }
                    onKeyDown={movePdfAnnotationStyleOption}
                  />
                  <span>{strokeWidth}</span>
                </label>
              ))}
            </div>
          </fieldset>
        )}
      </div>
    </Popover>
  );
}

function movePdfAnnotationStyleOption(
  event: KeyboardEvent<HTMLInputElement>,
): void {
  if (
    ![
      'ArrowDown',
      'ArrowLeft',
      'ArrowRight',
      'ArrowUp',
      'Home',
      'End',
    ].includes(event.key)
  ) {
    return;
  }
  const options = [
    ...(event.currentTarget
      .closest('fieldset')
      ?.querySelectorAll<HTMLInputElement>(
        'input[type="radio"]:not(:disabled)',
      ) ?? []),
  ];
  if (!options.length) return;
  event.preventDefault();
  const current = options.indexOf(event.currentTarget);
  const nextIndex =
    event.key === 'Home'
      ? 0
      : event.key === 'End'
        ? options.length - 1
        : event.key === 'ArrowDown' || event.key === 'ArrowRight'
          ? (current + 1 + options.length) % options.length
          : (current - 1 + options.length) % options.length;
  const next = options[nextIndex];
  next?.focus({ preventScroll: true });
  next?.click();
}

function PdfToolbarOverflow({
  annotationState,
  can,
  commands,
  editable,
  state,
}: {
  annotationState: PdfAnnotationControllerState;
  can: PdfEditorCanCommands;
  commands: PdfEditorCommands;
  editable: boolean;
  state: PdfViewerControllerState;
}) {
  const hasOverflowTools =
    (editable && annotationState.available) ||
    (editable && state.features.history) ||
    state.features.navigation ||
    state.features.search ||
    state.features.zoom;
  return (
    <Popover
      label="更多 PDF 工具"
      panelLabel="更多 PDF 工具"
      panelRole="menu"
      placement="bottom-end"
      portal
      focusFirstOnOpen
      onPanelKeyDown={moveOfficeMenuFocus}
      disabled={!hasOverflowTools}
      className="work-pdf-overflow"
      panelClassName="work-pdf-overflow-panel"
      trigger={(triggerProps) => (
        <button
          {...triggerProps}
          className="ds-icon-button work-pdf-overflow-trigger"
          title="更多 PDF 工具"
        >
          <Ellipsis size={16} />
        </button>
      )}
    >
      {(close) => {
        const select = (command: () => void) => {
          close();
          command();
        };
        return (
          <>
            {editable && annotationState.available && (
              <fieldset
                className="work-pdf-overflow-group"
                aria-label="批注工具"
              >
                <PdfOverflowAction
                  label="选择"
                  active={annotationState.activeToolId === null}
                  disabled={!can.selectAnnotationTool(null)}
                  onSelect={() =>
                    select(() => commands.selectAnnotationTool(null))
                  }
                >
                  <MousePointer2 size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label="高亮"
                  active={annotationState.activeToolId === 'highlight'}
                  disabled={!can.selectAnnotationTool('highlight')}
                  onSelect={() =>
                    select(() => commands.selectAnnotationTool('highlight'))
                  }
                >
                  <Highlighter size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label="画笔"
                  active={annotationState.activeToolId === 'ink'}
                  disabled={!can.selectAnnotationTool('ink')}
                  onSelect={() =>
                    select(() => commands.selectAnnotationTool('ink'))
                  }
                >
                  <Pencil size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label="下划线批注"
                  active={annotationState.activeToolId === 'underline'}
                  disabled={!can.selectAnnotationTool('underline')}
                  onSelect={() =>
                    select(() => commands.selectAnnotationTool('underline'))
                  }
                >
                  <Underline size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label="删除线批注"
                  active={annotationState.activeToolId === 'strikeout'}
                  disabled={!can.selectAnnotationTool('strikeout')}
                  onSelect={() =>
                    select(() => commands.selectAnnotationTool('strikeout'))
                  }
                >
                  <Strikethrough size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label="文字批注"
                  active={annotationState.activeToolId === 'freeText'}
                  disabled={!can.selectAnnotationTool('freeText')}
                  onSelect={() =>
                    select(() => commands.selectAnnotationTool('freeText'))
                  }
                >
                  <Type size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label="删除所选批注"
                  ariaKeyShortcuts={pdfKeyboardShortcuts.deleteAnnotation}
                  disabled={!can.deleteAnnotationSelection()}
                  onSelect={() => select(commands.deleteAnnotationSelection)}
                >
                  <Trash2 size={15} />
                </PdfOverflowAction>
              </fieldset>
            )}
            {editable &&
              annotationState.available &&
              annotationState.supportsOpacity && (
                <fieldset
                  className="work-pdf-overflow-group"
                  aria-label="透明度"
                >
                  {PDF_ANNOTATION_OPACITY_OPTIONS.map((opacity) => {
                    const label = `透明度 ${Math.round(opacity * 100)}%`;
                    return (
                      <PdfOverflowAction
                        key={opacity}
                        label={label}
                        active={annotationState.annotationOpacity === opacity}
                        disabled={!can.setAnnotationOpacity(opacity)}
                        onSelect={() =>
                          select(() => commands.setAnnotationOpacity(opacity))
                        }
                      >
                        <SlidersHorizontal size={15} />
                      </PdfOverflowAction>
                    );
                  })}
                </fieldset>
              )}
            {editable &&
              annotationState.available &&
              annotationState.supportsStrokeWidth && (
                <fieldset className="work-pdf-overflow-group" aria-label="线宽">
                  {PDF_ANNOTATION_STROKE_WIDTH_OPTIONS.map((strokeWidth) => (
                    <PdfOverflowAction
                      key={strokeWidth}
                      label={`线宽 ${strokeWidth}`}
                      active={
                        annotationState.annotationStrokeWidth === strokeWidth
                      }
                      disabled={!can.setAnnotationStrokeWidth(strokeWidth)}
                      onSelect={() =>
                        select(() =>
                          commands.setAnnotationStrokeWidth(strokeWidth),
                        )
                      }
                    >
                      <Pencil size={15} />
                    </PdfOverflowAction>
                  ))}
                </fieldset>
              )}
            {editable && state.features.history && (
              <fieldset
                className="work-pdf-overflow-group"
                aria-label="历史记录"
              >
                <PdfOverflowAction
                  label="撤销"
                  ariaKeyShortcuts={pdfKeyboardShortcuts.undo}
                  disabled={!can.undo()}
                  onSelect={() => select(commands.undo)}
                >
                  <Undo2 size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label="重做"
                  ariaKeyShortcuts={pdfKeyboardShortcuts.redo}
                  disabled={!can.redo()}
                  onSelect={() => select(commands.redo)}
                >
                  <Redo2 size={15} />
                </PdfOverflowAction>
              </fieldset>
            )}
            {state.features.navigation && (
              <fieldset className="work-pdf-overflow-group" aria-label="翻页">
                <PdfOverflowAction
                  label="上一页"
                  disabled={!can.previousPage()}
                  onSelect={() => select(commands.previousPage)}
                >
                  <ChevronLeft size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label="下一页"
                  disabled={!can.nextPage()}
                  onSelect={() => select(commands.nextPage)}
                >
                  <ChevronRight size={15} />
                </PdfOverflowAction>
              </fieldset>
            )}
            {state.features.zoom && (
              <fieldset className="work-pdf-overflow-group" aria-label="缩放">
                <PdfOverflowAction
                  label="缩小"
                  ariaKeyShortcuts={pdfKeyboardShortcuts.zoomOut}
                  disabled={!can.zoomOut()}
                  onSelect={() => select(commands.zoomOut)}
                >
                  <Minus size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label="放大"
                  ariaKeyShortcuts={pdfKeyboardShortcuts.zoomIn}
                  disabled={!can.zoomIn()}
                  onSelect={() => select(commands.zoomIn)}
                >
                  <Plus size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label="整页"
                  active={state.zoomMode === 'fit-page'}
                  ariaKeyShortcuts={pdfKeyboardShortcuts.fitPage}
                  disabled={!can.fitPage()}
                  onSelect={() => select(commands.fitPage)}
                >
                  <Scan size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label="页宽"
                  active={state.zoomMode === 'fit-width'}
                  disabled={!can.fitWidth()}
                  onSelect={() => select(commands.fitWidth)}
                >
                  <MoveHorizontal size={15} />
                </PdfOverflowAction>
              </fieldset>
            )}
            {state.features.search && (
              <fieldset
                className="work-pdf-overflow-group work-pdf-overflow-narrow"
                aria-label="搜索结果"
              >
                <PdfOverflowAction
                  label="上一个搜索结果"
                  disabled={!can.previousSearchResult()}
                  onSelect={() => select(commands.previousSearchResult)}
                >
                  <ChevronUp size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label="下一个搜索结果"
                  disabled={!can.nextSearchResult()}
                  onSelect={() => select(commands.nextSearchResult)}
                >
                  <ChevronDown size={15} />
                </PdfOverflowAction>
              </fieldset>
            )}
          </>
        );
      }}
    </Popover>
  );
}

function PdfOverflowAction({
  active,
  ariaKeyShortcuts,
  children,
  disabled,
  label,
  onSelect,
}: {
  active?: boolean;
  ariaKeyShortcuts?: string;
  children: ReactNode;
  disabled: boolean;
  label: string;
  onSelect: () => void;
}) {
  const content = (
    <>
      <span aria-hidden="true">{children}</span>
      <span>{label}</span>
      {active && <Check size={14} aria-hidden="true" />}
    </>
  );
  if (active !== undefined) {
    return (
      <button
        type="button"
        role="menuitemradio"
        aria-checked={active}
        aria-keyshortcuts={ariaKeyShortcuts}
        tabIndex={-1}
        data-active={active ? 'true' : undefined}
        disabled={disabled}
        onClick={onSelect}
      >
        {content}
      </button>
    );
  }
  return (
    <button
      type="button"
      role="menuitem"
      aria-keyshortcuts={ariaKeyShortcuts}
      tabIndex={-1}
      disabled={disabled}
      onClick={onSelect}
    >
      {content}
    </button>
  );
}

function searchStatus(state: PdfViewerControllerState): string {
  const { search } = state;
  if (search.loading) return '搜索中';
  if (search.error) return '失败';
  if (!search.query && !search.active) return '';
  if (search.total === 0) return '0 / 0';
  return `${search.activeResultIndex + 1} / ${search.total}`;
}
