import {
  Check,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  ChevronsLeft,
  ChevronsRight,
  ChevronUp,
  Ellipsis,
  GalleryVerticalEnd,
  Highlighter,
  LayoutGrid,
  Loader2,
  Minus,
  MousePointer2,
  MoveHorizontal,
  Pencil,
  Plus,
  Ratio,
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
import { officeMessage } from '../../../i18n/office-locale';
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
import { WorkOfficeCollaborationParticipants } from './office-collaboration-participants';
import { useOfficeMessages } from './office-messages-context';
import { OfficeColorPicker } from './office-color-picker';
import { OfficeTextField } from './office-controls';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import { moveOfficeToolbarFocus } from './office-toolbar-keyboard';
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
  actualSize: 'Control+1 Meta+1',
  deleteAnnotation: 'Delete Backspace',
  fitPage: 'Control+0 Meta+0',
  fitWidth: 'Control+2 Meta+2',
  firstPage: 'Control+Home Meta+Home',
  lastPage: 'Control+End Meta+End',
  nextPage: 'PageDown Space',
  previousPage: 'PageUp Shift+Space',
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
  pageOrganizationAvailable,
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
  pageOrganizationAvailable?: boolean;
  pageNavigation?: PdfPageNavigationControl;
  saveAvailable?: boolean;
  saveLabel: string;
  saveState: PdfSaveState;
  searchInputRef: RefObject<HTMLInputElement | null>;
  state: PdfViewerControllerState;
}) {
  const messages = useOfficeMessages();
  const showSave = saveAvailable ?? editable;
  const showPageOrganization = pageOrganizationAvailable ?? editable;
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
    <header
      className="work-pdf-toolbar"
      role="toolbar"
      aria-label={officeMessage(messages, 'pdf.toolbar.aria')}
      onKeyDown={moveOfficeToolbarFocus}
    >
      {showSave && (
        <div className="work-pdf-toolbar-group work-pdf-save">
          <output aria-label={officeMessage(messages, 'pdf.toolbar.saveStatusAria')} aria-live="polite">
            {saveState === 'saving' && (
              <StatusBadge tone="info">
                <Loader2 className="spin" size={12} /> {officeMessage(messages, 'pdf.toolbar.saving')}
              </StatusBadge>
            )}
            {saveState === 'saved' && (
              <StatusBadge tone="success">
                <Check size={12} /> {officeMessage(messages, 'pdf.toolbar.saved')}
              </StatusBadge>
            )}
            {saveState === 'error' && (
              <StatusBadge tone="danger">{officeMessage(messages, 'pdf.toolbar.saveFailed')}</StatusBadge>
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
            label={officeMessage(messages, 'pdf.toolbar.undo')}
            title={officeMessage(messages, 'pdf.toolbar.undoTitle')}
            aria-keyshortcuts={pdfKeyboardShortcuts.undo}
            disabled={!can.undo()}
            onClick={commands.undo}
          >
            <Undo2 size={15} />
          </IconButton>
          <IconButton
            label={officeMessage(messages, 'pdf.toolbar.redo')}
            title={officeMessage(messages, 'pdf.toolbar.redoTitle')}
            aria-keyshortcuts={pdfKeyboardShortcuts.redo}
            disabled={!can.redo()}
            onClick={commands.redo}
          >
            <Redo2 size={15} />
          </IconButton>
        </div>
      )}

      {showPageOrganization && (
        <div className="work-pdf-toolbar-group work-pdf-page-organization">
          <Button
            tone="secondary"
            className="work-pdf-page-organizer-trigger"
            data-pdf-page-organizer-trigger
            aria-label={officeMessage(messages, 'pdf.toolbar.organizeAria')}
            title={officeMessage(messages, 'pdf.toolbar.organizeTitle')}
            disabled={!can.openPageOrganizer()}
            onClick={commands.openPageOrganizer}
          >
            <LayoutGrid size={15} />
            <span className="work-pdf-page-organizer-label">{officeMessage(messages, 'pdf.toolbar.organizeLabel')}</span>
          </Button>
        </div>
      )}

      {editable && (
        <fieldset className="work-pdf-toolbar-group work-pdf-annotation">
          <legend className="sr-only">{officeMessage(messages, 'pdf.toolbar.annotationToolsLegend')}</legend>
          <IconButton
            className="work-pdf-annotation-selection"
            label={officeMessage(messages, 'pdf.toolbar.select')}
            selected={annotationState.activeToolId === null}
            disabled={!can.selectAnnotationTool(null)}
            onClick={() => commands.selectAnnotationTool(null)}
          >
            <MousePointer2 size={14} />
          </IconButton>
          <IconButton
            label={officeMessage(messages, 'pdf.toolbar.highlight')}
            selected={annotationState.activeToolId === 'highlight'}
            disabled={!can.selectAnnotationTool('highlight')}
            onClick={() => commands.selectAnnotationTool('highlight')}
          >
            <Highlighter size={14} />
          </IconButton>
          <IconButton
            className="work-pdf-annotation-optional"
            label={officeMessage(messages, 'pdf.toolbar.underline')}
            selected={annotationState.activeToolId === 'underline'}
            disabled={!can.selectAnnotationTool('underline')}
            onClick={() => commands.selectAnnotationTool('underline')}
          >
            <Underline size={14} />
          </IconButton>
          <IconButton
            className="work-pdf-annotation-optional"
            label={officeMessage(messages, 'pdf.toolbar.strikeout')}
            selected={annotationState.activeToolId === 'strikeout'}
            disabled={!can.selectAnnotationTool('strikeout')}
            onClick={() => commands.selectAnnotationTool('strikeout')}
          >
            <Strikethrough size={14} />
          </IconButton>
          <IconButton
            className="work-pdf-annotation-ink"
            label={officeMessage(messages, 'pdf.toolbar.ink')}
            selected={annotationState.activeToolId === 'ink'}
            disabled={!can.selectAnnotationTool('ink')}
            onClick={() => commands.selectAnnotationTool('ink')}
          >
            <Pencil size={14} />
          </IconButton>
          <IconButton
            className="work-pdf-annotation-optional"
            label={officeMessage(messages, 'pdf.toolbar.freeText')}
            selected={annotationState.activeToolId === 'freeText'}
            disabled={!can.selectAnnotationTool('freeText')}
            onClick={() => commands.selectAnnotationTool('freeText')}
          >
            <Type size={14} />
          </IconButton>
          <OfficeColorPicker
            ariaLabel={officeMessage(messages, 'pdf.toolbar.colorAria')}
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
            label={officeMessage(messages, 'pdf.toolbar.deleteAnnotation')}
            title={officeMessage(messages, 'pdf.toolbar.deleteAnnotationTitle')}
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
        pageOrganizationAvailable={showPageOrganization}
        state={state}
      />

      <search className="work-pdf-search">
        <form onSubmit={submitSearch}>
          <Search size={14} aria-hidden="true" />
          <OfficeTextField
            ref={searchInputRef}
            type="search"
            aria-label={officeMessage(messages, 'pdf.toolbar.searchAria')}
            aria-keyshortcuts={pdfKeyboardShortcuts.search}
            placeholder={officeMessage(messages, 'pdf.toolbar.searchPlaceholder')}
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
              label={officeMessage(messages, 'pdf.toolbar.clearSearch')}
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
            {searchStatus(state, messages)}
          </output>
          <IconButton
            label={officeMessage(messages, 'pdf.toolbar.prevMatch')}
            disabled={!can.previousSearchResult()}
            onClick={commands.previousSearchResult}
          >
            <ChevronUp size={14} />
          </IconButton>
          <IconButton
            label={officeMessage(messages, 'pdf.toolbar.nextMatch')}
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
            label={officeMessage(messages, 'pdf.toolbar.openNav')}
            tooltip={officeMessage(messages, 'pdf.toolbar.navTooltip')}
            aria-controls={pageNavigation.controlsId}
            aria-expanded={pageNavigation.expanded}
            onClick={pageNavigation.onOpen}
          >
            <GalleryVerticalEnd size={15} />
            <span className="sr-only">
              {officeMessage(messages, 'pdf.toolbar.pageBadge', { page: String(Math.max(1, state.currentPage)) })}
            </span>
          </IconButton>
        )}
        <IconButton
          className="work-pdf-page-step"
          label={officeMessage(messages, 'pdf.toolbar.prevPage')}
          title={officeMessage(messages, 'pdf.toolbar.prevPageTitle')}
          aria-keyshortcuts={pdfKeyboardShortcuts.previousPage}
          disabled={!can.previousPage()}
          onClick={commands.previousPage}
        >
          <ChevronLeft size={15} />
        </IconButton>
        <OfficeTextField
          className="work-pdf-page-field"
          aria-label={officeMessage(messages, 'pdf.toolbar.pageNumberAria')}
          inputMode="numeric"
          value={pageValue}
          disabled={!can.goToPage(state.currentPage || 1)}
          data-office-escape-consumer={
            pageValue !==
              (state.currentPage > 0 ? String(state.currentPage) : '') ||
            undefined
          }
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
            } else if (
              event.key === 'Escape' &&
              pageValue !==
                (state.currentPage > 0 ? String(state.currentPage) : '')
            ) {
              event.preventDefault();
              event.stopPropagation();
              cancelPageBlurCommitRef.current = true;
              setPageValue(
                state.currentPage > 0 ? String(state.currentPage) : '',
              );
              event.currentTarget.blur();
            }
          }}
        />
        <span className="work-pdf-page-total">/ {state.totalPages || '—'}</span>
        <IconButton
          className="work-pdf-page-step"
          label={officeMessage(messages, 'pdf.toolbar.nextPage')}
          title={officeMessage(messages, 'pdf.toolbar.nextPageTitle')}
          aria-keyshortcuts={pdfKeyboardShortcuts.nextPage}
          disabled={!can.nextPage()}
          onClick={commands.nextPage}
        >
          <ChevronRight size={15} />
        </IconButton>
      </div>

      <div className="work-pdf-toolbar-group work-pdf-zoom-controls">
        <IconButton
          label={officeMessage(messages, 'pdf.toolbar.zoomOut')}
          title={officeMessage(messages, 'pdf.toolbar.zoomOutTitle')}
          aria-keyshortcuts={pdfKeyboardShortcuts.zoomOut}
          disabled={!can.zoomOut()}
          onClick={commands.zoomOut}
        >
          <Minus size={14} />
        </IconButton>
        <output aria-label={officeMessage(messages, 'pdf.toolbar.zoomAria')}>{state.zoomPercent}%</output>
        <IconButton
          label={officeMessage(messages, 'pdf.toolbar.zoomIn')}
          title={officeMessage(messages, 'pdf.toolbar.zoomInTitle')}
          aria-keyshortcuts={pdfKeyboardShortcuts.zoomIn}
          disabled={!can.zoomIn()}
          onClick={commands.zoomIn}
        >
          <Plus size={14} />
        </IconButton>
        <button
          type="button"
          className="work-pdf-fit-button"
          aria-label={officeMessage(messages, 'pdf.toolbar.actualSizeAria')}
          title={officeMessage(messages, 'pdf.toolbar.actualSizeTitle')}
          aria-keyshortcuts={pdfKeyboardShortcuts.actualSize}
          aria-pressed={isPdfActualSize(state)}
          disabled={!can.actualSize()}
          onClick={commands.actualSize}
        >
          {officeMessage(messages, 'pdf.toolbar.actualSize')}
        </button>
        <button
          type="button"
          className="work-pdf-fit-button"
          aria-label={officeMessage(messages, 'pdf.toolbar.fitPageAria')}
          title={officeMessage(messages, 'pdf.toolbar.fitPageTitle')}
          aria-keyshortcuts={pdfKeyboardShortcuts.fitPage}
          aria-pressed={state.zoomMode === 'fit-page'}
          disabled={!can.fitPage()}
          onClick={commands.fitPage}
        >
          {officeMessage(messages, 'pdf.toolbar.fitPage')}
        </button>
        <button
          type="button"
          className="work-pdf-fit-button"
          aria-label={officeMessage(messages, 'pdf.toolbar.fitWidthAria')}
          title={officeMessage(messages, 'pdf.toolbar.fitWidthTitle')}
          aria-keyshortcuts={pdfKeyboardShortcuts.fitWidth}
          aria-pressed={state.zoomMode === 'fit-width'}
          disabled={!can.fitWidth()}
          onClick={commands.fitWidth}
        >
          {officeMessage(messages, 'pdf.toolbar.fitWidth')}
        </button>
      </div>
      <WorkOfficeCollaborationParticipants variant="toolbar" />
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
  const messages = useOfficeMessages();
  const [open, setOpen] = useState(false);
  const panelRef = useRef<HTMLElement>(null);
  const opacityGroupName = useId();
  const strokeWidthGroupName = useId();
  const opacityPercent = Math.round(annotationState.annotationOpacity * 100);
  const disabled =
    !can.setAnnotationOpacity(annotationState.annotationOpacity) &&
    !can.setAnnotationStrokeWidth(annotationState.annotationStrokeWidth);
  const title = annotationState.supportsStrokeWidth
    ? officeMessage(messages, 'pdf.toolbar.styleSummaryStroke', { opacity: String(opacityPercent), stroke: String(annotationState.annotationStrokeWidth) })
    : officeMessage(messages, 'pdf.toolbar.styleSummary', { opacity: String(opacityPercent) });

  return (
    <Popover
      label={officeMessage(messages, 'pdf.toolbar.style')}
      panelLabel={officeMessage(messages, 'pdf.toolbar.style')}
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
            <legend>{officeMessage(messages, 'pdf.toolbar.opacityLegend')}</legend>
            <div className="work-pdf-annotation-style-options">
              {PDF_ANNOTATION_OPACITY_OPTIONS.map((opacity) => {
                const label = `${Math.round(opacity * 100)}%`;
                return (
                  <label key={opacity}>
                    <input
                      type="radio"
                      name={opacityGroupName}
                      aria-label={officeMessage(messages, 'pdf.toolbar.opacityAria', { label })}
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
            <legend>{officeMessage(messages, 'pdf.toolbar.strokeLegend')}</legend>
            <div className="work-pdf-annotation-style-options">
              {PDF_ANNOTATION_STROKE_WIDTH_OPTIONS.map((strokeWidth) => (
                <label key={strokeWidth}>
                  <input
                    type="radio"
                    name={strokeWidthGroupName}
                    aria-label={officeMessage(messages, 'pdf.toolbar.strokeAria', { width: String(strokeWidth) })}
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
  pageOrganizationAvailable,
  state,
}: {
  annotationState: PdfAnnotationControllerState;
  can: PdfEditorCanCommands;
  commands: PdfEditorCommands;
  editable: boolean;
  pageOrganizationAvailable: boolean;
  state: PdfViewerControllerState;
}) {
  const messages = useOfficeMessages();
  const hasOverflowTools =
    (editable && annotationState.available) ||
    (editable && state.features.history) ||
    state.features.navigation ||
    state.features.search ||
    state.features.zoom;
  return (
    <Popover
      label={officeMessage(messages, 'pdf.toolbar.more')}
      panelLabel={officeMessage(messages, 'pdf.toolbar.more')}
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
          title={officeMessage(messages, 'pdf.toolbar.more')}
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
              <PdfOverflowGroup ariaLabel={officeMessage(messages, 'pdf.toolbar.overflowAnnotation')}>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.select')}
                  active={annotationState.activeToolId === null}
                  disabled={!can.selectAnnotationTool(null)}
                  onSelect={() =>
                    select(() => commands.selectAnnotationTool(null))
                  }
                >
                  <MousePointer2 size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.highlight')}
                  active={annotationState.activeToolId === 'highlight'}
                  disabled={!can.selectAnnotationTool('highlight')}
                  onSelect={() =>
                    select(() => commands.selectAnnotationTool('highlight'))
                  }
                >
                  <Highlighter size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.ink')}
                  active={annotationState.activeToolId === 'ink'}
                  disabled={!can.selectAnnotationTool('ink')}
                  onSelect={() =>
                    select(() => commands.selectAnnotationTool('ink'))
                  }
                >
                  <Pencil size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.underline')}
                  active={annotationState.activeToolId === 'underline'}
                  disabled={!can.selectAnnotationTool('underline')}
                  onSelect={() =>
                    select(() => commands.selectAnnotationTool('underline'))
                  }
                >
                  <Underline size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.strikeout')}
                  active={annotationState.activeToolId === 'strikeout'}
                  disabled={!can.selectAnnotationTool('strikeout')}
                  onSelect={() =>
                    select(() => commands.selectAnnotationTool('strikeout'))
                  }
                >
                  <Strikethrough size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.freeText')}
                  active={annotationState.activeToolId === 'freeText'}
                  disabled={!can.selectAnnotationTool('freeText')}
                  onSelect={() =>
                    select(() => commands.selectAnnotationTool('freeText'))
                  }
                >
                  <Type size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.deleteAnnotation')}
                  ariaKeyShortcuts={pdfKeyboardShortcuts.deleteAnnotation}
                  disabled={!can.deleteAnnotationSelection()}
                  onSelect={() => select(commands.deleteAnnotationSelection)}
                >
                  <Trash2 size={15} />
                </PdfOverflowAction>
              </PdfOverflowGroup>
            )}
            {editable &&
              annotationState.available &&
              annotationState.supportsOpacity && (
                <PdfOverflowGroup ariaLabel={officeMessage(messages, 'pdf.toolbar.overflowOpacity')}>
                  {PDF_ANNOTATION_OPACITY_OPTIONS.map((opacity) => {
                    const label = officeMessage(messages, 'pdf.toolbar.overflowOpacityItem', { percent: String(Math.round(opacity * 100)) });
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
                </PdfOverflowGroup>
              )}
            {editable &&
              annotationState.available &&
              annotationState.supportsStrokeWidth && (
                <PdfOverflowGroup ariaLabel={officeMessage(messages, 'pdf.toolbar.overflowStroke')}>
                  {PDF_ANNOTATION_STROKE_WIDTH_OPTIONS.map((strokeWidth) => (
                    <PdfOverflowAction
                      key={strokeWidth}
                      label={officeMessage(messages, 'pdf.toolbar.overflowStrokeItem', { width: String(strokeWidth) })}
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
                </PdfOverflowGroup>
              )}
            {editable && state.features.history && (
              <PdfOverflowGroup ariaLabel={officeMessage(messages, 'pdf.toolbar.overflowHistory')}>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.undo')}
                  ariaKeyShortcuts={pdfKeyboardShortcuts.undo}
                  disabled={!can.undo()}
                  onSelect={() => select(commands.undo)}
                >
                  <Undo2 size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.redo')}
                  ariaKeyShortcuts={pdfKeyboardShortcuts.redo}
                  disabled={!can.redo()}
                  onSelect={() => select(commands.redo)}
                >
                  <Redo2 size={15} />
                </PdfOverflowAction>
              </PdfOverflowGroup>
            )}
            {pageOrganizationAvailable && (
              <PdfOverflowGroup ariaLabel={officeMessage(messages, 'pdf.toolbar.overflowOrganize')}>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.organizeLabel')}
                  disabled={!can.openPageOrganizer()}
                  onSelect={() => select(commands.openPageOrganizer)}
                >
                  <LayoutGrid size={15} />
                </PdfOverflowAction>
              </PdfOverflowGroup>
            )}
            {state.features.navigation && (
              <PdfOverflowGroup ariaLabel={officeMessage(messages, 'pdf.toolbar.overflowNav')}>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.firstPage')}
                  ariaKeyShortcuts={pdfKeyboardShortcuts.firstPage}
                  disabled={!can.goToPage(1)}
                  onSelect={() => select(() => commands.goToPage(1))}
                >
                  <ChevronsLeft size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.prevPage')}
                  ariaKeyShortcuts={pdfKeyboardShortcuts.previousPage}
                  disabled={!can.previousPage()}
                  onSelect={() => select(commands.previousPage)}
                >
                  <ChevronLeft size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.nextPage')}
                  ariaKeyShortcuts={pdfKeyboardShortcuts.nextPage}
                  disabled={!can.nextPage()}
                  onSelect={() => select(commands.nextPage)}
                >
                  <ChevronRight size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.lastPage')}
                  ariaKeyShortcuts={pdfKeyboardShortcuts.lastPage}
                  disabled={!can.goToPage(state.totalPages || 1)}
                  onSelect={() =>
                    select(() => commands.goToPage(state.totalPages || 1))
                  }
                >
                  <ChevronsRight size={15} />
                </PdfOverflowAction>
              </PdfOverflowGroup>
            )}
            {state.features.zoom && (
              <PdfOverflowGroup ariaLabel={officeMessage(messages, 'pdf.toolbar.overflowZoom')}>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.zoomOut')}
                  ariaKeyShortcuts={pdfKeyboardShortcuts.zoomOut}
                  disabled={!can.zoomOut()}
                  onSelect={() => select(commands.zoomOut)}
                >
                  <Minus size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.zoomIn')}
                  ariaKeyShortcuts={pdfKeyboardShortcuts.zoomIn}
                  disabled={!can.zoomIn()}
                  onSelect={() => select(commands.zoomIn)}
                >
                  <Plus size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.actualSize')}
                  active={isPdfActualSize(state)}
                  ariaKeyShortcuts={pdfKeyboardShortcuts.actualSize}
                  disabled={!can.actualSize()}
                  onSelect={() => select(commands.actualSize)}
                >
                  <Ratio size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.fitPage')}
                  active={state.zoomMode === 'fit-page'}
                  ariaKeyShortcuts={pdfKeyboardShortcuts.fitPage}
                  disabled={!can.fitPage()}
                  onSelect={() => select(commands.fitPage)}
                >
                  <Scan size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.fitWidth')}
                  active={state.zoomMode === 'fit-width'}
                  ariaKeyShortcuts={pdfKeyboardShortcuts.fitWidth}
                  disabled={!can.fitWidth()}
                  onSelect={() => select(commands.fitWidth)}
                >
                  <MoveHorizontal size={15} />
                </PdfOverflowAction>
              </PdfOverflowGroup>
            )}
            {state.features.search && (
              <PdfOverflowGroup
                ariaLabel={officeMessage(messages, 'pdf.toolbar.overflowSearch')}
                className="work-pdf-overflow-group work-pdf-overflow-narrow"
              >
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.prevMatch')}
                  disabled={!can.previousSearchResult()}
                  onSelect={() => select(commands.previousSearchResult)}
                >
                  <ChevronUp size={15} />
                </PdfOverflowAction>
                <PdfOverflowAction
                  label={officeMessage(messages, 'pdf.toolbar.nextMatch')}
                  disabled={!can.nextSearchResult()}
                  onSelect={() => select(commands.nextSearchResult)}
                >
                  <ChevronDown size={15} />
                </PdfOverflowAction>
              </PdfOverflowGroup>
            )}
          </>
        );
      }}
    </Popover>
  );
}

function PdfOverflowGroup({
  ariaLabel,
  children,
  className = 'work-pdf-overflow-group',
}: {
  ariaLabel: string;
  children: ReactNode;
  className?: string;
}) {
  return (
    // Menu children must be menuitem* | group | separator — not fieldset.
    // biome-ignore lint/a11y/useSemanticElements: role="menu" forbids fieldset; group scopes menuitemradio sets.
    <div role="group" className={className} aria-label={ariaLabel}>
      {children}
    </div>
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

function isPdfActualSize(state: PdfViewerControllerState): boolean {
  return state.zoomMode === null && state.zoomPercent === 100;
}

function searchStatus(
  state: PdfViewerControllerState,
  messages: ReturnType<typeof useOfficeMessages>,
): string {
  const { search } = state;
  if (search.loading) {
    return officeMessage(messages, 'pdf.toolbar.searchLoading');
  }
  if (search.error) {
    return officeMessage(messages, 'pdf.toolbar.searchFailed');
  }
  if (!search.query && !search.active) return '';
  if (search.total === 0) return '0 / 0';
  return `${search.activeResultIndex + 1} / ${search.total}`;
}
