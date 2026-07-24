import {
  Check,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  ChevronUp,
  Highlighter,
  Loader2,
  Minus,
  MousePointer2,
  Pencil,
  Plus,
  Redo2,
  Save,
  Search,
  Strikethrough,
  Trash2,
  Type,
  Undo2,
  Underline,
  X,
} from 'lucide-react';
import {
  type FormEvent,
  type KeyboardEvent,
  type RefObject,
  useEffect,
  useState,
} from 'react';
import {
  Button,
  IconButton,
  StatusBadge,
} from '../../../design-system/primitives';
import { OfficeTextField } from './office-controls';
import type { PdfAnnotationController } from './pdf-annotation-controller';
import type { PdfViewerController } from './pdf-viewer-controller';

export type PdfSaveState = 'idle' | 'saving' | 'saved' | 'error';

export function PdfToolbar({
  annotation,
  controller,
  editable,
  onSave,
  saveLabel,
  saveState,
  searchInputRef,
}: {
  annotation: PdfAnnotationController;
  controller: PdfViewerController;
  editable: boolean;
  onSave?: () => void;
  saveLabel: string;
  saveState: PdfSaveState;
  searchInputRef: RefObject<HTMLInputElement | null>;
}) {
  const { state } = controller;
  const [pageValue, setPageValue] = useState('');
  const [searchValue, setSearchValue] = useState('');
  const documentReady = state.ready && state.documentOpen;

  useEffect(() => {
    setPageValue(state.currentPage > 0 ? String(state.currentPage) : '');
  }, [state.currentPage]);

  useEffect(() => {
    setSearchValue(state.search.query);
  }, [state.search.query]);

  const commitPage = () => {
    const page = Number(pageValue);
    if (Number.isInteger(page) && page >= 1 && page <= state.totalPages) {
      controller.goToPage(page);
      return;
    }
    setPageValue(state.currentPage > 0 ? String(state.currentPage) : '');
  };

  const submitSearch = (event: FormEvent) => {
    event.preventDefault();
    const query = searchValue.trim();
    if (
      query &&
      query === state.search.query &&
      state.search.total > 0 &&
      !state.search.loading
    ) {
      controller.nextSearchResult();
      return;
    }
    controller.search(query);
  };

  return (
    <header className="work-pdf-toolbar" role="toolbar" aria-label="PDF 工具栏">
      <div className="work-pdf-toolbar-group work-pdf-history">
        <IconButton
          label="撤销"
          disabled={!documentReady || !state.canUndo}
          onClick={controller.undo}
        >
          <Undo2 size={15} />
        </IconButton>
        <IconButton
          label="重做"
          disabled={!documentReady || !state.canRedo}
          onClick={controller.redo}
        >
          <Redo2 size={15} />
        </IconButton>
      </div>

      {editable && (
        <fieldset className="work-pdf-toolbar-group work-pdf-annotation">
          <legend className="sr-only">PDF 批注工具</legend>
          <IconButton
            label="选择"
            selected={annotation.state.activeToolId === null}
            disabled={!documentReady || !annotation.state.available}
            onClick={() => annotation.selectTool(null)}
          >
            <MousePointer2 size={14} />
          </IconButton>
          <IconButton
            label="高亮"
            selected={annotation.state.activeToolId === 'highlight'}
            disabled={!documentReady || !annotation.state.available}
            onClick={() => annotation.selectTool('highlight')}
          >
            <Highlighter size={14} />
          </IconButton>
          <IconButton
            className="work-pdf-annotation-optional"
            label="下划线批注"
            selected={annotation.state.activeToolId === 'underline'}
            disabled={!documentReady || !annotation.state.available}
            onClick={() => annotation.selectTool('underline')}
          >
            <Underline size={14} />
          </IconButton>
          <IconButton
            className="work-pdf-annotation-optional"
            label="删除线批注"
            selected={annotation.state.activeToolId === 'strikeout'}
            disabled={!documentReady || !annotation.state.available}
            onClick={() => annotation.selectTool('strikeout')}
          >
            <Strikethrough size={14} />
          </IconButton>
          <IconButton
            label="画笔"
            selected={annotation.state.activeToolId === 'ink'}
            disabled={!documentReady || !annotation.state.available}
            onClick={() => annotation.selectTool('ink')}
          >
            <Pencil size={14} />
          </IconButton>
          <IconButton
            className="work-pdf-annotation-optional"
            label="文字批注"
            selected={annotation.state.activeToolId === 'freeText'}
            disabled={!documentReady || !annotation.state.available}
            onClick={() => annotation.selectTool('freeText')}
          >
            <Type size={14} />
          </IconButton>
          <IconButton
            label="删除所选批注"
            disabled={
              !documentReady ||
              !annotation.state.available ||
              annotation.state.selectedCount === 0
            }
            onClick={annotation.deleteSelection}
          >
            <Trash2 size={14} />
          </IconButton>
        </fieldset>
      )}

      <search className="work-pdf-search">
        <form onSubmit={submitSearch}>
          <Search size={14} aria-hidden="true" />
          <OfficeTextField
            ref={searchInputRef}
            type="search"
            aria-label="在 PDF 中搜索"
            placeholder="搜索"
            value={searchValue}
            disabled={!documentReady || !state.features.search}
            onChange={(event) => setSearchValue(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Escape') {
                event.preventDefault();
                setSearchValue('');
                controller.clearSearch();
              } else if (
                event.key === 'Enter' &&
                event.shiftKey &&
                searchValue.trim() === state.search.query &&
                state.search.total > 0
              ) {
                event.preventDefault();
                controller.previousSearchResult();
              }
            }}
          />
          {(searchValue || state.search.active) && (
            <IconButton
              className="work-pdf-search-clear"
              label="清除搜索"
              onClick={() => {
                setSearchValue('');
                controller.clearSearch();
                searchInputRef.current?.focus();
              }}
            >
              <X size={13} />
            </IconButton>
          )}
          <output className="work-pdf-search-state" aria-live="polite">
            {searchStatus(controller)}
          </output>
          <IconButton
            label="上一个搜索结果"
            disabled={state.search.total === 0}
            onClick={controller.previousSearchResult}
          >
            <ChevronUp size={14} />
          </IconButton>
          <IconButton
            label="下一个搜索结果"
            disabled={state.search.total === 0}
            onClick={controller.nextSearchResult}
          >
            <ChevronDown size={14} />
          </IconButton>
        </form>
      </search>

      <div className="work-pdf-toolbar-group work-pdf-page-controls">
        <IconButton
          label="上一页"
          disabled={
            !documentReady ||
            !state.features.navigation ||
            state.currentPage <= 1
          }
          onClick={controller.previousPage}
        >
          <ChevronLeft size={15} />
        </IconButton>
        <OfficeTextField
          className="work-pdf-page-field"
          aria-label="页码"
          inputMode="numeric"
          value={pageValue}
          disabled={!documentReady || !state.features.navigation}
          onBlur={commitPage}
          onChange={(event) =>
            setPageValue(event.target.value.replace(/\D/g, ''))
          }
          onKeyDown={(event: KeyboardEvent<HTMLInputElement>) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              commitPage();
              event.currentTarget.select();
            } else if (event.key === 'Escape') {
              setPageValue(String(state.currentPage));
              event.currentTarget.blur();
            }
          }}
        />
        <span className="work-pdf-page-total">/ {state.totalPages || '—'}</span>
        <IconButton
          label="下一页"
          disabled={
            !documentReady ||
            !state.features.navigation ||
            state.currentPage >= state.totalPages
          }
          onClick={controller.nextPage}
        >
          <ChevronRight size={15} />
        </IconButton>
      </div>

      <div className="work-pdf-toolbar-group work-pdf-zoom-controls">
        <IconButton
          label="缩小"
          disabled={!documentReady || !state.features.zoom}
          onClick={controller.zoomOut}
        >
          <Minus size={14} />
        </IconButton>
        <output aria-label="PDF 缩放比例">{state.zoomPercent}%</output>
        <IconButton
          label="放大"
          disabled={!documentReady || !state.features.zoom}
          onClick={controller.zoomIn}
        >
          <Plus size={14} />
        </IconButton>
        <button
          type="button"
          className="work-pdf-fit-button"
          aria-pressed={state.zoomMode === 'fit-page'}
          disabled={!documentReady || !state.features.zoom}
          onClick={controller.fitPage}
        >
          整页
        </button>
        <button
          type="button"
          className="work-pdf-fit-button"
          aria-pressed={state.zoomMode === 'fit-width'}
          disabled={!documentReady || !state.features.zoom}
          onClick={controller.fitWidth}
        >
          页宽
        </button>
      </div>

      {onSave && (
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
            disabled={
              !documentReady || !state.features.export || saveState === 'saving'
            }
            onClick={onSave}
          >
            <Save size={14} />
            {saveLabel}
          </Button>
        </div>
      )}
    </header>
  );
}

function searchStatus(controller: PdfViewerController): string {
  const { search } = controller.state;
  if (search.loading) return '搜索中';
  if (search.error) return '失败';
  if (!search.query && !search.active) return '';
  if (search.total === 0) return '0 / 0';
  return `${search.activeResultIndex + 1} / ${search.total}`;
}
