import type { PluginRegistry } from '@embedpdf/react-pdf-viewer';
import {
  ChevronLeft,
  ChevronRight,
  Download,
  FilePlus2,
  Files,
  RotateCcw,
  RotateCw,
  Scissors,
  Trash2,
} from 'lucide-react';
import { type MouseEvent, useMemo, useRef, useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import type {
  PdfEditorCanCommands,
  PdfEditorCommands,
} from './pdf-editor-extensions';
import { reorderedPdfPageIndexes } from './pdf-page-organization';
import {
  calculatePdfThumbnailRange,
  usePdfThumbnailSource,
} from './pdf-thumbnail-rail';
import type {
  PdfPageOrganizationControllerError,
  PdfPageOrganizationControllerState,
} from './use-pdf-page-organization';
import { OfficeFileInput } from './office-controls';

const PDF_ORGANIZER_ITEM_HEIGHT = 184;
const PDF_ORGANIZER_VIEWPORT_HEIGHT = 552;

export interface PdfPageOrganizerDialogProps {
  busy: boolean;
  can: PdfEditorCanCommands;
  commands: PdfEditorCommands;
  currentPage: number;
  diagnostics: PdfPageOrganizationControllerState['diagnostics'];
  error: PdfPageOrganizationControllerError | null;
  registry: PluginRegistry;
  restoreFocusTarget: () => HTMLElement | null;
  totalPages: number;
  onClose: () => void;
  onDismissError: () => void;
}

export function PdfPageOrganizerDialog({
  busy,
  can,
  commands,
  currentPage,
  diagnostics,
  error,
  registry,
  restoreFocusTarget,
  totalPages,
  onClose,
  onDismissError,
}: PdfPageOrganizerDialogProps) {
  const initialIndex = Math.min(
    Math.max(0, totalPages - 1),
    Math.max(0, currentPage - 1),
  );
  const [selected, setSelected] = useState<Set<number>>(
    () => new Set([initialIndex]),
  );
  const [anchorIndex, setAnchorIndex] = useState(initialIndex);
  const [submitting, setSubmitting] = useState(false);
  const selectionAnchorRef = useRef(initialIndex);
  const mergeInputRef = useRef<HTMLInputElement>(null);
  const effectiveBusy = busy || submitting;
  const selectedIndexes = useMemo(
    () => [...selected].sort((left, right) => left - right),
    [selected],
  );
  const range = useMemo(
    () =>
      calculatePdfThumbnailRange({
        anchorIndex,
        itemHeight: PDF_ORGANIZER_ITEM_HEIGHT,
        totalPages,
        viewportHeight: PDF_ORGANIZER_VIEWPORT_HEIGHT,
      }),
    [anchorIndex, totalPages],
  );
  const visibleIndexes = useMemo(
    () =>
      Array.from(
        { length: Math.max(0, range.end - range.start) },
        (_, index) => range.start + index,
      ),
    [range.end, range.start],
  );
  const insertionIndex =
    selectedIndexes.length > 0
      ? Math.min(totalPages, (selectedIndexes.at(-1) ?? 0) + 1)
      : totalPages;
  const splitBoundaries = selectedIndexes.filter(
    (pageIndex) => pageIndex < totalPages - 1,
  );
  const moveLeftOrder = pageMoveOrder(totalPages, selectedIndexes, 'left');
  const moveRightOrder = pageMoveOrder(totalPages, selectedIndexes, 'right');

  const commitMutation = async (
    execute: () => Promise<boolean>,
  ): Promise<void> => {
    if (effectiveBusy) return;
    setSubmitting(true);
    const applied = await execute();
    if (applied) onClose();
    else setSubmitting(false);
  };
  const runExport = async (execute: () => Promise<boolean>): Promise<void> => {
    if (effectiveBusy) return;
    setSubmitting(true);
    await execute();
    setSubmitting(false);
  };

  return (
    <Dialog
      title="组织 PDF 页面"
      description="选择页面后插入、删除、旋转、重排、抽取、合并或拆分。"
      className="work-pdf-page-organizer-dialog"
      closeDisabled={effectiveBusy}
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <Button tone="primary" disabled={effectiveBusy} onClick={onClose}>
          完成
        </Button>
      }
    >
      <OfficeFileInput
        ref={mergeInputRef}
        accept=".pdf,application/pdf"
        aria-label="选择要合并的 PDF"
        disabled={effectiveBusy}
        onFileSelect={(file) => {
          if (!can.mergePages(insertionIndex, file)) return;
          void commitMutation(() => commands.mergePages(insertionIndex, file));
        }}
      />

      <div
        className="work-pdf-page-organizer-actions"
        role="toolbar"
        aria-label="PDF 页面组织命令"
      >
        <Button
          tone="secondary"
          aria-label="插入空白页"
          disabled={!can.insertBlankPage(insertionIndex) || effectiveBusy}
          onClick={() =>
            void commitMutation(() => commands.insertBlankPage(insertionIndex))
          }
        >
          <FilePlus2 size={15} /> 插入空白页
        </Button>
        <Button
          tone="secondary"
          aria-label="合并另一个 PDF"
          disabled={effectiveBusy}
          onClick={() => mergeInputRef.current?.click()}
        >
          <Files size={15} /> 合并 PDF
        </Button>
        <Button
          tone="secondary"
          aria-label="前移所选页"
          disabled={
            !moveLeftOrder || !can.reorderPages(moveLeftOrder) || effectiveBusy
          }
          onClick={() => {
            if (moveLeftOrder) {
              void commitMutation(() => commands.reorderPages(moveLeftOrder));
            }
          }}
        >
          <ChevronLeft size={15} /> 前移
        </Button>
        <Button
          tone="secondary"
          aria-label="后移所选页"
          disabled={
            !moveRightOrder ||
            !can.reorderPages(moveRightOrder) ||
            effectiveBusy
          }
          onClick={() => {
            if (moveRightOrder) {
              void commitMutation(() => commands.reorderPages(moveRightOrder));
            }
          }}
        >
          <ChevronRight size={15} /> 后移
        </Button>
        <Button
          tone="secondary"
          aria-label="向左旋转所选页"
          disabled={!can.rotatePages(selectedIndexes, 270) || effectiveBusy}
          onClick={() =>
            void commitMutation(() =>
              commands.rotatePages(selectedIndexes, 270),
            )
          }
        >
          <RotateCcw size={15} /> 左转
        </Button>
        <Button
          tone="secondary"
          aria-label="向右旋转所选页"
          disabled={!can.rotatePages(selectedIndexes, 90) || effectiveBusy}
          onClick={() =>
            void commitMutation(() => commands.rotatePages(selectedIndexes, 90))
          }
        >
          <RotateCw size={15} /> 右转
        </Button>
        <Button
          tone="secondary"
          aria-label="抽取所选页"
          disabled={!can.extractPages(selectedIndexes) || effectiveBusy}
          onClick={() =>
            void runExport(() => commands.extractPages(selectedIndexes))
          }
        >
          <Download size={15} /> 抽取
        </Button>
        <Button
          tone="secondary"
          aria-label="在所选页后拆分"
          disabled={!can.splitPages(splitBoundaries) || effectiveBusy}
          onClick={() =>
            void runExport(() => commands.splitPages(splitBoundaries))
          }
        >
          <Scissors size={15} /> 拆分
        </Button>
        <Button
          tone="danger"
          aria-label="删除所选页"
          disabled={!can.deletePages(selectedIndexes) || effectiveBusy}
          onClick={() =>
            void commitMutation(() => commands.deletePages(selectedIndexes))
          }
        >
          <Trash2 size={15} /> 删除
        </Button>
      </div>

      <div className="work-pdf-page-organizer-selection">
        <output aria-live="polite">
          已选择 {selectedIndexes.length} / {totalPages} 页
        </output>
        <button
          type="button"
          disabled={effectiveBusy}
          onClick={() =>
            setSelected(
              new Set(Array.from({ length: totalPages }, (_, index) => index)),
            )
          }
        >
          全选
        </button>
        <button
          type="button"
          disabled={effectiveBusy}
          onClick={() => setSelected(new Set([initialIndex]))}
        >
          仅当前页
        </button>
      </div>

      <section
        className="work-pdf-page-organizer-viewport"
        aria-label="可重排 PDF 页面"
        onScroll={(event) =>
          setAnchorIndex(
            Math.floor(
              event.currentTarget.scrollTop / PDF_ORGANIZER_ITEM_HEIGHT,
            ),
          )
        }
      >
        <div className="work-pdf-page-organizer-list">
          <OrganizerSpacer height={range.start * PDF_ORGANIZER_ITEM_HEIGHT} />
          {visibleIndexes.map((pageIndex) => (
            <OrganizerPage
              current={pageIndex === initialIndex}
              key={pageIndex}
              pageIndex={pageIndex}
              registry={registry}
              selected={selected.has(pageIndex)}
              disabled={effectiveBusy}
              onDrop={(targetIndex) => {
                const order = reorderedPdfPageIndexes(
                  totalPages,
                  selectedIndexes,
                  targetIndex,
                );
                if (can.reorderPages(order)) {
                  void commitMutation(() => commands.reorderPages(order));
                }
              }}
              onSelect={(event) => {
                setSelected((current) =>
                  nextPageSelection(
                    current,
                    pageIndex,
                    selectionAnchorRef.current,
                    event,
                  ),
                );
                if (!event.shiftKey) selectionAnchorRef.current = pageIndex;
              }}
            />
          ))}
          <OrganizerSpacer
            height={
              Math.max(0, totalPages - range.end) * PDF_ORGANIZER_ITEM_HEIGHT
            }
          />
        </div>
      </section>

      {error && (
        <div className="work-pdf-page-organizer-error" role="alert">
          <div>
            <strong>无法安全完成页面操作</strong>
            <p>{error.message}</p>
            <code>{error.code}</code>
          </div>
          <button
            type="button"
            aria-label="关闭错误提示"
            onClick={onDismissError}
          >
            ×
          </button>
        </div>
      )}
      {diagnostics.length > 0 && (
        <ul className="work-pdf-page-organizer-diagnostics" role="status">
          {diagnostics.map((diagnostic) => (
            <li key={diagnostic.code}>{diagnostic.message}</li>
          ))}
        </ul>
      )}
      <p className="work-pdf-page-organizer-boundary">
        页面操作在 Worker 中生成新
        PDF，并由一个撤销记录恢复。签名、加密、表单、目录或标记结构无法安全重写时会明确停止。
      </p>
    </Dialog>
  );
}

function OrganizerPage({
  current,
  disabled,
  pageIndex,
  registry,
  selected,
  onDrop,
  onSelect,
}: {
  current: boolean;
  disabled: boolean;
  pageIndex: number;
  registry: PluginRegistry;
  selected: boolean;
  onDrop: (targetIndex: number) => void;
  onSelect: (event: MouseEvent<HTMLButtonElement>) => void;
}) {
  const { sourceUrl, state } = usePdfThumbnailSource(registry, pageIndex + 1);
  return (
    <button
      type="button"
      className={selected ? 'selected' : undefined}
      aria-current={current ? 'page' : undefined}
      aria-label={`选择第 ${pageIndex + 1} 页`}
      aria-pressed={selected}
      data-pdf-organizer-page-index={pageIndex}
      disabled={disabled}
      draggable={!disabled}
      onClick={onSelect}
      onDragOver={(event) => {
        event.preventDefault();
        event.dataTransfer.dropEffect = 'move';
      }}
      onDragStart={(event) => {
        event.dataTransfer.effectAllowed = 'move';
        event.dataTransfer.setData('text/plain', String(pageIndex));
      }}
      onDrop={(event) => {
        event.preventDefault();
        onDrop(pageIndex);
      }}
    >
      <span className="work-pdf-page-organizer-number">
        {pageIndex + 1}
        {current && <small>当前</small>}
      </span>
      <span className="work-pdf-page-organizer-preview" data-state={state}>
        {sourceUrl && <img src={sourceUrl} alt="" draggable={false} />}
      </span>
    </button>
  );
}

function nextPageSelection(
  current: Set<number>,
  pageIndex: number,
  anchorIndex: number,
  event: MouseEvent<HTMLButtonElement>,
): Set<number> {
  if (event.shiftKey) {
    const start = Math.min(anchorIndex, pageIndex);
    const end = Math.max(anchorIndex, pageIndex);
    return new Set(
      Array.from({ length: end - start + 1 }, (_, index) => start + index),
    );
  }
  if (event.metaKey || event.ctrlKey) {
    const next = new Set(current);
    if (next.has(pageIndex)) next.delete(pageIndex);
    else next.add(pageIndex);
    return next;
  }
  return new Set([pageIndex]);
}

function pageMoveOrder(
  pageCount: number,
  selected: number[],
  direction: 'left' | 'right',
): number[] | null {
  if (selected.length === 0) return null;
  if (direction === 'left') {
    const first = selected[0] ?? 0;
    return first <= 0
      ? null
      : reorderedPdfPageIndexes(pageCount, selected, first - 1);
  }
  const last = selected.at(-1) ?? pageCount - 1;
  return last >= pageCount - 1
    ? null
    : reorderedPdfPageIndexes(pageCount, selected, last + 2);
}

function OrganizerSpacer({ height }: { height: number }) {
  if (height <= 0) return null;
  return <span aria-hidden="true" style={{ height }} />;
}
