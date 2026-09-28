import type { PluginRegistry } from '@embedpdf/react-pdf-viewer';
import { officeMessage } from '../../../i18n/office-locale';
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
import { useOfficeMessages } from './office-messages-context';
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
import { moveOfficeToolbarFocus } from './office-toolbar-keyboard';

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
  const messages = useOfficeMessages();
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
      title={officeMessage(messages, 'pdf.organizer.title')}
      description={officeMessage(messages, 'pdf.organizer.description')}
      className="work-pdf-page-organizer-dialog"
      closeDisabled={effectiveBusy}
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <Button tone="primary" disabled={effectiveBusy} onClick={onClose}>
          {officeMessage(messages, 'pdf.organizer.done')}
        </Button>
      }
    >
      <OfficeFileInput
        ref={mergeInputRef}
        accept=".pdf,application/pdf"
        aria-label={officeMessage(messages, 'pdf.organizer.mergeInputAria')}
        disabled={effectiveBusy}
        onFileSelect={(file) => {
          if (!can.mergePages(insertionIndex, file)) return;
          void commitMutation(() => commands.mergePages(insertionIndex, file));
        }}
      />

      <div
        className="work-pdf-page-organizer-actions"
        role="toolbar"
        aria-label={officeMessage(messages, 'pdf.organizer.commandsAria')}
        onKeyDown={moveOfficeToolbarFocus}
      >
        <Button
          tone="secondary"
          aria-label={officeMessage(messages, 'pdf.organizer.insertBlankAria')}
          disabled={!can.insertBlankPage(insertionIndex) || effectiveBusy}
          onClick={() =>
            void commitMutation(() => commands.insertBlankPage(insertionIndex))
          }
        >
          <FilePlus2 size={15} /> {officeMessage(messages, 'pdf.organizer.insertBlank')}
        </Button>
        <Button
          tone="secondary"
          aria-label={officeMessage(messages, 'pdf.organizer.mergeAria')}
          disabled={effectiveBusy}
          onClick={() => mergeInputRef.current?.click()}
        >
          <Files size={15} /> {officeMessage(messages, 'pdf.organizer.merge')}
        </Button>
        <Button
          tone="secondary"
          aria-label={officeMessage(messages, 'pdf.organizer.moveEarlierAria')}
          disabled={
            !moveLeftOrder || !can.reorderPages(moveLeftOrder) || effectiveBusy
          }
          onClick={() => {
            if (moveLeftOrder) {
              void commitMutation(() => commands.reorderPages(moveLeftOrder));
            }
          }}
        >
          <ChevronLeft size={15} /> {officeMessage(messages, 'pdf.organizer.moveEarlier')}
        </Button>
        <Button
          tone="secondary"
          aria-label={officeMessage(messages, 'pdf.organizer.moveLaterAria')}
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
          <ChevronRight size={15} /> {officeMessage(messages, 'pdf.organizer.moveLater')}
        </Button>
        <Button
          tone="secondary"
          aria-label={officeMessage(messages, 'pdf.organizer.rotateLeftAria')}
          disabled={!can.rotatePages(selectedIndexes, 270) || effectiveBusy}
          onClick={() =>
            void commitMutation(() =>
              commands.rotatePages(selectedIndexes, 270),
            )
          }
        >
          <RotateCcw size={15} /> {officeMessage(messages, 'pdf.organizer.rotateLeft')}
        </Button>
        <Button
          tone="secondary"
          aria-label={officeMessage(messages, 'pdf.organizer.rotateRightAria')}
          disabled={!can.rotatePages(selectedIndexes, 90) || effectiveBusy}
          onClick={() =>
            void commitMutation(() => commands.rotatePages(selectedIndexes, 90))
          }
        >
          <RotateCw size={15} /> {officeMessage(messages, 'pdf.organizer.rotateRight')}
        </Button>
        <Button
          tone="secondary"
          aria-label={officeMessage(messages, 'pdf.organizer.extractAria')}
          disabled={!can.extractPages(selectedIndexes) || effectiveBusy}
          onClick={() =>
            void runExport(() => commands.extractPages(selectedIndexes))
          }
        >
          <Download size={15} /> {officeMessage(messages, 'pdf.organizer.extract')}
        </Button>
        <Button
          tone="secondary"
          aria-label={officeMessage(messages, 'pdf.organizer.splitAria')}
          disabled={!can.splitPages(splitBoundaries) || effectiveBusy}
          onClick={() =>
            void runExport(() => commands.splitPages(splitBoundaries))
          }
        >
          <Scissors size={15} /> {officeMessage(messages, 'pdf.organizer.split')}
        </Button>
        <Button
          tone="danger"
          aria-label={officeMessage(messages, 'pdf.organizer.deleteAria')}
          disabled={!can.deletePages(selectedIndexes) || effectiveBusy}
          onClick={() =>
            void commitMutation(() => commands.deletePages(selectedIndexes))
          }
        >
          <Trash2 size={15} /> {officeMessage(messages, 'pdf.organizer.delete')}
        </Button>
      </div>

      <div className="work-pdf-page-organizer-selection">
        <output aria-live="polite">
          {officeMessage(messages, 'pdf.organizer.selectedCount', { selected: String(selectedIndexes.length), total: String(totalPages) })}
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
          {officeMessage(messages, 'pdf.organizer.selectAll')}
        </button>
        <button
          type="button"
          disabled={effectiveBusy}
          onClick={() => setSelected(new Set([initialIndex]))}
        >
          {officeMessage(messages, 'pdf.organizer.currentOnly')}
        </button>
      </div>

      <section
        className="work-pdf-page-organizer-viewport"
        aria-label={officeMessage(messages, 'pdf.organizer.listAria')}
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
            <strong>{officeMessage(messages, 'pdf.organizer.errorTitle')}</strong>
            <p>{error.message}</p>
            <code>{error.code}</code>
          </div>
          <button
            type="button"
            aria-label={officeMessage(messages, 'pdf.organizer.closeError')}
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
        {officeMessage(messages, 'pdf.organizer.workerHint')}
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
  const messages = useOfficeMessages();
  const { sourceUrl, state } = usePdfThumbnailSource(registry, pageIndex + 1);
  return (
    <button
      type="button"
      className={selected ? 'selected' : undefined}
      aria-current={current ? 'page' : undefined}
      aria-label={officeMessage(messages, 'pdf.organizer.selectPageAria', { page: String(pageIndex + 1) })}
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
        {current && <small>{officeMessage(messages, 'pdf.organizer.currentBadge')}</small>}
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
