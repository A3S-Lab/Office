import type { Editor } from '@tiptap/core';
import { Grid2X2 } from 'lucide-react';
import { useState } from 'react';
import { Popover } from '../../../design-system/primitives';
import {
  DEFAULT_DOCUMENT_TABLE_CELL_MARGINS,
  type DocumentTableCellMarginSide,
  type DocumentTableCellMargins,
} from '../work-document-table-geometry';
import { documentTableSizing } from '../work-document-table-sizing';
import { OfficeNumberField } from './office-controls';

const PIXELS_PER_CENTIMETER = 96 / 2.54;
const marginFields = [
  { side: 'top', label: '上', ariaLabel: '单元格上边距（厘米）' },
  { side: 'bottom', label: '下', ariaLabel: '单元格下边距（厘米）' },
  { side: 'left', label: '左', ariaLabel: '单元格左边距（厘米）' },
  { side: 'right', label: '右', ariaLabel: '单元格右边距（厘米）' },
] as const satisfies readonly {
  side: DocumentTableCellMarginSide;
  label: string;
  ariaLabel: string;
}[];

export function DocumentTableMarginsPopover({ editor }: { editor: Editor }) {
  const margins = currentMargins(editor);
  const [open, setOpen] = useState(false);
  const [drafts, setDrafts] = useState(() => marginDrafts(margins));
  const customized = !sameMargins(margins, DEFAULT_DOCUMENT_TABLE_CELL_MARGINS);

  const commit = (side: DocumentTableCellMarginSide, rawValue: string) => {
    const centimeters = Number(rawValue);
    if (!Number.isFinite(centimeters) || centimeters < 0 || centimeters > 5) {
      setDrafts(marginDrafts(currentMargins(editor)));
      return;
    }
    const value = Math.round(centimeters * PIXELS_PER_CENTIMETER * 100) / 100;
    const next = { ...currentMargins(editor), [side]: value };
    if (editor.commands.setDocumentTableCellMargins(next)) {
      setDrafts(marginDrafts(next));
    } else {
      setDrafts(marginDrafts(currentMargins(editor)));
    }
  };

  const restoreDefault = () => {
    const next = { ...DEFAULT_DOCUMENT_TABLE_CELL_MARGINS };
    if (editor.commands.setDocumentTableCellMargins(next)) {
      setDrafts(marginDrafts(next));
    }
  };

  return (
    <Popover
      label="单元格边距"
      panelLabel="单元格边距设置"
      panelRole="dialog"
      portal
      open={open}
      onOpenChange={(nextOpen) => {
        setOpen(nextOpen);
        if (nextOpen) setDrafts(marginDrafts(currentMargins(editor)));
      }}
      focusFirstOnOpen
      className="work-document-table-margins-popover"
      panelClassName="work-document-table-margins-panel"
      trigger={(triggerProps, { open: popoverOpen }) => (
        <button
          {...triggerProps}
          className={`with-label${customized || popoverOpen ? ' active' : ''}`}
          aria-pressed={customized}
        >
          <Grid2X2 size={18} />
          <span>单元格边距</span>
        </button>
      )}
    >
      <fieldset>
        <legend>单元格边距</legend>
        <p>设置文字到单元格边框的距离。</p>
        <div className="work-document-table-margins-grid">
          {marginFields.map(({ side, label, ariaLabel }) => (
            <fieldset key={side} className="work-document-table-margin-field">
              <legend className="sr-only">{label}边距</legend>
              <span aria-hidden="true">{label}</span>
              <OfficeNumberField
                ariaLabel={ariaLabel}
                value={drafts[side]}
                min={0}
                max={5}
                step={0.05}
                onValueChange={(value) =>
                  setDrafts((current) => ({ ...current, [side]: value }))
                }
                onCommit={(value) => commit(side, value)}
                onCancel={() => setDrafts(marginDrafts(currentMargins(editor)))}
              />
              <small>厘米</small>
            </fieldset>
          ))}
        </div>
        <button
          type="button"
          className="work-document-table-margins-reset"
          disabled={!customized}
          onClick={restoreDefault}
        >
          恢复标准边距
        </button>
      </fieldset>
    </Popover>
  );
}

function currentMargins(editor: Editor): DocumentTableCellMargins {
  return (
    documentTableSizing(editor.state)?.cellMargins ?? {
      ...DEFAULT_DOCUMENT_TABLE_CELL_MARGINS,
    }
  );
}

function marginDrafts(
  margins: DocumentTableCellMargins,
): Record<DocumentTableCellMarginSide, string> {
  return {
    top: centimeters(margins.top),
    right: centimeters(margins.right),
    bottom: centimeters(margins.bottom),
    left: centimeters(margins.left),
  };
}

function centimeters(pixels: number): string {
  return (pixels / PIXELS_PER_CENTIMETER).toFixed(2).replace(/\.?0+$/, '');
}

function sameMargins(
  left: DocumentTableCellMargins,
  right: DocumentTableCellMargins,
): boolean {
  return marginFields.every(
    ({ side }) => Math.abs(left[side] - right[side]) < 0.01,
  );
}
