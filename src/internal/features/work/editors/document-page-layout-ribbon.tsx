import type { Editor } from '@tiptap/core';
import { Columns3, FilePlus2, Palette, Settings2 } from 'lucide-react';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import { normalizeDocumentColumns } from '../work-document-columns';
import {
  updateDocumentPageOrientation,
  updateDocumentPaperSizePreset,
} from '../work-document-page-size';
import type {
  WorkDocumentMargins,
  WorkDocumentSectionLayout,
} from '../work-types';
import { getDocumentCommandDefinition } from './document-command-catalog';
import {
  documentCommandLabel,
  documentCommandTitleWithShortcut,
} from './document-command-i18n';
import type { DocumentLayoutPanelTab } from './document-layout-panel';
import { DocumentPaginationPopover } from './document-pagination-popover';
import { DocumentParagraphSpacingPopover } from './document-paragraph-spacing-popover';
import { OfficeColorPicker, OfficeSelect } from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

type DocumentMarginPreset =
  | 'custom'
  | 'moderate'
  | 'narrow'
  | 'normal'
  | 'wide';
type DocumentColumnCount = '1' | '2' | '3' | '4' | '5' | '6';
type DocumentColumnPreset = DocumentColumnCount | 'more';

const documentMarginPresets = {
  normal: { top: 25.4, right: 25.4, bottom: 25.4, left: 25.4 },
  narrow: { top: 12.7, right: 12.7, bottom: 12.7, left: 12.7 },
  moderate: { top: 25.4, right: 19.1, bottom: 25.4, left: 19.1 },
  wide: { top: 25.4, right: 50.8, bottom: 25.4, left: 50.8 },
} as const satisfies Record<
  Exclude<DocumentMarginPreset, 'custom'>,
  WorkDocumentMargins
>;

export function DocumentPageLayoutRibbon({
  editor,
  layout,
  layoutOpen,
  pageColor,
  onLayoutChange,
  onOpenLayout,
  onToggleLayout,
  onPageColorChange,
  onInsertSection,
}: {
  editor: Editor;
  layout: WorkDocumentSectionLayout;
  layoutOpen: boolean;
  pageColor: string;
  onLayoutChange: (layout: WorkDocumentSectionLayout) => void;
  onOpenLayout: (target: DocumentLayoutPanelTab) => void;
  onToggleLayout: () => void;
  onPageColorChange: (color: string) => void;
  onInsertSection: () => void;
}) {
  const messages = useOfficeMessages();
  const pageBreakCommand = getDocumentCommandDefinition('insertPageBreak');
  const marginPreset = documentMarginPreset(layout.margins);
  const columnPreset = documentColumnClosedValue(layout.columns.count);
  const update = (patch: Partial<WorkDocumentSectionLayout>) =>
    onLayoutChange({ ...layout, ...patch });

  return (
    <>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.pageLayout.group.setup')}
        priority="high"
      >
        <div className="work-office-field work-document-page-setup-choice">
          <span>
            {officeMessage(messages, 'document.pageLayout.margins')}
          </span>
          <OfficeSelect
            ariaLabel={officeMessage(
              messages,
              'document.pageLayout.marginsAria',
            )}
            value={marginPreset}
            options={marginPresetOptions(messages)}
            onValueChange={(preset) => {
              if (preset === 'custom') {
                onOpenLayout('page');
                return;
              }
              update({ margins: { ...documentMarginPresets[preset] } });
            }}
          />
        </div>
        <div className="work-office-field work-document-page-setup-choice">
          <span>
            {officeMessage(messages, 'document.layout.paper.orientation')}
          </span>
          <OfficeSelect
            ariaLabel={officeMessage(
              messages,
              'document.layout.paper.orientationAria',
            )}
            value={layout.orientation}
            options={[
              {
                value: 'portrait',
                label: officeMessage(
                  messages,
                  'document.layout.paper.portrait',
                ),
              },
              {
                value: 'landscape',
                label: officeMessage(
                  messages,
                  'document.layout.paper.landscape',
                ),
              },
            ]}
            onValueChange={(orientation) =>
              onLayoutChange(updateDocumentPageOrientation(layout, orientation))
            }
          />
        </div>
        <div className="work-office-field work-document-page-setup-choice">
          <span>
            {officeMessage(messages, 'document.layout.paper.heading')}
          </span>
          <OfficeSelect
            ariaLabel={officeMessage(
              messages,
              'document.layout.paper.sizeAria',
            )}
            value={layout.pageSize}
            options={[
              { value: 'a3', label: 'A3' },
              { value: 'a4', label: 'A4' },
              { value: 'a5', label: 'A5' },
              { value: 'letter', label: 'Letter' },
              { value: 'legal', label: 'Legal' },
              { value: 'tabloid', label: 'Tabloid' },
              {
                value: 'custom',
                label: officeMessage(messages, 'document.layout.paper.custom'),
              },
            ]}
            onValueChange={(pageSize) => {
              if (pageSize === 'custom') {
                onOpenLayout('page');
                return;
              }
              onLayoutChange(updateDocumentPaperSizePreset(layout, pageSize));
            }}
          />
        </div>
        <div className="work-office-field work-document-page-setup-choice">
          <span>
            {officeMessage(messages, 'document.pageLayout.columns')}
          </span>
          <OfficeSelect<DocumentColumnPreset>
            ariaLabel={officeMessage(
              messages,
              'document.pageLayout.columnsAria',
            )}
            value={columnPreset}
            options={[
              ...documentColumnCountOptions(messages),
              {
                value: 'more',
                label: officeMessage(
                  messages,
                  'document.pageLayout.columns.more',
                ),
              },
            ]}
            onValueChange={(preset) => {
              if (preset === 'more') {
                onOpenLayout('columns');
                return;
              }
              update({
                columns: normalizeDocumentColumns({
                  ...layout.columns,
                  count: Number(preset),
                  custom: undefined,
                }),
              });
            }}
          />
        </div>
        <WorkOfficeRibbonButton
          label={officeMessage(messages, 'document.pageLayout.setup')}
          active={layoutOpen}
          onClick={onToggleLayout}
        >
          <Settings2 size={19} />
        </WorkOfficeRibbonButton>
        <WorkOfficeRibbonButton
          label={documentCommandLabel('insertPageBreak', messages)}
          title={documentCommandTitleWithShortcut('insertPageBreak', messages)}
          aria-keyshortcuts={pageBreakCommand.shortcut?.aria}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => {
            // Keep ribbon focus on Page Layout page-break. chain().focus()
            // schedules into the editor and breaks L2 loops.
            editor.commands.insertContent({ type: 'pageBreak' });
          }}
        >
          <FilePlus2 size={19} />
        </WorkOfficeRibbonButton>
        <WorkOfficeRibbonButton
          label={officeMessage(messages, 'document.pageLayout.insertSection')}
          onClick={onInsertSection}
        >
          <Columns3 size={19} />
        </WorkOfficeRibbonButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.pageLayout.group.paragraph')}
        priority="high"
      >
        <DocumentParagraphSpacingPopover editor={editor} />
        <DocumentPaginationPopover editor={editor} />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.pageLayout.group.background')}
        priority="low"
      >
        <OfficeColorPicker
          ariaLabel={officeMessage(
            messages,
            'document.pageLayout.pageColorAria',
          )}
          className="work-document-page-color-picker"
          triggerLabel={officeMessage(
            messages,
            'document.pageLayout.pageColorTrigger',
          )}
          triggerIcon={<Palette size={18} />}
          value={pageColor}
          onValueChange={onPageColorChange}
        />
      </WorkOfficeRibbonGroup>
    </>
  );
}

function documentColumnCountOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: '1' as const,
      label: officeMessage(messages, 'document.pageLayout.columns.count1'),
    },
    {
      value: '2' as const,
      label: officeMessage(messages, 'document.pageLayout.columns.count2'),
    },
    {
      value: '3' as const,
      label: officeMessage(messages, 'document.pageLayout.columns.count3'),
    },
    {
      value: '4' as const,
      label: officeMessage(messages, 'document.pageLayout.columns.count4'),
    },
    {
      value: '5' as const,
      label: officeMessage(messages, 'document.pageLayout.columns.count5'),
    },
    {
      value: '6' as const,
      label: officeMessage(messages, 'document.pageLayout.columns.count6'),
    },
  ];
}

function marginPresetOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'normal' as const,
      label: officeMessage(messages, 'document.pageLayout.margin.normal'),
      meta: officeMessage(messages, 'document.pageLayout.margin.normalMeta'),
    },
    {
      value: 'narrow' as const,
      label: officeMessage(messages, 'document.pageLayout.margin.narrow'),
      meta: officeMessage(messages, 'document.pageLayout.margin.narrowMeta'),
    },
    {
      value: 'moderate' as const,
      label: officeMessage(messages, 'document.pageLayout.margin.moderate'),
      meta: officeMessage(messages, 'document.pageLayout.margin.moderateMeta'),
    },
    {
      value: 'wide' as const,
      label: officeMessage(messages, 'document.pageLayout.margin.wide'),
      meta: officeMessage(messages, 'document.pageLayout.margin.wideMeta'),
    },
    {
      value: 'custom' as const,
      label: officeMessage(messages, 'document.pageLayout.margin.custom'),
      meta: officeMessage(messages, 'document.pageLayout.margin.customMeta'),
    },
  ];
}

function documentColumnClosedValue(count: number): DocumentColumnCount {
  const clamped = Math.min(6, Math.max(1, Math.round(count) || 1));
  return String(clamped) as DocumentColumnCount;
}

function documentMarginPreset(
  margins: WorkDocumentMargins,
): DocumentMarginPreset {
  for (const [preset, values] of Object.entries(documentMarginPresets)) {
    if (
      Object.keys(values).every(
        (side) =>
          Math.abs(
            margins[side as keyof WorkDocumentMargins] -
              values[side as keyof WorkDocumentMargins],
          ) < 0.05,
      )
    ) {
      return preset as DocumentMarginPreset;
    }
  }
  return 'custom';
}
