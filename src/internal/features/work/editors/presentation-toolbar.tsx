import {
  AlignCenter,
  AlignHorizontalSpaceBetween,
  AlignLeft,
  AlignRight,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  AlignVerticalSpaceBetween,
  ArrowDownToLine,
  ArrowUpToLine,
  BarChart3,
  Bold,
  ChevronDown,
  ClipboardPaste,
  Copy,
  Grid2X2,
  Group,
  Image,
  Italic,
  LayoutTemplate,
  Link2,
  MessageSquarePlus,
  MessagesSquare,
  NotebookPen,
  PanelsTopLeft,
  Play,
  Plus,
  Redo2,
  Scissors,
  Square,
  SquarePlay,
  Trash2,
  Type,
  Underline,
  Undo2,
  Ungroup,
} from 'lucide-react';
import { useEffect, useRef, useState } from 'react';
import { Popover } from '../../../design-system/primitives';
import type { OfficeKernelPresentationAlignment } from '../../../kernel/office-kernel-protocol';
import {
  DOCUMENT_LINK_VALIDATION_MESSAGE,
  normalizeDocumentHref,
} from '../work-document-links';
import { workSlideAnimationForElement } from '../work-presentation-animation';
import type {
  WorkSlide,
  WorkSlideElement,
  WorkSlideTextAlign,
} from '../work-types';
import {
  OfficeColorPicker,
  OfficeNumberField,
  OfficeSelect,
  useOfficeDialog,
} from './office-controls';
import {
  normalizeOfficeFontFamily,
  officeFontFamilies,
  officeFontFamilyGroupLabel,
  officeFontFamilyLabel,
  officeFontFamilyLocalizedLabel,
} from './office-font-families';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import { OfficeTableInsertPopover } from './office-table-insert-popover';
import { PresentationAnimationPanel } from './presentation-animation-panel';
import type {
  PresentationEditorCanCommands,
  PresentationEditorCommands,
} from './presentation-command-types';
import { PresentationTransitionPanel } from './presentation-transition-panel';
import {
  type WorkOfficeFileAction,
  WorkOfficeRibbon,
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';
import {
  officeMessage,
  resolveOfficeMessages,
  type OfficeMessageCatalog,
} from '../../../i18n/office-locale';
import { useOfficeMessages } from './office-messages-context';

function presentationRibbonTabs(catalog: OfficeMessageCatalog) {
  return [
    { id: 'home' as const, label: officeMessage(catalog, 'presentation.ribbon.tab.home') },
    { id: 'insert' as const, label: officeMessage(catalog, 'presentation.ribbon.tab.insert') },
    { id: 'design' as const, label: officeMessage(catalog, 'presentation.ribbon.tab.design') },
    {
      id: 'transitions' as const,
      label: officeMessage(catalog, 'presentation.ribbon.tab.transitions'),
    },
    {
      id: 'animations' as const,
      label: officeMessage(catalog, 'presentation.ribbon.tab.animations'),
    },
    {
      id: 'slideshow' as const,
      label: officeMessage(catalog, 'presentation.ribbon.tab.slideshow'),
      compactLabel: officeMessage(
        catalog,
        'presentation.ribbon.tab.slideshowCompact',
      ),
    },
    { id: 'review' as const, label: officeMessage(catalog, 'presentation.ribbon.tab.review') },
    { id: 'view' as const, label: officeMessage(catalog, 'presentation.ribbon.tab.view') },
  ];
}

const presentationFontMessages = resolveOfficeMessages();
const basePresentationFontFamilyOptions = officeFontFamilies.map((family) => {
  const label = officeFontFamilyLocalizedLabel(
    family,
    presentationFontMessages,
  );
  return {
    value: family.cssValue,
    group: officeFontFamilyGroupLabel(family.group, presentationFontMessages),
    label,
    previewStyle: { fontFamily: family.cssFamily },
    searchText: `${family.name} ${label}`,
  };
});

function presentationAlignmentActions(catalog: OfficeMessageCatalog) {
  return [
    {
      value: 'left' as const,
      label: officeMessage(catalog, 'presentation.align.left'),
      Icon: AlignLeft,
    },
    {
      value: 'center' as const,
      label: officeMessage(catalog, 'presentation.align.centerH'),
      Icon: AlignCenter,
    },
    {
      value: 'right' as const,
      label: officeMessage(catalog, 'presentation.align.right'),
      Icon: AlignRight,
    },
    {
      value: 'top' as const,
      label: officeMessage(catalog, 'presentation.align.top'),
      Icon: AlignVerticalJustifyStart,
    },
    {
      value: 'middle' as const,
      label: officeMessage(catalog, 'presentation.align.middle'),
      Icon: AlignVerticalJustifyCenter,
    },
    {
      value: 'bottom' as const,
      label: officeMessage(catalog, 'presentation.align.bottom'),
      Icon: AlignVerticalJustifyEnd,
    },
  ] as const satisfies readonly {
    value: OfficeKernelPresentationAlignment;
    label: string;
    Icon: typeof AlignLeft;
  }[];
}

export function PresentationToolbar({
  selectedSlide,
  selectedElement,
  selectedUnitCount,
  can,
  textFormattingAvailable,
  commentsOpen,
  commentCount,
  designOpen,
  editingDesign,
  background,
  transition,
  fileActions,
  notesVisible = true,
  viewMode = 'normal',
  onToggleNotes,
  commands,
}: {
  selectedSlide: WorkSlide;
  selectedElement: WorkSlideElement | null;
  selectedUnitCount: number;
  can: PresentationEditorCanCommands;
  textFormattingAvailable: boolean;
  commentsOpen: boolean;
  commentCount: number;
  designOpen: boolean;
  editingDesign: boolean;
  background?: string;
  transition: WorkSlide['transition'];
  fileActions?: readonly WorkOfficeFileAction[];
  notesVisible?: boolean;
  viewMode?: 'normal' | 'sorter';
  onToggleNotes?: () => void;
  commands: PresentationEditorCommands;
}) {
  const officeDialog = useOfficeDialog();
  const messages = useOfficeMessages();
  const ribbonTabs = presentationRibbonTabs(messages);
  const alignmentActions = presentationAlignmentActions(messages);
  const selectedFontSize = selectedElement
    ? String(selectedElement.fontSize)
    : '';
  const selectedFontSizeRef = useRef({
    elementId: selectedElement?.id ?? null,
    value: selectedFontSize,
  });
  const [fontSizeDraft, setFontSizeDraft] = useState(selectedFontSize);
  const fontFamilyValue = presentationFontFamilyValue(
    selectedElement?.fontFamily,
  );
  const selectedEntranceAnimation = workSlideAnimationForElement(
    selectedSlide,
    selectedElement?.id,
    'entrance',
  );
  const selectedExitAnimation = workSlideAnimationForElement(
    selectedSlide,
    selectedElement?.id,
    'exit',
  );
  useEffect(() => {
    const previous = selectedFontSizeRef.current;
    const elementId = selectedElement?.id ?? null;
    selectedFontSizeRef.current = { elementId, value: selectedFontSize };
    setFontSizeDraft((draft) =>
      previous.elementId !== elementId || draft === previous.value
        ? selectedFontSize
        : draft,
    );
  }, [selectedElement?.id, selectedFontSize]);
  const commitFontSize = (value: string): void => {
    if (!selectedElement) return;
    const fontSize = normalizedPresentationFontSize(
      value,
      selectedElement.fontSize,
    );
    setFontSizeDraft(String(fontSize));
    if (fontSize === selectedElement.fontSize) return;
    commands.updateElement({ fontSize }, { restoreTextFocus: false });
  };
  return (
    <>
      <WorkOfficeRibbon
        ariaLabel={officeMessage(messages, 'presentation.toolbar.aria')}
        tabs={ribbonTabs}
        defaultTab="home"
        fileActions={fileActions}
        collapsible
        className="work-presentation-ribbon"
        toolbarClassName="presentation-toolbar"
        panels={{
          home: (
            <>
              <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.group.undo')}>
                <WorkOfficeRibbonButton
                  label={officeMessage(messages, 'presentation.action.undo')}
                  title={officeMessage(messages, 'presentation.action.undoTitle')}
                  aria-keyshortcuts="Control+Z Meta+Z"
                  disabled={!can.undo()}
                  onClick={commands.undo}
                >
                  <Undo2 size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label={officeMessage(messages, 'presentation.action.redo')}
                  title={officeMessage(messages, 'presentation.action.redoTitle')}
                  aria-keyshortcuts="Control+Shift+Z Meta+Shift+Z Control+Y Meta+Y"
                  disabled={!can.redo()}
                  onClick={commands.redo}
                >
                  <Redo2 size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.group.slides')}>
                <WorkOfficeRibbonButton
                  label={officeMessage(messages, 'presentation.action.newSlide')}
                  title={officeMessage(messages, 'presentation.action.newSlideTitle')}
                  aria-keyshortcuts="Control+M Meta+Shift+N"
                  disabled={!can.addSlide()}
                  onClick={commands.addSlide}
                >
                  <Plus size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label={officeMessage(messages, 'presentation.action.duplicateSlide')}
                  title={officeMessage(messages, 'presentation.action.duplicateSlideTitle')}
                  aria-keyshortcuts="Control+D Meta+D"
                  disabled={!can.duplicateSlide()}
                  onClick={commands.duplicateSlide}
                >
                  <Copy size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label={officeMessage(messages, 'presentation.action.deleteSlide')}
                  title={officeMessage(messages, 'presentation.action.deleteSlideTitle')}
                  aria-keyshortcuts="Delete Backspace"
                  disabled={!can.deleteSlide()}
                  onClick={commands.deleteSlide}
                >
                  <Trash2 size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.group.clipboard')}>
                <WorkOfficeRibbonButton
                  label={officeMessage(messages, 'presentation.action.copy')}
                  title={officeMessage(messages, 'presentation.action.copyTitle')}
                  aria-keyshortcuts="Control+C Meta+C"
                  disabled={!can.copySelection()}
                  onClick={commands.copySelection}
                >
                  <Copy size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label={officeMessage(messages, 'presentation.action.cut')}
                  title={officeMessage(messages, 'presentation.action.cutTitle')}
                  aria-keyshortcuts="Control+X Meta+X"
                  disabled={!can.cutSelection()}
                  onClick={commands.cutSelection}
                >
                  <Scissors size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label={officeMessage(messages, 'presentation.action.paste')}
                  title={officeMessage(messages, 'presentation.action.pasteTitle')}
                  aria-keyshortcuts="Control+V Meta+V"
                  disabled={!can.pasteSelection()}
                  onClick={commands.pasteSelection}
                >
                  <ClipboardPaste size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
              {selectedElement && (
                <>
                  {textFormattingAvailable && (
                    <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.group.font')}>
                      <OfficeSelect
                        ariaLabel={officeMessage(messages, 'presentation.font.familyAria')}
                        className="presentation-font-family-select"
                        value={fontFamilyValue}
                        options={presentationFontFamilyOptions(fontFamilyValue)}
                        onValueChange={(fontFamily) =>
                          commands.updateElement(
                            { fontFamily },
                            { restoreTextFocus: false },
                          )
                        }
                      />
                      <OfficeNumberField
                        ariaLabel={officeMessage(messages, 'presentation.font.sizeAria')}
                        className="presentation-font-size-field"
                        min={8}
                        max={96}
                        step={1}
                        value={fontSizeDraft}
                        escapeConsumer={
                          fontSizeDraft !== String(selectedElement.fontSize)
                        }
                        onValueChange={setFontSizeDraft}
                        onCommit={commitFontSize}
                        onCancel={() =>
                          setFontSizeDraft(String(selectedElement.fontSize))
                        }
                      />
                      <WorkOfficeRibbonButton
                        label={officeMessage(messages, 'presentation.action.bold')}
                        title={officeMessage(messages, 'presentation.action.boldTitle')}
                        aria-keyshortcuts="Control+B Meta+B"
                        displayLabel={false}
                        active={Boolean(selectedElement.bold)}
                        disabled={!can.toggleBold()}
                        onClick={commands.toggleBold}
                      >
                        <Bold size={15} />
                      </WorkOfficeRibbonButton>
                      <WorkOfficeRibbonButton
                        label={officeMessage(messages, 'presentation.action.italic')}
                        title={officeMessage(messages, 'presentation.action.italicTitle')}
                        aria-keyshortcuts="Control+I Meta+I"
                        displayLabel={false}
                        active={Boolean(selectedElement.italic)}
                        disabled={!can.toggleItalic()}
                        onClick={commands.toggleItalic}
                      >
                        <Italic size={15} />
                      </WorkOfficeRibbonButton>
                      <WorkOfficeRibbonButton
                        label={officeMessage(messages, 'presentation.action.underline')}
                        title={officeMessage(messages, 'presentation.action.underlineTitle')}
                        aria-keyshortcuts="Control+U Meta+U"
                        displayLabel={false}
                        active={Boolean(selectedElement.underline)}
                        disabled={!can.toggleUnderline()}
                        onClick={commands.toggleUnderline}
                      >
                        <Underline size={15} />
                      </WorkOfficeRibbonButton>
                      {(
                        ['left', 'center', 'right'] as WorkSlideTextAlign[]
                      ).map((align) => (
                        <WorkOfficeRibbonButton
                          label={
                            align === 'left'
                              ? officeMessage(
                                  messages,
                                  'presentation.align.textLeft',
                                )
                              : align === 'center'
                                ? officeMessage(
                                    messages,
                                    'presentation.align.textCenter',
                                  )
                                : officeMessage(
                                    messages,
                                    'presentation.align.textRight',
                                  )
                          }
                          displayLabel={false}
                          active={selectedElement.align === align}
                          key={align}
                          onClick={() => commands.updateElement({ align })}
                        >
                          {align === 'left' ? (
                            <AlignLeft size={15} />
                          ) : align === 'center' ? (
                            <AlignCenter size={15} />
                          ) : (
                            <AlignRight size={15} />
                          )}
                        </WorkOfficeRibbonButton>
                      ))}
                      <OfficeColorPicker
                        compact
                        className="work-color-tool"
                        value={selectedElement.color}
                        ariaLabel={officeMessage(messages, 'presentation.font.colorAria')}
                        onValueChange={(color) =>
                          commands.updateElement(
                            { color },
                            { restoreTextFocus: false },
                          )
                        }
                      />
                    </WorkOfficeRibbonGroup>
                  )}
                  <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.group.arrange')}>
                    <PresentationAlignMenu
                      selectedUnitCount={selectedUnitCount}
                      can={can}
                      commands={commands}
                    />
                    {selectedUnitCount >= 3 && (
                      <>
                        <WorkOfficeRibbonButton
                          label={officeMessage(messages, 'presentation.action.distributeH')}
                          displayLabel={false}
                          disabled={!can.distributeElements('horizontal')}
                          onClick={() =>
                            commands.distributeElements('horizontal')
                          }
                        >
                          <AlignHorizontalSpaceBetween size={17} />
                        </WorkOfficeRibbonButton>
                        <WorkOfficeRibbonButton
                          label={officeMessage(messages, 'presentation.action.distributeV')}
                          displayLabel={false}
                          disabled={!can.distributeElements('vertical')}
                          onClick={() =>
                            commands.distributeElements('vertical')
                          }
                        >
                          <AlignVerticalSpaceBetween size={17} />
                        </WorkOfficeRibbonButton>
                      </>
                    )}
                    <WorkOfficeRibbonButton
                      label={officeMessage(messages, 'presentation.action.group')}
                      title={officeMessage(messages, 'presentation.action.groupTitle')}
                      aria-keyshortcuts="Control+G Meta+G"
                      displayLabel={false}
                      disabled={!can.groupElements()}
                      onClick={commands.groupElements}
                    >
                      <Group size={19} />
                    </WorkOfficeRibbonButton>
                    <WorkOfficeRibbonButton
                      label={officeMessage(messages, 'presentation.action.ungroup')}
                      title={officeMessage(messages, 'presentation.action.ungroupTitle')}
                      aria-keyshortcuts="Control+Shift+G Meta+Shift+G"
                      displayLabel={false}
                      disabled={!can.ungroupElements()}
                      onClick={commands.ungroupElements}
                    >
                      <Ungroup size={19} />
                    </WorkOfficeRibbonButton>
                    <WorkOfficeRibbonButton
                      label={officeMessage(messages, 'presentation.action.sendBackward')}
                      displayLabel={false}
                      disabled={!can.reorderElement(-1)}
                      onClick={() => commands.reorderElement(-1)}
                    >
                      <ArrowDownToLine size={19} />
                    </WorkOfficeRibbonButton>
                    <WorkOfficeRibbonButton
                      label={officeMessage(messages, 'presentation.action.bringForward')}
                      displayLabel={false}
                      disabled={!can.reorderElement(1)}
                      onClick={() => commands.reorderElement(1)}
                    >
                      <ArrowUpToLine size={19} />
                    </WorkOfficeRibbonButton>
                  </WorkOfficeRibbonGroup>
                </>
              )}
            </>
          ),
          insert: (
            <>
              <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.group.textShapes')}>
                <WorkOfficeRibbonButton
                  label={officeMessage(messages, 'presentation.action.textBox')}
                  disabled={!can.addElement('text')}
                  onClick={() => commands.addElement('text')}
                >
                  <Type size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label={officeMessage(messages, 'presentation.action.shape')}
                  disabled={!can.addElement('shape')}
                  onClick={() => commands.addElement('shape')}
                >
                  <Square size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.group.content')}>
                <WorkOfficeRibbonButton
                  label={officeMessage(messages, 'presentation.action.picture')}
                  disabled={!can.requestImage()}
                  onClick={commands.requestImage}
                >
                  <Image size={19} />
                </WorkOfficeRibbonButton>
                {!editingDesign && (
                  <>
                    <OfficeTableInsertPopover
                      label={officeMessage(messages, 'presentation.action.table')}
                      disabled={!can.addTable({ rows: 1, columns: 1 })}
                      onInsert={commands.addTable}
                    />
                    <WorkOfficeRibbonButton
                      label={officeMessage(messages, 'presentation.action.chart')}
                      disabled={!can.addChart()}
                      onClick={commands.addChart}
                    >
                      <BarChart3 size={19} />
                    </WorkOfficeRibbonButton>
                  </>
                )}
              </WorkOfficeRibbonGroup>
              {selectedElement &&
                (selectedElement.type === 'text' ||
                  selectedElement.type === 'shape') && (
                  <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.group.link')}>
                    <WorkOfficeRibbonButton
                      label={officeMessage(messages, 'presentation.action.link')}
                      active={Boolean(selectedElement.href)}
                      onClick={() =>
                        void officeDialog
                          .prompt({
                            title: officeMessage(messages, 'presentation.link.title'),
                            description: officeMessage(
                              messages,
                              'presentation.link.description',
                            ),
                            fieldLabel: officeMessage(
                              messages,
                              'presentation.link.field',
                            ),
                            initialValue: selectedElement.href ?? 'https://',
                            placeholder: 'https://',
                            inputMode: 'url',
                            confirmLabel: officeMessage(
                              messages,
                              'presentation.link.confirm',
                            ),
                            validate: (value) =>
                              value.trim() && !normalizeDocumentHref(value)
                                ? DOCUMENT_LINK_VALIDATION_MESSAGE
                                : null,
                          })
                          .then((href) => {
                            if (href !== null) {
                              const normalized = href.trim()
                                ? normalizeDocumentHref(href)
                                : undefined;
                              commands.updateElement({
                                href: normalized ?? undefined,
                              });
                            }
                          })
                      }
                    >
                      <Link2 size={19} />
                    </WorkOfficeRibbonButton>
                  </WorkOfficeRibbonGroup>
                )}
            </>
          ),
          design: (
            <>
              <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.group.master')}>
                <WorkOfficeRibbonButton
                  label={officeMessage(messages, 'presentation.action.masterLayouts')}
                  active={designOpen}
                  onClick={commands.toggleDesign}
                >
                  <LayoutTemplate size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.group.background')}>
                <OfficeColorPicker
                  compact
                  className="work-color-tool slide-background-tool"
                  value={background ?? selectedSlide.background}
                  ariaLabel={editingDesign ? officeMessage(messages, 'presentation.background.designAria') : officeMessage(messages, 'presentation.background.slideAria')}
                  onValueChange={commands.setBackground}
                />
              </WorkOfficeRibbonGroup>
            </>
          ),
          transitions: (
            <PresentationTransitionPanel
              slideId={selectedSlide.id}
              transition={transition}
              editable={can.setTransition(transition)}
              canApplyToAll={can.applyTransitionToAll}
              onChange={commands.setTransition}
              onApplyToAll={commands.applyTransitionToAll}
            />
          ),
          animations: (
            <PresentationAnimationPanel
              animations={{
                entrance: selectedEntranceAnimation,
                exit: selectedExitAnimation,
              }}
              canMove={can.moveAnimation}
              canPreview={can.previewAnimations()}
              canUpdate={can.updateAnimation}
              editable={can.setAnimation(
                'entrance',
                selectedEntranceAnimation?.effect,
              )}
              elementId={selectedElement?.id}
              onMove={commands.moveAnimation}
              onPreview={commands.previewAnimations}
              onSetEffect={commands.setAnimation}
              onUpdate={commands.updateAnimation}
            />
          ),
          slideshow: (
            <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.group.startShow')}>
              <WorkOfficeRibbonButton
                label={officeMessage(messages, 'presentation.action.showFromStart')}
                title={officeMessage(messages, 'presentation.action.showFromStartTitle')}
                aria-keyshortcuts="F5"
                data-presentation-slideshow-source="beginning"
                disabled={!can.startSlideshow('beginning')}
                onClick={() => commands.startSlideshow('beginning')}
              >
                <Play size={19} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label={officeMessage(messages, 'presentation.action.showFromCurrent')}
                title={officeMessage(messages, 'presentation.action.showFromCurrentTitle')}
                aria-keyshortcuts="Shift+F5"
                data-presentation-slideshow-source="current"
                disabled={!can.startSlideshow('current')}
                onClick={() => commands.startSlideshow('current')}
              >
                <SquarePlay size={19} />
              </WorkOfficeRibbonButton>
            </WorkOfficeRibbonGroup>
          ),
          review: (
            <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.group.comments')}>
              <WorkOfficeRibbonButton
                label={officeMessage(messages, 'presentation.action.newComment')}
                disabled={editingDesign || !can.addComment()}
                onClick={commands.addComment}
              >
                <MessageSquarePlus size={19} />
              </WorkOfficeRibbonButton>
              <WorkOfficeRibbonButton
                label={
                  commentCount
                    ? officeMessage(
                        messages,
                        'presentation.action.viewCommentsWithCount',
                        { count: String(commentCount) },
                      )
                    : officeMessage(messages, 'presentation.action.viewComments')
                }
                disabled={editingDesign}
                active={commentsOpen}
                onClick={commands.toggleComments}
              >
                <MessagesSquare size={19} />
              </WorkOfficeRibbonButton>
            </WorkOfficeRibbonGroup>
          ),
          view: (
            <>
              <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.group.presentationViews')}>
                <WorkOfficeRibbonButton
                  label={officeMessage(messages, 'presentation.action.normalView')}
                  active={viewMode === 'normal'}
                  onClick={() => commands.setViewMode('normal')}
                >
                  <PanelsTopLeft size={19} />
                </WorkOfficeRibbonButton>
                <WorkOfficeRibbonButton
                  label={officeMessage(messages, 'presentation.action.slideSorter')}
                  active={viewMode === 'sorter'}
                  onClick={() => commands.setViewMode('sorter')}
                >
                  <Grid2X2 size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.group.show')}>
                <WorkOfficeRibbonButton
                  label={officeMessage(messages, 'presentation.action.notes')}
                  active={notesVisible}
                  disabled={viewMode !== 'normal' || !onToggleNotes}
                  title={
                    viewMode === 'normal'
                      ? notesVisible
                        ? officeMessage(messages, 'presentation.action.hideNotes')
                        : officeMessage(messages, 'presentation.action.showNotes')
                      : officeMessage(messages, 'presentation.action.notesNormalOnly')
                  }
                  onClick={() => onToggleNotes?.()}
                >
                  <NotebookPen size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.group.master')}>
                <WorkOfficeRibbonButton
                  label={officeMessage(messages, 'presentation.action.masterView')}
                  active={designOpen}
                  onClick={commands.toggleDesign}
                >
                  <LayoutTemplate size={19} />
                </WorkOfficeRibbonButton>
              </WorkOfficeRibbonGroup>
            </>
          ),
        }}
      />
      {officeDialog.dialog}
    </>
  );
}

function PresentationAlignMenu({
  selectedUnitCount,
  can,
  commands,
}: {
  selectedUnitCount: number;
  can: PresentationEditorCanCommands;
  commands: PresentationEditorCommands;
}) {
  const messages = useOfficeMessages();
  const alignmentActions = presentationAlignmentActions(messages);
  const modeHint =
    selectedUnitCount > 1
      ? officeMessage(messages, 'presentation.align.selectedHint')
      : officeMessage(messages, 'presentation.align.slideHint');
  const disabled = !can.alignElement('left');

  return (
    <Popover
      label={officeMessage(messages, 'presentation.align.objects')}
      panelLabel={modeHint}
      panelRole="menu"
      portal
      className="presentation-align-menu"
      panelClassName="work-office-context-menu presentation-align-menu-panel"
      disabled={disabled}
      focusFirstOnOpen
      onPanelKeyDown={moveOfficeMenuFocus}
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          type="button"
          className={`presentation-align-trigger${open ? ' open' : ''}`}
          title={modeHint}
          aria-description={modeHint}
        >
          <span>{officeMessage(messages, 'presentation.align.objects')}</span>
          <ChevronDown size={14} aria-hidden="true" />
        </button>
      )}
    >
      {(close) =>
        alignmentActions.map(
          ({ value, label: itemLabel, Icon }) => (
            <button
              key={value}
              type="button"
              role="menuitem"
              tabIndex={-1}
              disabled={!can.alignElement(value)}
              onClick={() => {
                close();
                commands.alignElement(value);
              }}
            >
              <Icon size={15} aria-hidden="true" />
              <span>{itemLabel}</span>
            </button>
          ),
        )
      }
    </Popover>
  );
}

function normalizedPresentationFontSize(
  value: string,
  current: number,
): number {
  if (!value.trim()) return current;
  const number = Number(value);
  if (!Number.isFinite(number)) return current;
  return Math.min(96, Math.max(8, Math.round(number)));
}

function presentationFontFamilyValue(current: string | undefined): string {
  const value = current?.trim() || 'Aptos';
  const normalized = normalizePresentationFontFamily(value);
  return (
    basePresentationFontFamilyOptions.find(
      (option) => normalizePresentationFontFamily(option.value) === normalized,
    )?.value ?? value
  );
}

function presentationFontFamilyOptions(current: string) {
  if (
    basePresentationFontFamilyOptions.some((option) => option.value === current)
  ) {
    return basePresentationFontFamilyOptions;
  }
  return [
    ...basePresentationFontFamilyOptions,
    {
      value: current,
      group: officeMessage(resolveOfficeMessages(), 'presentation.font.documentGroup'),
      label: officeFontFamilyLabel(current),
      previewStyle: { fontFamily: current },
      searchText: current,
    },
  ];
}

function normalizePresentationFontFamily(value: string): string {
  return normalizeOfficeFontFamily(value);
}
