import { CheckCheck, Cloud, FileText, Globe2 } from 'lucide-react';
import { officeMessage } from '../../../i18n/office-locale';
import { getDocumentCommandDefinition } from './document-command-catalog';
import {
  documentCommandLabel,
  documentCommandTitleWithShortcut,
} from './document-command-i18n';
import type { DocumentViewMode } from './document-toolbar';
import {
  clampDocumentZoom,
  MAX_DOCUMENT_ZOOM,
  MIN_DOCUMENT_ZOOM,
} from './document-zoom';
import { useOfficeMessages } from './office-messages-context';
import {
  WorkOfficeStatusBar,
  WorkOfficeZoomControls,
} from './work-office-chrome';

interface DocumentStatusBarProps {
  bibliographyCount: number;
  citationCount: number;
  currentPage: number;
  pageCount: number;
  saveStatus: string;
  sectionCount: number;
  sectionIndex: number;
  spellcheckEnabled: boolean;
  viewMode: DocumentViewMode;
  wordCount: number;
  zoom: number;
  onSpellcheckChange: (enabled: boolean) => void;
  onOpenWordCount: () => void;
  onViewModeChange: (mode: DocumentViewMode) => void;
  onZoomChange: (zoom: number) => void;
}

export function DocumentStatusBar({
  bibliographyCount,
  citationCount,
  currentPage,
  pageCount,
  saveStatus,
  sectionCount,
  sectionIndex,
  spellcheckEnabled,
  viewMode,
  wordCount,
  zoom,
  onSpellcheckChange,
  onOpenWordCount,
  onViewModeChange,
  onZoomChange,
}: DocumentStatusBarProps) {
  const messages = useOfficeMessages();
  const wordCountCommand = getDocumentCommandDefinition('wordCount');
  const proofingLabel = officeMessage(
    messages,
    spellcheckEnabled
      ? 'document.status.proofing.on'
      : 'document.status.proofing.off',
  );

  return (
    <WorkOfficeStatusBar
      ariaLabel={officeMessage(messages, 'document.status.aria')}
      controlsLabel={officeMessage(messages, 'document.status.controlsAria')}
      className="work-document-footer"
      controls={
        <>
          <button
            type="button"
            data-document-status-control="view-mode"
            aria-label={officeMessage(messages, 'document.status.pageView')}
            title={officeMessage(messages, 'document.status.pageView')}
            aria-pressed={viewMode === 'page'}
            onClick={() => onViewModeChange('page')}
          >
            <FileText size={13} />
          </button>
          <button
            type="button"
            data-document-status-control="view-mode"
            aria-label={officeMessage(messages, 'document.status.webView')}
            title={officeMessage(messages, 'document.status.webView')}
            aria-pressed={viewMode === 'web'}
            onClick={() => onViewModeChange('web')}
          >
            <Globe2 size={13} />
          </button>
          <span
            className="work-office-status-divider"
            data-document-status-control="view-mode"
          />
          <WorkOfficeZoomControls
            zoom={zoom}
            minimum={MIN_DOCUMENT_ZOOM}
            maximum={MAX_DOCUMENT_ZOOM}
            step={5}
            decreaseLabel={officeMessage(messages, 'document.status.zoomOut')}
            increaseLabel={officeMessage(messages, 'document.status.zoomIn')}
            outputLabel={officeMessage(messages, 'document.status.zoomOutput')}
            sliderLabel={officeMessage(messages, 'document.status.zoomSlider')}
            onChange={(value) => onZoomChange(clampDocumentZoom(value))}
          />
        </>
      }
    >
      <output
        aria-label={officeMessage(messages, 'document.status.pageAria')}
        data-document-status-item="page"
      >
        {officeMessage(messages, 'document.status.page', {
          current: String(currentPage),
          total: String(pageCount),
        })}
      </output>
      <output
        aria-label={officeMessage(messages, 'document.status.sectionAria')}
        data-document-status-item="section"
      >
        {officeMessage(messages, 'document.status.section', {
          current: String(sectionIndex + 1),
          total: String(sectionCount),
        })}
      </output>
      <button
        type="button"
        className="work-office-status-text-button"
        data-document-status-item="word-count"
        aria-label={officeMessage(messages, 'document.status.wordCountAria', {
          count: String(wordCount),
        })}
        aria-keyshortcuts={wordCountCommand.shortcut?.aria}
        title={documentCommandTitleWithShortcut('wordCount', messages)}
        onClick={onOpenWordCount}
      >
        {officeMessage(messages, 'document.status.words', {
          count: String(wordCount),
        })}
      </button>
      <button
        type="button"
        data-document-status-item="spellcheck"
        aria-label={proofingLabel}
        title={proofingLabel}
        aria-pressed={spellcheckEnabled}
        onClick={() => onSpellcheckChange(!spellcheckEnabled)}
      >
        <CheckCheck size={12} />
      </button>
      <output aria-label={officeMessage(messages, 'document.status.citationAria')}>
        {officeMessage(messages, 'document.status.citation', {
          bibliography: String(bibliographyCount),
          citations: String(citationCount),
        })}
      </output>
      <output
        aria-label={officeMessage(messages, 'document.status.saveAria')}
        className="work-office-save-status"
      >
        <Cloud size={12} />
        {saveStatus}
      </output>
    </WorkOfficeStatusBar>
  );
}
