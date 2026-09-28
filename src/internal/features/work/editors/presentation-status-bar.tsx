import { Cloud, Grid2X2, PanelsTopLeft } from 'lucide-react';
import { officeMessage } from '../../../i18n/office-locale';
import type { WorkPresentationContent, WorkSlide } from '../work-types';
import { useOfficeMessages } from './office-messages-context';
import { presentationCommentCount } from './presentation-comments-panel';
import {
  WorkOfficeStatusBar,
  WorkOfficeZoomControls,
} from './work-office-chrome';

export function PresentationStatusBar({
  content,
  selectedSlide,
  viewMode,
  zoom,
  saveStatus,
  onViewModeChange,
  onZoomChange,
}: {
  content: WorkPresentationContent;
  selectedSlide: WorkSlide;
  viewMode: 'normal' | 'sorter';
  zoom: number;
  saveStatus: string;
  onViewModeChange: (mode: 'normal' | 'sorter') => void;
  onZoomChange: (zoom: number) => void;
}) {
  const messages = useOfficeMessages();
  const slideNumber =
    content.slides.findIndex((slide) => slide.id === selectedSlide.id) + 1;

  return (
    <WorkOfficeStatusBar
      className="work-presentation-status"
      controls={
        <>
          <button
            type="button"
            aria-label={officeMessage(
              messages,
              'presentation.status.normalViewAria',
            )}
            title={officeMessage(
              messages,
              'presentation.status.normalViewTitle',
            )}
            aria-pressed={viewMode === 'normal'}
            onClick={() => onViewModeChange('normal')}
          >
            <PanelsTopLeft size={13} />
          </button>
          <button
            type="button"
            aria-label={officeMessage(
              messages,
              'presentation.status.sorterViewAria',
            )}
            title={officeMessage(
              messages,
              'presentation.status.sorterViewTitle',
            )}
            aria-pressed={viewMode === 'sorter'}
            onClick={() => onViewModeChange('sorter')}
          >
            <Grid2X2 size={13} />
          </button>
          <span className="work-office-status-divider" />
          <WorkOfficeZoomControls
            zoom={zoom}
            decreaseLabel={officeMessage(
              messages,
              'presentation.status.zoomOut',
            )}
            increaseLabel={officeMessage(
              messages,
              'presentation.status.zoomIn',
            )}
            outputLabel={officeMessage(
              messages,
              'presentation.status.zoomOutput',
            )}
            sliderLabel={officeMessage(
              messages,
              'presentation.status.zoomSlider',
            )}
            onChange={onZoomChange}
          />
        </>
      }
    >
      <output
        aria-label={officeMessage(messages, 'presentation.status.slideAria')}
        className="work-presentation-status-primary"
      >
        {officeMessage(messages, 'presentation.status.slideProgress', {
          current: String(slideNumber),
          total: String(content.slides.length),
        })}
      </output>
      <output
        aria-label={officeMessage(messages, 'presentation.status.notesAria')}
        className="work-presentation-status-secondary"
      >
        {selectedSlide.notes?.trim()
          ? officeMessage(messages, 'presentation.status.notesPresent')
          : officeMessage(messages, 'presentation.status.notesAbsent')}
      </output>
      <output
        aria-label={officeMessage(
          messages,
          'presentation.status.commentsAria',
        )}
        className="work-presentation-status-secondary"
      >
        {officeMessage(messages, 'presentation.status.commentsCount', {
          count: String(presentationCommentCount(content.slides)),
        })}
      </output>
      <output
        aria-label={officeMessage(messages, 'presentation.status.saveAria')}
        className="work-office-save-status"
      >
        <Cloud size={12} />
        {saveStatus}
      </output>
    </WorkOfficeStatusBar>
  );
}
