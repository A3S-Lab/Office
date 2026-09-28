import { officeMessage } from '../../../i18n/office-locale';
import { useOfficeMessages } from './office-messages-context';
import {
  WorkOfficeStatusBar,
  WorkOfficeZoomControls,
} from './work-office-chrome';

const MARKDOWN_MIN_ZOOM = 60;
const MARKDOWN_MAX_ZOOM = 180;

export function MarkdownStatus({
  characterCount,
  lineCount,
  saveStatus,
  zoom,
  onZoomChange,
}: {
  characterCount: number;
  lineCount: number;
  saveStatus?: string;
  zoom: number;
  onZoomChange: (zoom: number) => void;
}) {
  const messages = useOfficeMessages();
  return (
    <WorkOfficeStatusBar
      className="work-markdown-status"
      controls={
        <WorkOfficeZoomControls
          zoom={zoom}
          minimum={MARKDOWN_MIN_ZOOM}
          maximum={MARKDOWN_MAX_ZOOM}
          decreaseLabel={officeMessage(messages, 'markdown.status.zoomOut')}
          increaseLabel={officeMessage(messages, 'markdown.status.zoomIn')}
          outputLabel={officeMessage(messages, 'markdown.status.zoomOutput')}
          sliderLabel={officeMessage(messages, 'markdown.status.zoomSlider')}
          onChange={onZoomChange}
        />
      }
    >
      <output>{officeMessage(messages, 'markdown.status.lines', { count: String(lineCount) })}</output>
      <output>{officeMessage(messages, 'markdown.status.chars', { count: String(characterCount) })}</output>
      {saveStatus && (
        <span className="work-office-save-status">{saveStatus}</span>
      )}
    </WorkOfficeStatusBar>
  );
}
