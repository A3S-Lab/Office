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
  return (
    <WorkOfficeStatusBar
      className="work-markdown-status"
      controls={
        <WorkOfficeZoomControls
          zoom={zoom}
          minimum={MARKDOWN_MIN_ZOOM}
          maximum={MARKDOWN_MAX_ZOOM}
          decreaseLabel="缩小内容"
          increaseLabel="放大内容"
          outputLabel="Markdown 缩放比例"
          sliderLabel="调整 Markdown 缩放比例"
          onChange={onZoomChange}
        />
      }
    >
      <output>{lineCount} 行</output>
      <output>{characterCount} 字符</output>
      {saveStatus && (
        <span className="work-office-save-status">{saveStatus}</span>
      )}
    </WorkOfficeStatusBar>
  );
}
