export type HomeLanguage = 'zh' | 'en';
export type ChapterKind =
  | 'document'
  | 'markdown'
  | 'spreadsheet'
  | 'presentation'
  | 'pdf';

export const chapterOrder: readonly ChapterKind[] = [
  'document',
  'markdown',
  'spreadsheet',
  'presentation',
  'pdf',
];

export function MotionArrow() {
  return (
    <svg
      aria-hidden="true"
      className="office-editor-chapters__arrow"
      viewBox="0 0 16 16"
    >
      <path d="M3 8h10" />
      <path d="m9 4 4 4-4 4" />
    </svg>
  );
}
