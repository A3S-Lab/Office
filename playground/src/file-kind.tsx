import type { OfficeArtifactKind } from '@a3s-lab/office/core';
import {
  FileCode2,
  FileSpreadsheet,
  FileText,
  Presentation,
} from 'lucide-react';

export function FileKindIcon({
  kind,
  size = 18,
}: {
  kind: OfficeArtifactKind;
  size?: number;
}) {
  if (kind === 'spreadsheet') return <FileSpreadsheet size={size} />;
  if (kind === 'presentation') return <Presentation size={size} />;
  if (kind === 'markdown') return <FileCode2 size={size} />;
  if (kind === 'pdf') return <PdfFileIcon size={size} />;
  return <FileText size={size} />;
}

function PdfFileIcon({ size }: { size: number }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M14 2H6.8A1.8 1.8 0 0 0 5 3.8v16.4A1.8 1.8 0 0 0 6.8 22h10.4a1.8 1.8 0 0 0 1.8-1.8V7Z" />
      <path d="M14 2v5h5" />
      <rect
        x="7"
        y="12.5"
        width="10"
        height="6"
        rx="1.2"
        fill="currentColor"
        stroke="none"
      />
      <text
        x="12"
        y="16.95"
        fill="white"
        stroke="none"
        textAnchor="middle"
        fontFamily="Arial, sans-serif"
        fontSize="4.6"
        fontWeight="800"
      >
        PDF
      </text>
    </svg>
  );
}

export function fileKindLabel(kind: OfficeArtifactKind): string {
  if (kind === 'spreadsheet') return '表格';
  if (kind === 'presentation') return '演示';
  if (kind === 'markdown') return 'Markdown';
  if (kind === 'pdf') return 'PDF';
  return '文字';
}

export function fileKindExtension(kind: OfficeArtifactKind): string {
  if (kind === 'spreadsheet') return 'XLSX';
  if (kind === 'presentation') return 'PPTX';
  if (kind === 'markdown') return 'MD';
  if (kind === 'pdf') return 'PDF';
  return 'DOCX';
}
