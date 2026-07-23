import type { OfficeArtifactKind } from '@a3s-lab/office/core';
import {
  FileCode2,
  FileSpreadsheet,
  FileText,
  FileType2,
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
  if (kind === 'pdf') return <FileType2 size={size} />;
  return <FileText size={size} />;
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
