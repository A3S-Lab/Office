export function fileNameWithoutExtension(value: string): string {
  return value.replace(/\.[^.]+$/, '').trim() || '导入的文件';
}

export function safeFileName(value: string): string {
  return (
    value.replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_').trim() || 'A3S Work 文件'
  );
}

export function downloadBlob(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = name;
  anchor.click();
  // Headless browsers can cancel the download if the object URL is revoked
  // before Chrome finishes reading it. Keep the URL alive through that read.
  window.setTimeout(() => URL.revokeObjectURL(url), 15_000);
}
