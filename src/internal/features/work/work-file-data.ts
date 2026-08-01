export async function materializeWorkFile(file: File): Promise<File> {
  const bytes = await file.arrayBuffer();
  return new File([bytes], file.name, {
    lastModified: file.lastModified,
    type: file.type,
  });
}
