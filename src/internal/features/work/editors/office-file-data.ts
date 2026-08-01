import { materializeWorkFile } from '../work-file-data';

export async function readOfficeFileAsDataUrl(file: File): Promise<string> {
  const ownedFile = await materializeWorkFile(file);
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener('load', () => resolve(String(reader.result)));
    reader.addEventListener('error', () =>
      reject(reader.error ?? new Error('File could not be read')),
    );
    reader.readAsDataURL(ownedFile);
  });
}
