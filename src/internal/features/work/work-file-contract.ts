import { WORK_SUPPORTED_FILE_EXTENSIONS } from './work-file-kind';

export const WORK_IMPORT_ACCEPT = WORK_SUPPORTED_FILE_EXTENSIONS.map(
  (extension) => `.${extension}`,
).join(',');
