import { WorkFileImportController } from './work-file-import';

const WORK_FILE_READ_CHUNK_BYTES = 4 * 1024 * 1024;

export interface MaterializedWorkFile {
  bytes: ArrayBuffer;
  file: File;
}

export async function materializeWorkFile(
  file: File,
  controller = new WorkFileImportController({}, file.size),
): Promise<File> {
  return (await materializeWorkFileSource(file, controller)).file;
}

export async function materializeWorkFileSource(
  file: File,
  controller: WorkFileImportController,
): Promise<MaterializedWorkFile> {
  const bytes = await readWorkFileBytes(file, controller);
  return {
    bytes,
    file: new File([bytes], file.name, {
      lastModified: file.lastModified,
      type: file.type,
    }),
  };
}

async function readWorkFileBytes(
  file: File,
  controller: WorkFileImportController,
): Promise<ArrayBuffer> {
  controller.report('reading', 0, 0);
  if (file.size <= WORK_FILE_READ_CHUNK_BYTES) {
    const bytes = await abortableArrayBuffer(file.arrayBuffer(), controller);
    controller.report('reading', 1, bytes.byteLength);
    return bytes;
  }

  const bytes = new Uint8Array(file.size);
  for (
    let offset = 0;
    offset < file.size;
    offset += WORK_FILE_READ_CHUNK_BYTES
  ) {
    controller.throwIfAborted();
    const end = Math.min(file.size, offset + WORK_FILE_READ_CHUNK_BYTES);
    const chunk = await abortableArrayBuffer(
      file.slice(offset, end).arrayBuffer(),
      controller,
    );
    bytes.set(new Uint8Array(chunk), offset);
    controller.report('reading', end / file.size, end);
    if (end < file.size) await controller.yieldToMainThread();
  }
  return bytes.buffer;
}

async function abortableArrayBuffer(
  pending: Promise<ArrayBuffer>,
  controller: WorkFileImportController,
): Promise<ArrayBuffer> {
  controller.throwIfAborted();
  const signal = controller.signal;
  if (!signal) return pending;
  return new Promise<ArrayBuffer>((resolve, reject) => {
    const abort = () => {
      cleanup();
      try {
        controller.throwIfAborted();
      } catch (error) {
        reject(error);
      }
    };
    const cleanup = () => signal.removeEventListener('abort', abort);
    signal.addEventListener('abort', abort, { once: true });
    pending.then(
      (bytes) => {
        cleanup();
        resolve(bytes);
      },
      (error) => {
        cleanup();
        reject(error);
      },
    );
  });
}
