const SYNTHETIC_PICTURE_LABEL = /^(?:undefined|preencoded\.[a-z0-9]+)$/i;

/**
 * Author alt text, then a real picture name. pptxgenjs otherwise writes the
 * data-URL stand-in `preencoded.png` into `cNvPr/@descr`.
 */
export function presentationPictureAltText(
  ...candidates: Array<string | undefined>
): string | undefined {
  for (const candidate of candidates) {
    const text = candidate?.trim();
    if (text && !SYNTHETIC_PICTURE_LABEL.test(text)) return text;
  }
  return undefined;
}
