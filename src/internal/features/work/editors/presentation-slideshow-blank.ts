export type PresentationBlankScreen = 'off' | 'black' | 'white';

/** WPS/PowerPoint slideshow blank-screen shortcuts (B/. black, W/, white). */
export function nextPresentationBlankScreen(
  current: PresentationBlankScreen,
  key: string,
): PresentationBlankScreen | null {
  const normalized = key.length === 1 ? key.toLowerCase() : key;
  if (normalized === 'b' || key === '.') {
    return current === 'black' ? 'off' : 'black';
  }
  if (normalized === 'w' || key === ',') {
    return current === 'white' ? 'off' : 'white';
  }
  return null;
}

export function presentationBlankScreenLabel(
  mode: Exclude<PresentationBlankScreen, 'off'>,
): string {
  return mode === 'black' ? '黑屏' : '白屏';
}
