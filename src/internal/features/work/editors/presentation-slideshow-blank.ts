import {
  officeMessage,
  resolveOfficeMessages,
} from '../../../i18n/office-locale';
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
  const catalog = resolveOfficeMessages();
  return mode === 'black'
    ? officeMessage(catalog, 'presentation.blank.black')
    : officeMessage(catalog, 'presentation.blank.white');
}
