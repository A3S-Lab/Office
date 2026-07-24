import { useEffect, useState } from 'react';
import {
  documentLayoutFontKey,
  type WorkDocumentLayoutFont,
} from '../work-document-fonts';

interface BrowserFontRegistration {
  face: FontFace;
  loaded: Promise<boolean>;
  references: number;
}

const browserFontRegistrations = new Map<string, BrowserFontRegistration>();

export function useDocumentLayoutFonts(
  fonts: readonly WorkDocumentLayoutFont[],
): ReadonlySet<string> {
  const key = documentLayoutFontKey(fonts);
  const [loadedFontIds, setLoadedFontIds] = useState<ReadonlySet<string>>(
    () => new Set(),
  );

  useEffect(() => {
    if (
      typeof document === 'undefined' ||
      typeof FontFace === 'undefined' ||
      !document.fonts
    ) {
      setLoadedFontIds(new Set());
      return;
    }
    let disposed = false;
    const acquired = fonts.map(acquireBrowserFont);
    void Promise.all(
      acquired.map(async ({ font, registration }) => ({
        id: font.id,
        loaded: await registration.loaded,
      })),
    ).then((results) => {
      if (disposed) return;
      setLoadedFontIds(
        new Set(
          results.flatMap((result) => (result.loaded ? [result.id] : [])),
        ),
      );
    });

    return () => {
      disposed = true;
      for (const { registryKey } of acquired) releaseBrowserFont(registryKey);
    };
  }, [key]);

  return loadedFontIds;
}

function acquireBrowserFont(font: WorkDocumentLayoutFont): {
  font: WorkDocumentLayoutFont;
  registration: BrowserFontRegistration;
  registryKey: string;
} {
  const registryKey = documentLayoutFontKey([font]);
  let registration = browserFontRegistrations.get(registryKey);
  if (!registration) {
    const face = new FontFace(font.family, `url(${JSON.stringify(font.url)})`, {
      style: font.style ?? 'normal',
      weight: String(font.weight ?? 400),
    });
    registration = {
      face,
      loaded: face
        .load()
        .then((loadedFace) => {
          document.fonts.add(loadedFace);
          return true;
        })
        .catch(() => false),
      references: 0,
    };
    browserFontRegistrations.set(registryKey, registration);
  }
  registration.references += 1;
  return { font, registration, registryKey };
}

function releaseBrowserFont(registryKey: string): void {
  const registration = browserFontRegistrations.get(registryKey);
  if (!registration) return;
  registration.references -= 1;
  if (registration.references > 0) return;
  browserFontRegistrations.delete(registryKey);
  void registration.loaded.then((loaded) => {
    if (loaded && registration.references <= 0) {
      document.fonts.delete(registration.face);
    }
  });
}
