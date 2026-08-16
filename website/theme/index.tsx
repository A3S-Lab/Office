import '@fontsource-variable/geist';
import '@fontsource-variable/geist-mono';
import './docs-content.css';
import './index.css';

import { getCustomMDXComponent as getOriginalMDXComponent } from '@rspress/core/theme-original';
import { PlaygroundLink } from './playground-link';

export { Nav } from './components/Nav';
export * from '@rspress/core/theme-original';

export function getCustomMDXComponent() {
  return {
    ...getOriginalMDXComponent(),
    PlaygroundLink,
  };
}
