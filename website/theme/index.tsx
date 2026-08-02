import './index.css';

import { getCustomMDXComponent as getOriginalMDXComponent } from '@rspress/core/theme-original';
import { PlaygroundLink } from './playground-link';

export { Nav } from './nav';
export * from '@rspress/core/theme-original';

export function getCustomMDXComponent() {
  return {
    ...getOriginalMDXComponent(),
    PlaygroundLink,
  };
}
