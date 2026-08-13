import type { PresentationContent } from '../../src/core';

export function presentationCollaborationFixture(): PresentationContent {
  return {
    type: 'presentation',
    width: 13.333,
    height: 7.5,
    masters: [
      {
        id: 'master-1',
        name: 'Main master',
        background: '#FFFFFF',
        elements: [],
      },
    ],
    layouts: [
      {
        id: 'layout-1',
        name: 'Title',
        masterId: 'master-1',
        elements: [],
      },
    ],
    slides: [
      {
        id: 'slide-1',
        name: 'Cover',
        background: '#FFFFFF',
        layoutId: 'layout-1',
        elements: [
          {
            id: 'element-title',
            type: 'text',
            x: 10,
            y: 10,
            width: 80,
            height: 20,
            text: 'Shared presentation',
            fontSize: 32,
            color: '#172033',
            fill: 'transparent',
            bold: true,
            align: 'center',
          },
        ],
      },
      {
        id: 'slide-2',
        name: 'Details',
        background: '#FFFFFF',
        layoutId: 'layout-1',
        elements: [
          {
            id: 'element-body',
            type: 'shape',
            x: 20,
            y: 25,
            width: 60,
            height: 40,
            text: 'Body',
            fontSize: 18,
            color: '#172033',
            fill: '#DCE6FB',
            bold: false,
            align: 'left',
          },
        ],
      },
    ],
  };
}
