import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import JSZip from 'jszip';
import { createRef } from 'react';
import {
  presentationElementDisplayBox,
  rotatePresentationSelection,
} from '../src/internal/features/work/editors/presentation-selection';
import {
  PresentationWorkspace,
  type PresentationWorkspaceCommands,
} from '../src/internal/features/work/editors/presentation-workspace';
import {
  PptxGroupExportRegistry,
  patchPptxNativeGroups,
} from '../src/internal/features/work/work-pptx-groups';
import { ungroupPresentationElements } from '../src/internal/features/work/work-presentation-groups';
import { withPresentationDesign } from '../src/internal/features/work/work-presentation-layouts';
import type {
  WorkPresentationContent,
  WorkSlideElement,
} from '../src/internal/features/work/work-types';

test('rotates one group together and keeps child positions in group space', () => {
  const shape = element('shape-1', 10, 10);
  const picture = element('picture-1', 40, 12);
  const once = rotatePresentationSelection([shape, picture], ['shape-1'], 15);

  expect(once.map((item) => item.groupRotation)).toEqual([15, 15]);
  expect(once.map((item) => item.x)).toEqual([10, 40]);
  expect(once.map((item) => item.rotation)).toEqual([undefined, undefined]);

  const wrapped = rotatePresentationSelection(once, ['picture-1'], 350);
  expect(wrapped.map((item) => item.groupRotation)).toEqual([5, 5]);

  const ungrouped = ungroupPresentationElements(once, ['shape-1']);
  expect(ungrouped.map((item) => item.groupIds)).toEqual([
    undefined,
    undefined,
  ]);
  expect(ungrouped.map((item) => item.rotation)).toEqual([15, 15]);
  expect(ungrouped.map((item) => item.groupRotation)).toEqual([
    undefined,
    undefined,
  ]);
});

test('turns a rotated group around the group center on the slide', () => {
  const shape = element('shape-1', 0, 0);
  const picture = element('picture-1', 20, 0);
  const rotated = rotatePresentationSelection(
    [shape, picture],
    ['shape-1'],
    90,
  );
  const displayed = presentationElementDisplayBox(rotated[0], rotated);

  expect(displayed.rotation).toBe(90);
  expect(displayed.x).toBeCloseTo(10);
  expect(displayed.y).toBeCloseTo(-10);
});

test('writes rot on the exported group for Shape 1 and Picture 1', async () => {
  const registry = new PptxGroupExportRegistry();
  const shape = { ...element('shape-1', 10, 10), groupRotation: 15 };
  const picture = {
    ...element('picture-1', 40, 12),
    type: 'image' as const,
    groupRotation: 15,
  };
  const shapeMarker = registry.objectName('slide-1', shape, 'shape');
  const pictureMarker = registry.objectName('slide-1', picture, 'image');
  const zip = new JSZip();
  zip.file(
    'ppt/slides/slide1.xml',
    `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>${sceneShape(shapeMarker ?? '', 0, 0, 100, 80)}${sceneShape(pictureMarker ?? '', 120, 0, 90, 70)}</p:spTree></p:cSld></p:sld>`,
  );

  const patched = await patchPptxNativeGroups(
    await zip.generateAsync({ type: 'arraybuffer' }),
    registry,
  );
  const xml = await JSZip.loadAsync(patched).then((archive) =>
    archive.file('ppt/slides/slide1.xml')?.async('string'),
  );

  expect(xml).toContain('name="Group 1"');
  expect(xml).toContain('name="Shape 1"');
  expect(xml).toContain('name="Picture 1"');
  expect(xml).toContain('rot="900000"');
});

test('dragging the selection rotate handle turns the group', () => {
  const rotated: number[] = [];
  const shape = element('shape-1', 10, 10);
  const picture = element('picture-1', 40, 12);
  const content: WorkPresentationContent = {
    type: 'presentation',
    slides: [
      {
        id: 'slide-1',
        name: 'Slide 1',
        background: '#ffffff',
        elements: [shape, picture],
      },
    ],
  };
  render(
    <PresentationWorkspace
      activeBackground="#ffffff"
      activeCommentId={null}
      activeElements={[shape, picture]}
      aspectRatio="16 / 9"
      canvasName="Slide canvas"
      canvasRef={createRef<HTMLElement>()}
      commands={workspaceCommands({
        rotateSelection: (degrees) => {
          rotated.push(degrees);
          return true;
        },
      })}
      content={content}
      designContent={withPresentationDesign(content)}
      designMode="slide"
      editingElementId={null}
      inheritedElements={[]}
      placeholderGuides={[]}
      selectedElementIds={['shape-1', 'picture-1']}
      selectedLayout={undefined}
      selectedMaster={undefined}
      selectedSlide={content.slides[0]}
      snapGuides={[]}
      notesVisible
      viewMode="normal"
      zoom={100}
      onBeginDrag={() => undefined}
      onContinueDrag={() => undefined}
      onDragCancel={() => undefined}
      onDragEnd={() => undefined}
      onOpenContextMenu={() => undefined}
      onTextEditorChange={() => undefined}
      onTextSelectionChange={() => undefined}
    />,
  );

  const frame = document.querySelector<HTMLElement>(
    '[data-presentation-selection-frame]',
  );
  if (!frame) throw new Error('Missing selection frame.');
  frame.getBoundingClientRect = () =>
    ({
      x: 0,
      y: 0,
      top: 0,
      left: 0,
      right: 100,
      bottom: 100,
      width: 100,
      height: 100,
      toJSON() {
        return {};
      },
    }) as DOMRect;
  const handle = screen.getByRole('button', { name: '旋转手柄' });
  fireEvent.pointerDown(handle, {
    button: 0,
    clientX: 50,
    clientY: 0,
    pointerId: 1,
  });
  fireEvent.pointerUp(handle, {
    button: 0,
    clientX: 100,
    clientY: 50,
    pointerId: 1,
  });

  expect(rotated).toEqual([90]);
});

function element(id: string, x: number, y: number): WorkSlideElement {
  return {
    id,
    type: 'shape',
    groupIds: ['group-1'],
    x,
    y,
    width: 20,
    height: 10,
    text: id,
    fontSize: 14,
    color: '#172033',
    fill: '#dce6fb',
    bold: false,
    align: 'center',
  };
}

function sceneShape(
  name: string,
  x: number,
  y: number,
  width: number,
  height: number,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="2" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></a:xfrm></p:spPr></p:sp>`;
}

function workspaceCommands(
  overrides: Partial<PresentationWorkspaceCommands>,
): PresentationWorkspaceCommands {
  return {
    addSlide: () => undefined,
    deleteSlideById: () => false,
    editElement: () => undefined,
    exitEditing: () => undefined,
    instantiatePlaceholder: () => undefined,
    moveSlide: () => false,
    openComment: () => undefined,
    rotateSelection: () => false,
    selectElement: () => undefined,
    selectSlide: () => undefined,
    setViewMode: () => undefined,
    updateElement: () => undefined,
    updateNotes: () => undefined,
    updateTextElement: () => undefined,
    ...overrides,
  };
}
