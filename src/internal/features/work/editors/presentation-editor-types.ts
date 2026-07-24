import type { WorkEditorAgentRequest } from '../work-agent-request';
import type { WorkPresentationContent } from '../work-types';
import type { WorkOfficeFileAction } from './work-office-chrome';

export interface PresentationEditorProps {
  content: WorkPresentationContent;
  preview: boolean;
  saveStatus?: string;
  fileActions?: readonly WorkOfficeFileAction[];
  kernelWasmUrl?: string;
  onChange: (content: WorkPresentationContent) => void;
  onAgentRequest?: (request: WorkEditorAgentRequest) => void | Promise<void>;
  onStartSlideshow?: () => void;
}

export interface PresentationDragState {
  elementId: string;
  mode: 'move' | 'resize';
  pointerId: number;
  startX: number;
  startY: number;
  originX: number;
  originY: number;
  originWidth: number;
  originHeight: number;
}

export interface PresentationAgentMenuState {
  x: number;
  y: number;
  selection: string;
  target: 'slide' | 'element';
  slideId: string;
  elementId: string | null;
}
