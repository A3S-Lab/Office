import type { WorkOfficeCollaborationSession } from '../../../collaboration/office-collaboration';
import type { WorkEditorAgentRequest } from '../work-agent-request';
import type { WorkPresentationContent } from '../work-types';
import type { WorkOfficeFileAction } from './work-office-chrome';

export type PresentationDesignMode = 'slide' | 'layout' | 'master';

export interface PresentationEditorProps {
  autoFocus?: boolean;
  collaboration?: WorkOfficeCollaborationSession;
  content: WorkPresentationContent;
  preview: boolean;
  saveStatus?: string;
  fileActions?: readonly WorkOfficeFileAction[];
  kernelWasmUrl?: string;
  onChange: (content: WorkPresentationContent) => void;
  onAgentRequest?: (request: WorkEditorAgentRequest) => void | Promise<void>;
  onStartSlideshow?: () => void;
}

export interface PresentationAgentMenuState {
  x: number;
  y: number;
  selection: string;
  target: 'slide' | 'element';
  slideId: string;
  elementId: string | null;
}
