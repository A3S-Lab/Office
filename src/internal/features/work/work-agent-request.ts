import type { WorkAgentProposalRequest } from './work-agent-proposal';

export interface WorkEditorAgentRequest {
  instruction: string;
  selection?: string;
  proposal?: WorkAgentProposalRequest;
}
