import { ChatMessage, HostType, ToolCall, ToolDefinition } from '../types';

export interface AgentContext {
  host: HostType;
  selectedText?: string;
  activeCellOrRange?: string;
  documentSummary?: string;
}

export type ExecutionContext = AgentContext;

export interface AgentResponse {
  message: string;
  toolCalls: ToolCall[];
}

export interface IAgent {
  id: string;
  name: string;
  hostType: HostType;
  getSystemPrompt(context: AgentContext): string;
  getTools(): ToolDefinition[];
  executeTool(toolCall: ToolCall, context: AgentContext): Promise<{ success: boolean; result?: any; error?: string }>;
}

export type { ChatMessage };
