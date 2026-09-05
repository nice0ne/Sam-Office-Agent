import { ChatMessage, ProviderConfig, ToolCall, ToolDefinition } from '../../types';

export interface ChatRequest {
  messages: ChatMessage[];
  systemPrompt?: string;
  tools?: ToolDefinition[];
  temperature?: number;
}

export interface StreamEvent {
  type: 'content_delta' | 'tool_call' | 'done' | 'error';
  delta?: string;
  toolCall?: ToolCall;
  error?: string;
}

export interface ILLMProvider {
  id: string;
  name: string;
  sendMessage(req: ChatRequest, config: ProviderConfig): AsyncIterable<StreamEvent>;
  testConnection(config: ProviderConfig): Promise<{ success: boolean; message: string }>;
}
