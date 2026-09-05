export type HostType = 'Excel' | 'Word' | 'PowerPoint' | 'BrowserDev';

export function isValidHost(host: string): host is HostType {
  return ['Excel', 'Word', 'PowerPoint', 'BrowserDev'].includes(host);
}

export type ProviderId = 'gemini' | 'openai' | 'claude' | 'glm' | 'openrouter' | 'ollama';

export interface ProviderConfig {
  id: ProviderId;
  name: string;
  apiKey: string;
  baseUrl?: string;
  selectedModel: string;
  enabled: boolean;
}

export interface ChatMessage {
  id: string;
  role: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: number;
  toolCalls?: ToolCall[];
  status?: 'sending' | 'streaming' | 'done' | 'error';
  error?: string;
}

export interface ToolProperty {
  type: 'string' | 'number' | 'boolean' | 'array' | 'object';
  description: string;
  enum?: string[];
  items?: { type: string };
}

export interface ToolDefinition {
  name: string;
  description: string;
  parameters: {
    type: 'object';
    properties: Record<string, ToolProperty>;
    required?: string[];
  };
}

export interface ToolCall {
  id: string;
  name: string;
  arguments: Record<string, any>;
  status: 'pending' | 'applied' | 'rejected' | 'failed';
  error?: string;
}

export interface ActionQueueItem {
  id: string;
  stepNumber: number;
  title: string;
  description: string;
  toolCall: ToolCall;
  status: 'pending' | 'executing' | 'completed' | 'failed';
}

export type AgentAction = ActionQueueItem;

export type ExecutionMode = 'copilot' | 'autopilot';
