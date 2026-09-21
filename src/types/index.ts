export type HostType = 'Excel' | 'Word' | 'PowerPoint' | 'BrowserDev';

export function isValidHost(host: string): host is HostType {
  return ['Excel', 'Word', 'PowerPoint', 'BrowserDev'].includes(host);
}

export type ProviderId = 'gemini' | 'openai' | 'claude' | 'glm' | 'openrouter' | 'ollama' | 'openai-compatible';

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
  role: 'user' | 'assistant' | 'system' | 'tool';
  content: string;
  displayContent?: string;
  timestamp: number;
  toolCalls?: ToolCall[];
  toolCallId?: string;
  toolName?: string;
  isError?: boolean;
  status?: 'sending' | 'streaming' | 'done' | 'error';
  error?: string;
  stepProgress?: {
    step: number;
    maxSteps: number;
    description?: string;
  };
}

export interface ToolProperty {
  type: 'string' | 'number' | 'boolean' | 'array' | 'object';
  description: string;
  enum?: string[];
  items?: { type: string; properties?: Record<string, any>; required?: string[] } | Record<string, any>;
  properties?: Record<string, any>;
  required?: string[];
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
