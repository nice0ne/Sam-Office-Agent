import { describe, it, expect } from 'vitest';
import {
  HostType,
  isValidHost,
  ProviderConfig,
  ChatMessage,
  ToolDefinition,
  ToolCall,
  ActionQueueItem,
  AgentAction,
  ExecutionMode,
} from '../src/types';
import { AgentContext, AgentResponse, IAgent, ExecutionContext } from '../src/agents/types';

describe('HostType & Guards', () => {
  it('validates known Office hosts correctly', () => {
    const excelHost: HostType = 'Excel';
    expect(isValidHost(excelHost)).toBe(true);
    expect(isValidHost('Word')).toBe(true);
    expect(isValidHost('PowerPoint')).toBe(true);
    expect(isValidHost('BrowserDev')).toBe(true);
    expect(isValidHost('UnknownHost')).toBe(false);
  });
});

describe('Domain Models & Contracts', () => {
  it('instantiates valid ProviderConfig objects', () => {
    const config: ProviderConfig = {
      id: 'gemini',
      name: 'Google Gemini',
      apiKey: 'test-api-key',
      selectedModel: 'gemini-2.0-flash',
      enabled: true,
    };
    expect(config.id).toBe('gemini');
    expect(config.enabled).toBe(true);
  });

  it('instantiates valid ChatMessage with tool calls', () => {
    const toolCall: ToolCall = {
      id: 'call-1',
      name: 'writeRange',
      arguments: { range: 'A1', values: [['Hello']] },
      status: 'pending',
    };

    const msg: ChatMessage = {
      id: 'msg-1',
      role: 'assistant',
      content: 'Writing data to range',
      timestamp: Date.now(),
      toolCalls: [toolCall],
      status: 'done',
    };

    expect(msg.role).toBe('assistant');
    expect(msg.toolCalls?.length).toBe(1);
    expect(msg.toolCalls?.[0].status).toBe('pending');
  });

  it('supports ToolDefinition schema contract', () => {
    const tool: ToolDefinition = {
      name: 'formatRange',
      description: 'Format specified cells in active worksheet',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string', description: 'Target range like A1:B10' },
          bold: { type: 'boolean', description: 'Bold text formatting' },
        },
        required: ['range'],
      },
    };
    expect(tool.name).toBe('formatRange');
    expect(tool.parameters.required).toContain('range');
  });

  it('supports ActionQueueItem and AgentAction alias', () => {
    const item: ActionQueueItem = {
      id: 'action-1',
      stepNumber: 1,
      title: 'Set Header',
      description: 'Format row 1 with bold and fill color',
      toolCall: {
        id: 'tc-1',
        name: 'formatHeader',
        arguments: { row: 1 },
        status: 'pending',
      },
      status: 'pending',
    };

    const agentAction: AgentAction = item;
    expect(agentAction.stepNumber).toBe(1);
    expect(agentAction.status).toBe('pending');

    const mode: ExecutionMode = 'copilot';
    expect(mode).toBe('copilot');
  });

  it('supports AgentContext and IAgent implementation', async () => {
    const context: AgentContext = {
      host: 'Excel',
      activeCellOrRange: 'B2',
      documentSummary: 'Sales report 2026',
    };

    const execContext: ExecutionContext = context;
    expect(execContext.host).toBe('Excel');

    class MockAgent implements IAgent {
      id = 'mock-excel-agent';
      name = 'Mock Excel Agent';
      hostType: HostType = 'Excel';

      getSystemPrompt(ctx: AgentContext): string {
        return `Specialist for ${ctx.host}`;
      }

      getTools(): ToolDefinition[] {
        return [];
      }

      async executeTool(toolCall: ToolCall, _ctx: AgentContext) {
        return { success: true, result: toolCall.name };
      }
    }

    const agent = new MockAgent();
    expect(agent.getSystemPrompt(context)).toContain('Excel');
    const result = await agent.executeTool({ id: '1', name: 'ping', arguments: {}, status: 'pending' }, context);
    expect(result.success).toBe(true);
    expect(result.result).toBe('ping');

    const resp: AgentResponse = {
      message: 'Done',
      toolCalls: [],
    };
    expect(resp.message).toBe('Done');
  });
});

