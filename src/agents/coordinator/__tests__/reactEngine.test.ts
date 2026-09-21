import { describe, it, expect } from 'vitest';
import { ReActExecutionEngine, ReActCallbacks } from '../reactEngine';
import { ILLMProvider } from '../../../services/llm/types';
import { IAgent } from '../../types';

describe('ReActExecutionEngine', () => {
  it('executes single turn when no tool calls are returned', async () => {
    const mockProvider: ILLMProvider = {
      id: 'mock',
      name: 'Mock',
      testConnection: async () => ({ success: true, message: 'OK' }),
      sendMessage: async function* () {
        yield { type: 'content_delta', delta: 'Halo ada yang bisa dibantu?' };
      },
    };

    const mockSpecialist: IAgent = {
      id: 'test',
      name: 'Test',
      hostType: 'Excel',
      getSystemPrompt: () => 'Prompt',
      getTools: () => [],
      executeTool: async () => ({ success: true }),
    };

    const engine = new ReActExecutionEngine();
    const result = await engine.runLoop({
      provider: mockProvider,
      providerConfig: { id: 'openai', name: 'OpenAI', apiKey: 'k', selectedModel: 'gpt-4o', enabled: true },
      specialist: mockSpecialist,
      initialMessages: [{ id: '1', role: 'user', content: 'Halo', timestamp: 1 }],
      systemPrompt: 'Sys',
      maxIterations: 4,
    });

    expect(result.iterations).toBe(1);
    expect(result.finalContent).toBe('Halo ada yang bisa dibantu?');
    expect(result.completed).toBe(true);
    expect(result.allMessages).toHaveLength(2); // user + assistant
    expect(result.allMessages[1].role).toBe('assistant');
  });

  it('runs multi-turn loop and performs self-correction on tool error', async () => {
    let callCount = 0;
    const capturedMessagesPerCall: any[][] = [];

    const mockProvider: ILLMProvider = {
      id: 'mock',
      name: 'Mock',
      testConnection: async () => ({ success: true, message: 'OK' }),
      sendMessage: async function* (req) {
        callCount++;
        capturedMessagesPerCall.push([...req.messages]);
        if (callCount === 1) {
          yield {
            type: 'tool_call',
            toolCall: { id: 'call_1', name: 'write_cells', arguments: { range: 'INVALID' }, status: 'pending' },
          };
        } else {
          yield { type: 'content_delta', delta: 'Berhasil diperbaiki setelah error.' };
        }
      },
    };

    const mockSpecialist: IAgent = {
      id: 'test',
      name: 'Test',
      hostType: 'Excel',
      getSystemPrompt: () => 'Prompt',
      getTools: () => [],
      executeTool: async (tc) => {
        if (tc.arguments.range === 'INVALID') {
          return { success: false, error: 'Range tidak valid' };
        }
        return { success: true, result: { updated: 5 } };
      },
    };

    const engine = new ReActExecutionEngine();
    const result = await engine.runLoop({
      provider: mockProvider,
      providerConfig: { id: 'openai', name: 'OpenAI', apiKey: 'k', selectedModel: 'gpt-4o', enabled: true },
      specialist: mockSpecialist,
      initialMessages: [{ id: '1', role: 'user', content: 'Tulis sel', timestamp: 1 }],
      systemPrompt: 'Sys',
      maxIterations: 3,
    });

    expect(callCount).toBe(2);
    expect(result.iterations).toBe(2);
    expect(result.completed).toBe(true);
    expect(result.finalContent).toContain('Berhasil diperbaiki setelah error.');

    // Verify tool error was appended to conversation and passed to call 2
    expect(capturedMessagesPerCall[1]).toHaveLength(3); // user, assistant (with tool_call), tool (with error)
    const toolMsgInCall2 = capturedMessagesPerCall[1][2];
    expect(toolMsgInCall2.role).toBe('tool');
    expect(toolMsgInCall2.toolCallId).toBe('call_1');
    expect(toolMsgInCall2.toolName).toBe('write_cells');
    expect(toolMsgInCall2.isError).toBe(true);
    expect(toolMsgInCall2.content).toContain('Range tidak valid');
  });

  it('respects safety limit maxIterations (customizable and default 6)', async () => {
    // Test custom maxIterations: 3
    let callCount = 0;
    const mockLoopingProvider: ILLMProvider = {
      id: 'mock',
      name: 'Mock',
      testConnection: async () => ({ success: true, message: 'OK' }),
      sendMessage: async function* () {
        callCount++;
        yield {
          type: 'tool_call',
          toolCall: { id: `call_${callCount}`, name: 'read_cell', arguments: {}, status: 'pending' },
        };
      },
    };

    const mockSpecialist: IAgent = {
      id: 'test',
      name: 'Test',
      hostType: 'Excel',
      getSystemPrompt: () => 'Prompt',
      getTools: () => [],
      executeTool: async () => ({ success: true, result: 'data' }),
    };

    const engine = new ReActExecutionEngine();
    const customResult = await engine.runLoop({
      provider: mockLoopingProvider,
      providerConfig: { id: 'openai', name: 'OpenAI', apiKey: 'k', selectedModel: 'gpt-4o', enabled: true },
      specialist: mockSpecialist,
      initialMessages: [{ id: '1', role: 'user', content: 'Infinite loop test', timestamp: 1 }],
      systemPrompt: 'Sys',
      maxIterations: 3,
    });

    expect(customResult.iterations).toBe(3);
    expect(customResult.completed).toBe(false);

    // Test default maxIterations (should default to 6)
    callCount = 0;
    const defaultResult = await engine.runLoop({
      provider: mockLoopingProvider,
      providerConfig: { id: 'openai', name: 'OpenAI', apiKey: 'k', selectedModel: 'gpt-4o', enabled: true },
      specialist: mockSpecialist,
      initialMessages: [{ id: '1', role: 'user', content: 'Default limit test', timestamp: 1 }],
      systemPrompt: 'Sys',
    });

    expect(defaultResult.iterations).toBe(6);
    expect(defaultResult.completed).toBe(false);
  });

  it('triggers progress, content delta, and tool executed callbacks', async () => {
    const stepProgressCalls: Array<{ step: number; maxSteps: number; desc: string }> = [];
    const contentDeltas: string[] = [];
    const executedTools: Array<{ name: string; success: boolean }> = [];

    const callbacks: ReActCallbacks = {
      onStepProgress: (step, maxSteps, description) => {
        stepProgressCalls.push({ step, maxSteps, desc: description });
      },
      onContentDelta: (delta) => {
        contentDeltas.push(delta);
      },
      onToolExecuted: (toolCall, success) => {
        executedTools.push({ name: toolCall.name, success });
      },
    };

    let step = 0;
    const mockProvider: ILLMProvider = {
      id: 'mock',
      name: 'Mock',
      testConnection: async () => ({ success: true, message: 'OK' }),
      sendMessage: async function* () {
        step++;
        if (step === 1) {
          yield { type: 'content_delta', delta: 'Menganalisis ' };
          yield { type: 'content_delta', delta: 'data...' };
          yield {
            type: 'tool_call',
            toolCall: { id: 'call_1', name: 'format_table', arguments: {}, status: 'pending' },
          };
        } else {
          yield { type: 'content_delta', delta: 'Selesai diformat.' };
        }
      },
    };

    const mockSpecialist: IAgent = {
      id: 'test',
      name: 'Test',
      hostType: 'Excel',
      getSystemPrompt: () => 'Prompt',
      getTools: () => [],
      executeTool: async () => ({ success: true, result: { formatted: true } }),
    };

    const engine = new ReActExecutionEngine();
    const result = await engine.runLoop(
      {
        provider: mockProvider,
        providerConfig: { id: 'openai', name: 'OpenAI', apiKey: 'k', selectedModel: 'gpt-4o', enabled: true },
        specialist: mockSpecialist,
        initialMessages: [{ id: '1', role: 'user', content: 'Format ini', timestamp: 1 }],
        systemPrompt: 'Sys',
      },
      callbacks
    );

    expect(result.completed).toBe(true);
    expect(result.iterations).toBe(2);
    expect(contentDeltas).toEqual(['Menganalisis ', 'data...', 'Selesai diformat.']);
    expect(executedTools).toEqual([{ name: 'format_table', success: true }]);
    expect(stepProgressCalls.length).toBeGreaterThanOrEqual(2);
    expect(stepProgressCalls[0].step).toBe(1);
    expect(stepProgressCalls[0].maxSteps).toBe(6);
  });
});
