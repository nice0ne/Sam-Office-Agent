import { ChatMessage, ProviderConfig, ToolCall } from '../../types';
import { ILLMProvider } from '../../services/llm/types';
import { AgentContext, IAgent } from '../types';

export interface ReActCallbacks {
  onStepProgress?: (step: number, maxSteps: number, description: string) => void;
  onContentDelta?: (delta: string) => void;
  onToolExecuted?: (toolCall: ToolCall, success: boolean, resultOrError?: any) => void;
}

export interface ReActParams {
  provider: ILLMProvider;
  providerConfig: ProviderConfig;
  specialist: IAgent;
  initialMessages: ChatMessage[];
  systemPrompt: string;
  context?: AgentContext;
  maxIterations?: number;
  abortSignal?: AbortSignal;
}

export interface ReActRunResult {
  completed: boolean;
  aborted?: boolean;
  iterations: number;
  finalContent: string;
  allMessages: ChatMessage[];
}

export class ReActExecutionEngine {
  async runLoop(params: ReActParams, callbacks?: ReActCallbacks): Promise<ReActRunResult> {
    const maxIterations = params.maxIterations ?? 6;
    let iterations = 0;
    const conversation: ChatMessage[] = [...params.initialMessages];
    let finalContent = '';
    const tools = params.specialist.getTools();

    while (iterations < maxIterations) {
      if (params.abortSignal?.aborted) {
        return {
          completed: false,
          aborted: true,
          iterations,
          finalContent,
          allMessages: conversation,
        };
      }

      iterations++;
      callbacks?.onStepProgress?.(iterations, maxIterations, `Iterasi [${iterations}/${maxIterations}]...`);

      const assistantMsgId = `asst_step_${Date.now()}_${iterations}`;
      const assistantMsg: ChatMessage = {
        id: assistantMsgId,
        role: 'assistant',
        content: '',
        timestamp: Date.now(),
        toolCalls: [],
      };

      const stream = params.provider.sendMessage(
        {
          messages: conversation,
          systemPrompt: params.systemPrompt,
          tools,
        },
        params.providerConfig
      );

      for await (const chunk of stream) {
        if (params.abortSignal?.aborted) {
          break;
        }

        if (chunk.type === 'content_delta' && chunk.delta) {
          assistantMsg.content += chunk.delta;
          callbacks?.onContentDelta?.(chunk.delta);
        } else if (chunk.type === 'tool_call' && chunk.toolCall) {
          assistantMsg.toolCalls?.push(chunk.toolCall);
        } else if (chunk.type === 'error') {
          assistantMsg.content += `\n[Error: ${chunk.error}]`;
        }

        if (params.abortSignal?.aborted) {
          break;
        }
      }

      conversation.push(assistantMsg);
      finalContent = assistantMsg.content;

      if (params.abortSignal?.aborted) {
        return {
          completed: false,
          aborted: true,
          iterations,
          finalContent,
          allMessages: conversation,
        };
      }

      // If no tools were called, the agent has finished its turn
      if (!assistantMsg.toolCalls || assistantMsg.toolCalls.length === 0) {
        return {
          completed: true,
          iterations,
          finalContent,
          allMessages: conversation,
        };
      }

      // Execute each tool call
      for (const toolCall of assistantMsg.toolCalls) {
        if (params.abortSignal?.aborted) {
          return {
            completed: false,
            aborted: true,
            iterations,
            finalContent,
            allMessages: conversation,
          };
        }

        callbacks?.onStepProgress?.(iterations, maxIterations, `Mengeksekusi tool ${toolCall.name}...`);
        const context = params.context || { host: params.specialist.hostType };
        const execResult = await params.specialist.executeTool(toolCall, context);
        toolCall.status = execResult.success ? 'applied' : 'failed';
        toolCall.error = execResult.error;
        callbacks?.onToolExecuted?.(toolCall, execResult.success, execResult.result ?? execResult.error);

        // Append tool result into conversation history for the next iteration (Self-Correction & Observation)
        const toolMsg: ChatMessage = {
          id: `tool_${Date.now()}_${toolCall.id}`,
          role: 'tool',
          toolCallId: toolCall.id,
          toolName: toolCall.name,
          content: execResult.success
            ? (typeof execResult.result === 'string' ? execResult.result : JSON.stringify(execResult.result ?? { success: true }))
            : `Error: ${execResult.error || 'Eksekusi tool gagal'}`,
          isError: !execResult.success,
          timestamp: Date.now(),
        };
        conversation.push(toolMsg);
      }
    }

    return {
      completed: false,
      iterations,
      finalContent,
      allMessages: conversation,
    };
  }
}
