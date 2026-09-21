# Autonomous ReAct Engine, Latency Optimization, and Office Suite Capabilities Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Transform Sam Office Agent into an autonomous ReAct multi-turn assistant with context caching, single-sync batching, expanded native tools for Excel, Word, and PowerPoint, and a Quick Action Presets toolbar.

**Architecture:** Extend `ChatMessage` and LLM provider adapters to natively support multi-turn `role: 'tool'` messages. Implement `ReActExecutionEngine` with safety limits (max 6 steps) and error self-correction. Introduce `DocumentContextCache` for sub-50ms conversational response times. Implement high-value native tools across Excel, Word, and PowerPoint, and surface them via an adaptive Quick Action Presets toolbar in the chat UI.

**Tech Stack:** TypeScript 5.6, React 18.3, Office.js, Vitest, Tailwind CSS.

**Spec:** [docs/superpowers/specs/2026-09-21-agent-react-optimization-capabilities-design.md](file:///E:/VIBE-CODE-WS/Sam-Office-Agent/docs/superpowers/specs/2026-09-21-agent-react-optimization-capabilities-design.md)

## Global Constraints

- 100% Client-Side BYOK architecture: Direct browser `fetch` to provider endpoints with zero intermediary proxy or telemetry.
- Strict Office.js safety: Protect against unsupported APIs (e.g. PowerPoint `notesPage` API limitations) and batch Office.js mutations with single `context.sync()`.
- Backward compatibility: Existing custom tools and settings must continue to work without regression.
- All existing 219 Vitest tests must pass at every step.

---

### Task 1: Update ChatMessage Schema & LLM Multi-Turn Tool Adapters

**Files:**
- Modify: `src/types/index.ts:18-27`
- Modify: `src/services/llm/openai.ts:24-50`
- Modify: `src/services/llm/gemini.ts:17-45`
- Modify: `src/services/llm/anthropic.ts:25-60`
- Test: `src/services/llm/__tests__/llmAdapters.test.ts`

**Interfaces:**
- Consumes: Existing `ChatMessage`, `ToolCall`, `ChatRequest`
- Produces: Extended `ChatMessage` with `role: 'tool'`, `toolCallId?: string`, `toolName?: string`, `isError?: boolean`. Adapters format tool history compliant with OpenAI, Gemini, and Anthropic APIs.

- [ ] **Step 1: Write the failing test in `src/services/llm/__tests__/llmAdapters.test.ts`**

Add tests checking that `OpenAIProvider`, `GeminiProvider`, and `AnthropicProvider` serialize `role: 'tool'` and assistant `toolCalls` correctly in the request payload:
```typescript
it('OpenAIProvider serializes assistant tool_calls and tool result messages', async () => {
  const provider = new OpenAIProvider();
  let capturedBody: any = null;
  globalThis.fetch = vi.fn().mockImplementation(async (_url, options) => {
    capturedBody = JSON.parse(options.body);
    return {
      ok: true,
      body: {
        getReader: () => ({
          read: vi.fn()
            .mockResolvedValueOnce({ done: false, value: new TextEncoder().encode('data: {"choices":[{"delta":{"content":"Done"}}]}\n\n') })
            .mockResolvedValueOnce({ done: true, value: undefined }),
        }),
      },
    };
  });

  const msgs: ChatMessage[] = [
    { id: '1', role: 'user', content: 'Hitung total', timestamp: 1 },
    { id: '2', role: 'assistant', content: '', timestamp: 2, toolCalls: [{ id: 'call_1', name: 'read_sheet', arguments: {}, status: 'applied' }] },
    { id: '3', role: 'tool', toolCallId: 'call_1', toolName: 'read_sheet', content: '{"values":[[100]]}', timestamp: 3 },
  ];

  const events = [];
  for await (const e of provider.sendMessage({ messages: msgs }, { id: 'openai', name: 'OpenAI', apiKey: 'test', selectedModel: 'gpt-4o', enabled: true })) {
    events.push(e);
  }

  expect(capturedBody.messages).toHaveLength(3);
  expect(capturedBody.messages[1].tool_calls).toBeDefined();
  expect(capturedBody.messages[1].tool_calls[0].id).toBe('call_1');
  expect(capturedBody.messages[2].role).toBe('tool');
  expect(capturedBody.messages[2].tool_call_id).toBe('call_1');
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- src/services/llm/__tests__/llmAdapters.test.ts -t "OpenAIProvider serializes assistant tool_calls"`
Expected: FAIL due to missing properties or unhandled tool serialization.

- [ ] **Step 3: Update `src/types/index.ts` and LLM provider adapters**

In `src/types/index.ts`:
```typescript
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
}
```

In `src/services/llm/openai.ts`:
Format messages:
```typescript
for (const m of req.messages) {
  if (m.role === 'assistant' && m.toolCalls && m.toolCalls.length > 0) {
    messages.push({
      role: 'assistant',
      content: m.content || null,
      tool_calls: m.toolCalls.map(tc => ({
        id: tc.id,
        type: 'function',
        function: {
          name: tc.name,
          arguments: typeof tc.arguments === 'string' ? tc.arguments : JSON.stringify(tc.arguments || {}),
        },
      })),
    });
  } else if (m.role === 'tool') {
    messages.push({
      role: 'tool',
      tool_call_id: m.toolCallId || 'call_default',
      content: m.content,
    });
  } else {
    messages.push({ role: m.role, content: m.content });
  }
}
```

In `src/services/llm/gemini.ts` and `src/services/llm/anthropic.ts`:
Implement equivalent multi-turn mapping for Gemini (`functionCall` and `functionResponse`) and Anthropic (`tool_use` and `tool_result`).

- [ ] **Step 4: Run tests to verify they pass**

Run: `npm test -- src/services/llm/__tests__/llmAdapters.test.ts`
Expected: PASS (all tests pass).

- [ ] **Step 5: Commit changes**

```bash
git add src/types/index.ts src/services/llm/ src/services/llm/__tests__/llmAdapters.test.ts
git commit -m "feat(llm): add multi-turn tool calling protocol to ChatMessage and LLM adapters"
```

---

### Task 2: Document Context Snapshot Cache & Invalidation

**Files:**
- Create: `src/services/office/contextCache.ts`
- Test: `src/services/office/__tests__/contextCache.test.ts`

**Interfaces:**
- Consumes: `HostType`, `AgentContext`
- Produces: `DocumentContextCache` singleton with `getCached(host): AgentContext | null`, `setCached(host, ctx): void`, `invalidate(host?): void`.

- [ ] **Step 1: Write the failing test in `src/services/office/__tests__/contextCache.test.ts`**

```typescript
import { describe, it, expect, beforeEach } from 'vitest';
import { DocumentContextCache } from '../contextCache';

describe('DocumentContextCache', () => {
  let cache: DocumentContextCache;

  beforeEach(() => {
    cache = new DocumentContextCache(5000); // 5s TTL
  });

  it('stores and retrieves context within TTL', () => {
    cache.set('Excel', { host: 'Excel', documentSummary: 'Sheet1 Data' });
    const hit = cache.get('Excel');
    expect(hit).not.toBeNull();
    expect(hit?.documentSummary).toBe('Sheet1 Data');
  });

  it('invalidates cache explicitly', () => {
    cache.set('Excel', { host: 'Excel', documentSummary: 'Sheet1 Data' });
    cache.invalidate('Excel');
    expect(cache.get('Excel')).toBeNull();
  });

  it('returns null when expired', async () => {
    const shortCache = new DocumentContextCache(10);
    shortCache.set('Word', { host: 'Word', documentSummary: 'Doc' });
    await new Promise(r => setTimeout(r, 20));
    expect(shortCache.get('Word')).toBeNull();
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- src/services/office/__tests__/contextCache.test.ts`
Expected: FAIL with module not found.

- [ ] **Step 3: Implement `src/services/office/contextCache.ts`**

```typescript
import { AgentContext } from '../../agents/types';
import { HostType } from '../../types';

interface CacheEntry {
  context: AgentContext;
  timestamp: number;
}

export class DocumentContextCache {
  private cache = new Map<HostType, CacheEntry>();
  private ttlMs: number;

  constructor(ttlMs = 15000) {
    this.ttlMs = ttlMs;
  }

  get(host: HostType): AgentContext | null {
    const entry = this.cache.get(host);
    if (!entry) return null;
    if (Date.now() - entry.timestamp > this.ttlMs) {
      this.cache.delete(host);
      return null;
    }
    return entry.context;
  }

  set(host: HostType, context: AgentContext): void {
    this.cache.set(host, {
      context,
      timestamp: Date.now(),
    });
  }

  invalidate(host?: HostType): void {
    if (host) {
      this.cache.delete(host);
    } else {
      this.cache.clear();
    }
  }
}

export const documentContextCache = new DocumentContextCache();
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npm test -- src/services/office/__tests__/contextCache.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit changes**

```bash
git add src/services/office/contextCache.ts src/services/office/__tests__/contextCache.test.ts
git commit -m "feat(office): introduce DocumentContextCache for low-latency context retrieval"
```

---

### Task 3: Autonomous ReAct Engine with Safety Limits & Self-Correction

**Files:**
- Create: `src/agents/coordinator/reactEngine.ts`
- Test: `src/agents/coordinator/__tests__/reactEngine.test.ts`

**Interfaces:**
- Consumes: `ILLMProvider`, `IAgent`, `ProviderConfig`, `ChatMessage`, `ToolCall`
- Produces: `ReActExecutionEngine` with `runLoop(params, callbacks): Promise<ReActRunResult>`

- [ ] **Step 1: Write the failing test in `src/agents/coordinator/__tests__/reactEngine.test.ts`**

```typescript
import { describe, it, expect, vi } from 'vitest';
import { ReActExecutionEngine } from '../reactEngine';
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
  });

  it('runs multi-turn loop and performs self-correction on tool error', async () => {
    let callCount = 0;
    const mockProvider: ILLMProvider = {
      id: 'mock',
      name: 'Mock',
      testConnection: async () => ({ success: true, message: 'OK' }),
      sendMessage: async function* () {
        callCount++;
        if (callCount === 1) {
          yield {
            type: 'tool_call',
            toolCall: { id: 'call_1', name: 'write_cells', arguments: { range: 'INVALID' }, status: 'pending' },
          };
        } else {
          yield { type: 'content_delta', delta: 'Berhasil diperbaiki.' };
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
        return { success: true };
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
    expect(result.finalContent).toContain('Berhasil diperbaiki.');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- src/agents/coordinator/__tests__/reactEngine.test.ts`
Expected: FAIL with module not found.

- [ ] **Step 3: Implement `src/agents/coordinator/reactEngine.ts`**

```typescript
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
}

export interface ReActRunResult {
  completed: boolean;
  iterations: number;
  finalContent: string;
  allMessages: ChatMessage[];
}

export class ReActExecutionEngine {
  async runLoop(params: ReActParams, callbacks?: ReActCallbacks): Promise<ReActRunResult> {
    const maxIterations = params.maxIterations || 6;
    let iterations = 0;
    const conversation: ChatMessage[] = [...params.initialMessages];
    let finalContent = '';
    const tools = params.specialist.getTools();

    while (iterations < maxIterations) {
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
        if (chunk.type === 'content_delta' && chunk.delta) {
          assistantMsg.content += chunk.delta;
          callbacks?.onContentDelta?.(chunk.delta);
        } else if (chunk.type === 'tool_call' && chunk.toolCall) {
          assistantMsg.toolCalls?.push(chunk.toolCall);
        } else if (chunk.type === 'error') {
          assistantMsg.content += `\n[Error: ${chunk.error}]`;
        }
      }

      conversation.push(assistantMsg);
      finalContent = assistantMsg.content;

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
        callbacks?.onStepProgress?.(iterations, maxIterations, `Mengeksekusi tool ${toolCall.name}...`);
        const execResult = await params.specialist.executeTool(toolCall, params.context || { host: params.specialist.hostType });
        toolCall.status = execResult.success ? 'applied' : 'failed';
        toolCall.error = execResult.error;
        callbacks?.onToolExecuted?.(toolCall, execResult.success, execResult.result || execResult.error);

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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npm test -- src/agents/coordinator/__tests__/reactEngine.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit changes**

```bash
git add src/agents/coordinator/reactEngine.ts src/agents/coordinator/__tests__/reactEngine.test.ts
git commit -m "feat(agent): implement ReActExecutionEngine with safety bounds and self-correction"
```

---

### Task 4: Excel Native Capabilities Expansion (`clean_data` & `apply_conditional_formatting`)

**Files:**
- Modify: `src/services/office/types.ts:40-60`
- Modify: `src/services/office/excelDriver.ts`
- Modify: `src/services/office/mockDriver.ts`
- Modify: `src/agents/excel/excelAgent.ts`
- Test: `src/agents/__tests__/specialistAgents.test.ts`

**Interfaces:**
- Consumes: `ExcelDriver`, `MockOfficeDriver`
- Produces: `cleanData(options)` and `applyConditionalFormatting(options)` methods and tools registered in `ExcelAgent`.

- [x] **Step 1: Write the failing test in `src/agents/__tests__/specialistAgents.test.ts`**

Add tests asserting `clean_data` and `apply_conditional_formatting` tool definitions exist and execute properly:
```typescript
it('executes clean_data tool and removes duplicate rows', async () => {
  const agent = new ExcelAgent();
  const res = await agent.executeTool({
    id: 'c1',
    name: 'clean_data',
    arguments: { removeDuplicates: true, trimWhitespace: true },
    status: 'pending',
  }, { host: 'Excel' });

  expect(res.success).toBe(true);
  expect(res.result).toBeDefined();
});

it('executes apply_conditional_formatting tool', async () => {
  const agent = new ExcelAgent();
  const res = await agent.executeTool({
    id: 'cf1',
    name: 'apply_conditional_formatting',
    arguments: { range: 'A1:A10', type: 'color_scale' },
    status: 'pending',
  }, { host: 'Excel' });

  expect(res.success).toBe(true);
});
```

- [x] **Step 2: Run test to verify it fails**

Run: `npm test -- src/agents/__tests__/specialistAgents.test.ts -t "clean_data"`
Expected: FAIL (tool not found/unhandled).

- [x] **Step 3: Implement `cleanData` and `applyConditionalFormatting` in `mockDriver.ts`, `excelDriver.ts`, and `excelAgent.ts`**

In `excelAgent.ts`:
Add tools to `getTools()`:
- `clean_data` with parameters `range`, `removeDuplicates`, `trimWhitespace`, `fillEmptyValues`.
- `apply_conditional_formatting` with parameters `range`, `type` (`color_scale`, `data_bar`, `highlight_threshold`), `color`.

In `excelDriver.ts` and `mockDriver.ts`:
Implement single-sync batching:
- `cleanData`: Reads used range, computes deduplication & whitespace trimming in JS, updates values, and executes single `context.sync()`.
- `applyConditionalFormatting`: Adds conditional formatting rules to target range with single `context.sync()`.

- [x] **Step 4: Run tests to verify they pass**

Run: `npm test -- src/agents/__tests__/specialistAgents.test.ts`
Expected: PASS.

- [x] **Step 5: Commit changes**

```bash
git add src/services/office/ src/agents/excel/excelAgent.ts src/agents/__tests__/specialistAgents.test.ts
git commit -m "feat(excel): add clean_data and apply_conditional_formatting native tools with single-sync batching"
```

---

### Task 5: Word & PowerPoint Native Capabilities Expansion

**Files:**
- Modify: `src/services/office/types.ts`
- Modify: `src/services/office/wordDriver.ts`
- Modify: `src/services/office/pptDriver.ts`
- Modify: `src/services/office/mockDriver.ts`
- Modify: `src/agents/word/wordAgent.ts`
- Modify: `src/agents/powerpoint/pptAgent.ts`
- Test: `src/agents/__tests__/specialistAgents.test.ts`

**Interfaces:**
- Consumes: `wordDriver`, `pptDriver`, `mockDriver`
- Produces:
  - Word: `generate_structured_doc`, `polish_document_text`
  - PowerPoint: `generate_themed_deck`

- [x] **Step 1: Write the failing test in `src/agents/__tests__/specialistAgents.test.ts`**

```typescript
it('executes generate_structured_doc in WordAgent', async () => {
  const agent = new WordAgent();
  const res = await agent.executeTool({
    id: 'w1',
    name: 'generate_structured_doc',
    arguments: {
      templateType: 'SOP',
      title: 'SOP Pengajuan Cuti',
      sections: [{ heading: 'Tujuan', content: 'Standarisasi pengajuan cuti tahunan karyawan.' }],
    },
    status: 'pending',
  }, { host: 'Word' });

  expect(res.success).toBe(true);
});

it('executes generate_themed_deck in PPTAgent', async () => {
  const agent = new PPTAgent();
  const res = await agent.executeTool({
    id: 'p1',
    name: 'generate_themed_deck',
    arguments: {
      topic: 'Laporan Q3',
      theme: 'corporate_blue',
      slides: [{ title: 'Cover', layout: 'title_cover', content: ['Q3 Performance Overview'] }],
    },
    status: 'pending',
  }, { host: 'PowerPoint' });

  expect(res.success).toBe(true);
});
```

- [x] **Step 2: Run test to verify it fails**

Run: `npm test -- src/agents/__tests__/specialistAgents.test.ts -t "generate_structured_doc"`
Expected: FAIL.

- [x] **Step 3: Implement driver methods and tool handlers in `WordAgent` and `PPTAgent`**

- In `wordDriver.ts` / `mockDriver.ts`: Implement `generateStructuredDoc` (inserts title, hierarchical headings, body paragraphs, and structured bullet items) and `polishDocumentText`.
- In `pptDriver.ts` / `mockDriver.ts`: Implement `generateThemedDeck` (creates styled slides according to layout specifications).
- Register schemas in `wordAgent.ts` and `pptAgent.ts`.

- [x] **Step 4: Run tests to verify they pass**

Run: `npm test -- src/agents/__tests__/specialistAgents.test.ts`
Expected: PASS.

- [x] **Step 5: Commit changes**

```bash
git add src/services/office/ src/agents/word/wordAgent.ts src/agents/powerpoint/pptAgent.ts src/agents/__tests__/specialistAgents.test.ts
git commit -m "feat(office): add structured doc generator for Word and themed deck generator for PowerPoint"
```

---

### Task 6: Quick Action Presets Toolbar & Full UI ReAct Integration

**Files:**
- Create: `src/components/Chat/QuickActionPresets.tsx`
- Modify: `src/components/Chat/InputBar.tsx`
- Modify: `src/components/Chat/MessageBubble.tsx`
- Modify: `src/App.tsx`
- Test: `src/components/__tests__/chatAndActions.test.tsx`
- Test: `tests/app.test.tsx`

**Interfaces:**
- Consumes: `ReActExecutionEngine`, `documentContextCache`, `QuickActionPresets`
- Produces: Interactive 1-click preset buttons, live step indicator in chat messages, and seamless autonomous ReAct execution in `App.tsx`.

- [ ] **Step 1: Write failing test in `src/components/__tests__/chatAndActions.test.tsx`**

Test that presets render appropriate buttons based on host and trigger prompt emission when clicked:
```typescript
it('renders QuickActionPresets for Excel and triggers preset prompt on click', () => {
  const onSelectPreset = vi.fn();
  render(<QuickActionPresets host="Excel" onSelectPreset={onSelectPreset} disabled={false} />);

  expect(screen.getByText(/Bersihkan Data/i)).toBeInTheDocument();
  fireEvent.click(screen.getByText(/Bersihkan Data/i));
  expect(onSelectPreset).toHaveBeenCalledWith(expect.stringContaining('Bersihkan data'));
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- src/components/__tests__/chatAndActions.test.tsx -t "QuickActionPresets"`
Expected: FAIL with component not found.

- [ ] **Step 3: Implement `QuickActionPresets.tsx` and integrate into `InputBar.tsx` & `App.tsx`**

1. Create `src/components/Chat/QuickActionPresets.tsx` with customized chip presets for Excel, Word, and PowerPoint.
2. In `App.tsx`:
   - Replace single-turn manual tool handling with `ReActExecutionEngine`.
   - Use `documentContextCache` for document summary caching.
   - Invalidate cache automatically whenever a writing tool executes.
   - Update `MessageBubble.tsx` to display step progress badges (`Step [1/3]...`).

- [ ] **Step 4: Run all tests to verify entire test suite passes**

Run: `npm test -- --run`
Expected: All tests pass (>230 tests).

- [ ] **Step 5: Commit changes**

```bash
git add src/components/Chat/ src/App.tsx src/components/__tests__/chatAndActions.test.tsx tests/app.test.tsx
git commit -m "feat(ui): integrate ReActExecutionEngine, QuickActionPresets toolbar, and live step badges"
```
