import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { getLLMProvider, testProviderConnection } from '../factory';
import { OpenAICompatibleProvider } from '../openai';
import { GeminiProvider } from '../gemini';
import { AnthropicProvider } from '../anthropic';
import { GLMProvider } from '../glm';
import { ToolDefinition } from '../../../types';

describe('LLM Provider Factory', () => {
  it('returns valid adapter instances for all supported providers', () => {
    expect(getLLMProvider('gemini').id).toBe('gemini');
    expect(getLLMProvider('openai').id).toBe('openai');
    expect(getLLMProvider('claude').id).toBe('claude');
    expect(getLLMProvider('glm').id).toBe('glm');
    expect(getLLMProvider('openrouter').id).toBe('openrouter');
    expect(getLLMProvider('ollama').id).toBe('ollama');
  });

  it('falls back to gemini provider for unknown provider id', () => {
    // @ts-expect-error testing invalid provider id
    const fallback = getLLMProvider('unknown-provider');
    expect(fallback.id).toBe('gemini');
  });

  it('fails connection test gracefully if API key is missing', async () => {
    const provider = getLLMProvider('gemini');
    const result = await provider.testConnection({
      id: 'gemini',
      name: 'Google Gemini',
      apiKey: '',
      selectedModel: 'gemini-2.0-flash',
      enabled: true,
    });
    expect(result.success).toBe(false);
    expect(result.message).toContain('API Key');
  });

  it('delegates testProviderConnection to the appropriate provider', async () => {
    const result = await testProviderConnection({
      id: 'claude',
      name: 'Anthropic Claude',
      apiKey: '',
      selectedModel: 'claude-3-5-sonnet-20241022',
      enabled: true,
    });
    expect(result.success).toBe(false);
    expect(result.message).toContain('API Key');
  });
});

describe('OpenAICompatibleProvider', () => {
  const originalFetch = globalThis.fetch;

  beforeEach(() => {
    vi.restoreAllMocks();
  });

  afterEach(() => {
    globalThis.fetch = originalFetch;
  });

  it('fails connection test if API key is missing (except for ollama)', async () => {
    const openai = new OpenAICompatibleProvider('openai', 'OpenAI', 'https://api.openai.com/v1');
    const result = await openai.testConnection({
      id: 'openai',
      name: 'OpenAI',
      apiKey: '',
      selectedModel: 'gpt-4o',
      enabled: true,
    });
    expect(result.success).toBe(false);
    expect(result.message).toContain('API Key');

    const ollama = new OpenAICompatibleProvider('ollama', 'Local Ollama', 'http://localhost:11434/v1');
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: true,
      json: async () => ({ choices: [] }),
    } as any);

    const ollamaResult = await ollama.testConnection({
      id: 'ollama',
      name: 'Local Ollama',
      apiKey: '',
      selectedModel: 'llama3.2',
      enabled: true,
    });
    expect(ollamaResult.success).toBe(true);
  });

  it('succeeds connection test when API returns 200 OK', async () => {
    const openai = new OpenAICompatibleProvider('openai', 'OpenAI', 'https://api.openai.com/v1');
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: true,
    } as any);

    const result = await openai.testConnection({
      id: 'openai',
      name: 'OpenAI',
      apiKey: 'test-key',
      selectedModel: 'gpt-4o',
      enabled: true,
    });
    expect(result.success).toBe(true);
    expect(result.message).toContain('Koneksi berhasil ke OpenAI (gpt-4o)!');
  });

  it('handles HTTP error during connection test', async () => {
    const openai = new OpenAICompatibleProvider('openai', 'OpenAI', 'https://api.openai.com/v1');
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: false,
      status: 401,
      text: async () => 'Unauthorized: Invalid API key',
    } as any);

    const result = await openai.testConnection({
      id: 'openai',
      name: 'OpenAI',
      apiKey: 'bad-key',
      selectedModel: 'gpt-4o',
      enabled: true,
    });
    expect(result.success).toBe(false);
    expect(result.message).toContain('Gagal (HTTP 401)');
  });

  it('handles network error during connection test', async () => {
    const openai = new OpenAICompatibleProvider('openai', 'OpenAI', 'https://api.openai.com/v1');
    globalThis.fetch = vi.fn().mockRejectedValue(new Error('Network offline'));

    const result = await openai.testConnection({
      id: 'openai',
      name: 'OpenAI',
      apiKey: 'key',
      selectedModel: 'gpt-4o',
      enabled: true,
    });
    expect(result.success).toBe(false);
    expect(result.message).toContain('Error koneksi: Network offline');
  });

  it('streams content_delta and tool_call events from SSE', async () => {
    const openrouter = new OpenAICompatibleProvider('openrouter', 'OpenRouter', 'https://openrouter.ai/api/v1');

    const sseChunks = [
      'data: {"choices":[{"delta":{"content":"Hello"}}]}\n\n',
      'data: {"choices":[{"delta":{"content":" world"}}]}\n\n',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_123","function":{"name":"read_cell","arguments":"{\\"address\\":"}}]}}]}\n\n',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"\\"A1\\"}"}}]}}]}\n\n',
      'data: [DONE]\n\n',
    ];

    const encoder = new TextEncoder();
    let chunkIndex = 0;
    const stream = new ReadableStream({
      pull(controller) {
        if (chunkIndex < sseChunks.length) {
          controller.enqueue(encoder.encode(sseChunks[chunkIndex++]));
        } else {
          controller.close();
        }
      },
    });

    let fetchHeaders: any = null;
    let fetchBody: any = null;
    globalThis.fetch = vi.fn().mockImplementation((_url, opts) => {
      fetchHeaders = opts.headers;
      fetchBody = JSON.parse(opts.body);
      return Promise.resolve({
        ok: true,
        body: stream,
      });
    });

    const tools: ToolDefinition[] = [
      {
        name: 'read_cell',
        description: 'Read cell value',
        parameters: {
          type: 'object',
          properties: { address: { type: 'string', description: 'cell address' } },
          required: ['address'],
        },
      },
    ];

    const events = [];
    for await (const event of openrouter.sendMessage(
      {
        messages: [{ id: '1', role: 'user', content: 'What is A1?', timestamp: 12345 }],
        systemPrompt: 'You are an Office agent',
        tools,
        temperature: 0.2,
      },
      {
        id: 'openrouter',
        name: 'OpenRouter',
        apiKey: 'or-key',
        selectedModel: 'deepseek/deepseek-chat',
        enabled: true,
      }
    )) {
      events.push(event);
    }

    // Verify OpenRouter headers
    expect(fetchHeaders['HTTP-Referer']).toBe('https://github.com/Sam-Office-Agent');
    expect(fetchHeaders['X-Title']).toBe('Sam Office Agent');
    expect(fetchHeaders['Authorization']).toBe('Bearer or-key');
    expect(fetchBody.tools).toBeDefined();
    expect(fetchBody.tools[0].function.name).toBe('read_cell');
    expect(fetchBody.temperature).toBe(0.2);

    // Verify stream events
    expect(events).toEqual([
      { type: 'content_delta', delta: 'Hello' },
      { type: 'content_delta', delta: ' world' },
      {
        type: 'tool_call',
        toolCall: {
          id: 'call_123',
          name: 'read_cell',
          arguments: { address: 'A1' },
          status: 'pending',
        },
      },
      { type: 'done' },
    ]);
  });

  it('yields error event when HTTP response is not ok', async () => {
    const openai = new OpenAICompatibleProvider('openai', 'OpenAI', 'https://api.openai.com/v1');
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: false,
      status: 429,
      text: async () => 'Rate limit exceeded',
    } as any);

    const events = [];
    for await (const event of openai.sendMessage(
      { messages: [{ id: '1', role: 'user', content: 'Hi', timestamp: 1 }] },
      { id: 'openai', name: 'OpenAI', apiKey: 'k', selectedModel: 'gpt-4o', enabled: true }
    )) {
      events.push(event);
    }

    expect(events).toHaveLength(1);
    expect(events[0].type).toBe('error');
    expect(events[0].error).toContain('HTTP 429: Rate limit exceeded');
  });

  it('yields error event when fetch throws', async () => {
    const openai = new OpenAICompatibleProvider('openai', 'OpenAI', 'https://api.openai.com/v1');
    globalThis.fetch = vi.fn().mockRejectedValue(new Error('Connection aborted'));

    const events = [];
    for await (const event of openai.sendMessage(
      { messages: [{ id: '1', role: 'user', content: 'Hi', timestamp: 1 }] },
      { id: 'openai', name: 'OpenAI', apiKey: 'k', selectedModel: 'gpt-4o', enabled: true }
    )) {
      events.push(event);
    }

    expect(events).toHaveLength(1);
    expect(events[0].type).toBe('error');
    expect(events[0].error).toBe('Connection aborted');
  });
});

describe('GeminiProvider', () => {
  const originalFetch = globalThis.fetch;

  beforeEach(() => {
    vi.restoreAllMocks();
  });

  afterEach(() => {
    globalThis.fetch = originalFetch;
  });

  it('fails connection test if API key is empty', async () => {
    const gemini = new GeminiProvider();
    const result = await gemini.testConnection({
      id: 'gemini',
      name: 'Google Gemini',
      apiKey: '',
      selectedModel: 'gemini-2.0-flash',
      enabled: true,
    });
    expect(result.success).toBe(false);
    expect(result.message).toContain('API Key Gemini belum diisi.');
  });

  it('succeeds connection test when API responds 200 OK', async () => {
    const gemini = new GeminiProvider();
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: true,
    } as any);

    const result = await gemini.testConnection({
      id: 'gemini',
      name: 'Google Gemini',
      apiKey: 'gemini-key',
      selectedModel: 'gemini-2.0-flash',
      enabled: true,
    });
    expect(result.success).toBe(true);
    expect(result.message).toContain('Koneksi Google Gemini berhasil (gemini-2.0-flash)!');
  });

  it('handles connection test error when API responds with error status', async () => {
    const gemini = new GeminiProvider();
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: false,
      status: 400,
      text: async () => 'API_KEY_INVALID',
    } as any);

    const result = await gemini.testConnection({
      id: 'gemini',
      name: 'Google Gemini',
      apiKey: 'bad-key',
      selectedModel: 'gemini-2.0-flash',
      enabled: true,
    });
    expect(result.success).toBe(false);
    expect(result.message).toContain('Gemini Error (400): API_KEY_INVALID');
  });

  it('yields error event in sendMessage if API key is missing', async () => {
    const gemini = new GeminiProvider();
    const events = [];
    for await (const event of gemini.sendMessage(
      { messages: [] },
      { id: 'gemini', name: 'Google Gemini', apiKey: '', selectedModel: 'gemini-2.0-flash', enabled: true }
    )) {
      events.push(event);
    }
    expect(events).toEqual([{ type: 'error', error: 'API Key Google Gemini belum diisi.' }]);
  });

  it('streams content_delta from Gemini JSON stream', async () => {
    const gemini = new GeminiProvider();
    const chunks = [
      '[\n{\n"candidates": [{\n"content": {\n"parts": [{\n"text": "Hello "\n}]\n}\n}]\n}\n',
      ',\n{\n"candidates": [{\n"content": {\n"parts": [{\n"text": "from Gemini!"\n}]\n}\n}]\n}\n]\n',
    ];

    const encoder = new TextEncoder();
    let chunkIndex = 0;
    const stream = new ReadableStream({
      pull(controller) {
        if (chunkIndex < chunks.length) {
          controller.enqueue(encoder.encode(chunks[chunkIndex++]));
        } else {
          controller.close();
        }
      },
    });

    let requestBody: any = null;
    let requestUrl: string = '';
    globalThis.fetch = vi.fn().mockImplementation((url, opts) => {
      requestUrl = url;
      requestBody = JSON.parse(opts.body);
      return Promise.resolve({
        ok: true,
        body: stream,
      });
    });

    const events = [];
    for await (const event of gemini.sendMessage(
      {
        messages: [
          { id: '1', role: 'user', content: 'Hi', timestamp: 100 },
          { id: '2', role: 'assistant', content: 'Hello', timestamp: 200 },
          { id: '3', role: 'user', content: 'Help me', timestamp: 300 },
        ],
        systemPrompt: 'Be a helpful office assistant',
      },
      { id: 'gemini', name: 'Google Gemini', apiKey: 'test-gemini-key', selectedModel: 'gemini-2.0-flash', enabled: true }
    )) {
      events.push(event);
    }

    expect(requestUrl).toContain('gemini-2.0-flash:streamGenerateContent?key=test-gemini-key');
    expect(requestBody.systemInstruction.parts[0].text).toBe('Be a helpful office assistant');
    expect(requestBody.contents[0].role).toBe('user');
    expect(requestBody.contents[1].role).toBe('model');
    expect(requestBody.contents[2].role).toBe('user');

    expect(events).toEqual([
      { type: 'content_delta', delta: 'Hello ' },
      { type: 'content_delta', delta: 'from Gemini!' },
      { type: 'done' },
    ]);
  });
});

describe('AnthropicProvider', () => {
  const originalFetch = globalThis.fetch;

  beforeEach(() => {
    vi.restoreAllMocks();
  });

  afterEach(() => {
    globalThis.fetch = originalFetch;
  });

  it('fails connection test if API key is empty', async () => {
    const claude = new AnthropicProvider();
    const result = await claude.testConnection({
      id: 'claude',
      name: 'Anthropic Claude',
      apiKey: '',
      selectedModel: 'claude-3-5-sonnet-20241022',
      enabled: true,
    });
    expect(result.success).toBe(false);
    expect(result.message).toContain('API Key Anthropic belum diisi.');
  });

  it('succeeds connection test when API responds 200 OK', async () => {
    const claude = new AnthropicProvider();
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: true,
    } as any);

    const result = await claude.testConnection({
      id: 'claude',
      name: 'Anthropic Claude',
      apiKey: 'anthropic-key',
      selectedModel: 'claude-3-5-sonnet-20241022',
      enabled: true,
    });
    expect(result.success).toBe(true);
    expect(result.message).toContain('Koneksi Anthropic Claude berhasil!');
  });

  it('handles connection test error when API responds with error status', async () => {
    const claude = new AnthropicProvider();
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: false,
      status: 401,
      text: async () => 'Invalid x-api-key',
    } as any);

    const result = await claude.testConnection({
      id: 'claude',
      name: 'Anthropic Claude',
      apiKey: 'bad-key',
      selectedModel: 'claude-3-5-sonnet-20241022',
      enabled: true,
    });
    expect(result.success).toBe(false);
    expect(result.message).toContain('Claude Error (401): Invalid x-api-key');
  });

  it('yields error event in sendMessage if API key is missing', async () => {
    const claude = new AnthropicProvider();
    const events = [];
    for await (const event of claude.sendMessage(
      { messages: [] },
      { id: 'claude', name: 'Anthropic Claude', apiKey: '', selectedModel: 'claude-3-5-sonnet-20241022', enabled: true }
    )) {
      events.push(event);
    }
    expect(events).toEqual([{ type: 'error', error: 'API Key Anthropic belum diisi.' }]);
  });

  it('streams content_delta from SSE content_block_delta events', async () => {
    const claude = new AnthropicProvider();
    const sseChunks = [
      'data: {"type": "message_start", "message": {"id": "msg_123"}}\n\n',
      'data: {"type": "content_block_start", "index": 0, "content_block": {"type": "text", "text": ""}}\n\n',
      'data: {"type": "content_block_delta", "index": 0, "delta": {"type": "text_delta", "text": "Hello "}}\n\n',
      'data: {"type": "content_block_delta", "index": 0, "delta": {"type": "text_delta", "text": "from Claude!"}}\n\n',
      'data: {"type": "message_stop"}\n\n',
    ];

    const encoder = new TextEncoder();
    let chunkIndex = 0;
    const stream = new ReadableStream({
      pull(controller) {
        if (chunkIndex < sseChunks.length) {
          controller.enqueue(encoder.encode(sseChunks[chunkIndex++]));
        } else {
          controller.close();
        }
      },
    });

    let requestHeaders: any = null;
    let requestBody: any = null;
    globalThis.fetch = vi.fn().mockImplementation((_url, opts) => {
      requestHeaders = opts.headers;
      requestBody = JSON.parse(opts.body);
      return Promise.resolve({
        ok: true,
        body: stream,
      });
    });

    const events = [];
    for await (const event of claude.sendMessage(
      {
        messages: [{ id: '1', role: 'user', content: 'Hi Claude', timestamp: 10 }],
        systemPrompt: 'System instruction here',
      },
      { id: 'claude', name: 'Anthropic Claude', apiKey: 'claude-key', selectedModel: 'claude-3-5-sonnet-20241022', enabled: true }
    )) {
      events.push(event);
    }

    expect(requestHeaders['x-api-key']).toBe('claude-key');
    expect(requestHeaders['anthropic-version']).toBe('2023-06-01');
    expect(requestHeaders['anthropic-dangerous-direct-browser-access']).toBe('true');
    expect(requestBody.system).toBe('System instruction here');

    expect(events).toEqual([
      { type: 'content_delta', delta: 'Hello ' },
      { type: 'content_delta', delta: 'from Claude!' },
      { type: 'done' },
    ]);
  });
});

describe('GLMProvider', () => {
  it('instantiates correctly with default GLM endpoint', () => {
    const glm = new GLMProvider();
    expect(glm.id).toBe('glm');
    expect(glm.name).toBe('GLM Coding Global');
    expect(glm.defaultBaseUrl).toBe('https://open.bigmodel.cn/api/paas/v4');
  });
});
