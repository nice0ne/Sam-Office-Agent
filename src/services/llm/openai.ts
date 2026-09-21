import { ProviderConfig } from '../../types';
import { ChatRequest, ILLMProvider, StreamEvent } from './types';

export class OpenAICompatibleProvider implements ILLMProvider {
  id: string;
  name: string;
  defaultBaseUrl: string;

  constructor(id: string, name: string, defaultBaseUrl: string) {
    this.id = id;
    this.name = name;
    this.defaultBaseUrl = defaultBaseUrl;
  }

  private normalizeBaseUrl(rawUrl?: string): string {
    let url = (rawUrl || this.defaultBaseUrl).trim().replace(/\/+$/, '');
    if (!url.startsWith('http://') && !url.startsWith('https://')) {
      url = `https://${url}`;
    }
    const path = url.replace(/^https?:\/\/[^/]+/, '');
    if (!path || path === '') {
      url = `${url}/v1`;
    }
    return url;
  }

  private isKeyRequired(): boolean {
    return this.id === 'openai';
  }

  async *sendMessage(req: ChatRequest, config: ProviderConfig): AsyncIterable<StreamEvent> {
    if (!config.apiKey && this.isKeyRequired()) {
      yield { type: 'error', error: `API Key ${this.name} belum diisi.` };
      return;
    }

    const baseUrl = this.normalizeBaseUrl(config.baseUrl);
    const url = `${baseUrl}/chat/completions`;

    const messages = [];
    if (req.systemPrompt) {
      messages.push({ role: 'system', content: req.systemPrompt });
    }
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

    const isReasoning = /^(o1|o3|deepseek-r1)/i.test(config.selectedModel);
    const body: Record<string, any> = {
      model: config.selectedModel || 'gpt-4o-mini',
      messages,
      stream: true,
    };
    if (!isReasoning) {
      body.temperature = req.temperature ?? 0.7;
    }

    if (req.tools && req.tools.length > 0) {
      body.tools = req.tools.map(t => ({
        type: 'function',
        function: {
          name: t.name,
          description: t.description,
          parameters: t.parameters,
        },
      }));
      body.tool_choice = 'auto';
    }

    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
    };

    const effectiveKey = config.apiKey?.trim();
    if (effectiveKey) {
      headers['Authorization'] = `Bearer ${effectiveKey}`;
    } else if (this.id === 'ollama') {
      headers['Authorization'] = 'Bearer ollama';
    }

    if (this.id === 'openrouter') {
      headers['HTTP-Referer'] = 'https://github.com/Sam-Office-Agent';
      headers['X-Title'] = 'Sam Office Agent';
    }

    try {
      const response = await fetch(url, {
        method: 'POST',
        headers,
        body: JSON.stringify(body),
      });

      if (!response.ok) {
        const errorText = await response.text();
        yield { type: 'error', error: `HTTP ${response.status}: ${errorText}` };
        return;
      }

      const reader = response.body?.getReader();
      if (!reader) {
        yield { type: 'error', error: 'No response body stream available.' };
        return;
      }

      const decoder = new TextDecoder('utf-8');
      let buffer = '';
      const pendingToolCalls: Record<number, { id: string; name: string; arguments: string }> = {};

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop() || '';

        for (const line of lines) {
          const clean = line.trim();
          if (!clean || !clean.startsWith('data: ')) continue;
          if (clean === 'data: [DONE]') continue;

          try {
            const parsed = JSON.parse(clean.slice(6));
            const delta = parsed.choices?.[0]?.delta;
            if (delta?.content) {
              yield { type: 'content_delta', delta: delta.content };
            }

            if (delta?.tool_calls) {
              for (const tc of delta.tool_calls) {
                const idx = tc.index ?? 0;
                if (!pendingToolCalls[idx]) {
                  pendingToolCalls[idx] = {
                    id: tc.id || `call_${Date.now()}_${idx}`,
                    name: '',
                    arguments: '',
                  };
                }
                if (tc.id) pendingToolCalls[idx].id = tc.id;
                if (tc.function?.name) pendingToolCalls[idx].name += tc.function.name;
                if (tc.function?.arguments) pendingToolCalls[idx].arguments += tc.function.arguments;
              }
            }
          } catch {
            // Ignore stream parse chunks
          }
        }
      }

      for (const idx in pendingToolCalls) {
        const ptc = pendingToolCalls[idx];
        let args = {};
        try {
          args = JSON.parse(ptc.arguments);
        } catch {
          args = {};
        }
        yield {
          type: 'tool_call',
          toolCall: {
            id: ptc.id,
            name: ptc.name,
            arguments: args,
            status: 'pending',
          },
        };
      }

      yield { type: 'done' };
    } catch (err: any) {
      yield { type: 'error', error: err?.message || 'Gagal terhubung ke API OpenAI' };
    }
  }

  async testConnection(config: ProviderConfig): Promise<{ success: boolean; message: string }> {
    if (!config.apiKey?.trim() && this.isKeyRequired()) {
      return { success: false, message: 'API Key belum diisi.' };
    }

    const baseUrl = this.normalizeBaseUrl(config.baseUrl);
    const url = `${baseUrl}/chat/completions`;

    try {
      const headers: Record<string, string> = {
        'Content-Type': 'application/json',
      };
      if (config.apiKey?.trim()) {
        headers['Authorization'] = `Bearer ${config.apiKey.trim()}`;
      } else if (this.id === 'ollama') {
        headers['Authorization'] = 'Bearer ollama';
      }

      if (this.id === 'openrouter') {
        headers['HTTP-Referer'] = 'https://github.com/Sam-Office-Agent';
        headers['X-Title'] = 'Sam Office Agent';
      }

      const isReasoning = /^(o1|o3|deepseek-r1)/i.test(config.selectedModel);
      const testBody: Record<string, any> = {
        model: config.selectedModel || 'gpt-4o-mini',
        messages: [{ role: 'user', content: 'Ping' }],
      };
      if (isReasoning) {
        testBody.max_completion_tokens = 10;
      } else {
        testBody.max_tokens = 5;
      }

      const res = await fetch(url, {
        method: 'POST',
        headers,
        body: JSON.stringify(testBody),
      });

      if (res.ok) {
        const json = typeof res.json === 'function' ? await res.json().catch(() => null) : null;
        const reply = json?.choices?.[0]?.message?.content?.trim();
        const detail = reply ? ` (Balasan model: "${reply.slice(0, 30)}")` : '';
        const targetName = config.name || this.name;
        return {
          success: true,
          message: `Koneksi berhasil ke ${targetName} (${config.selectedModel || 'default'})${detail}!`,
        };
      }

      const errText = typeof res.text === 'function' ? await res.text().catch(() => '') : '';
      let friendlyError = `HTTP ${res.status}`;
      if (res.status === 401) {
        friendlyError = 'HTTP 401 (API Key salah atau tidak memiliki izin akses)';
      } else if (res.status === 404) {
        friendlyError = `HTTP 404 (Endpoint ${url} tidak ditemukan. Periksa Base URL)`;
      } else if (res.status === 403) {
        friendlyError = 'HTTP 403 (Akses ditolak atau saldo/kuota habis)';
      } else if (res.status === 429) {
        friendlyError = 'HTTP 429 (Rate limit terlampaui)';
      }
      return {
        success: false,
        message: `Gagal (HTTP ${res.status}): ${friendlyError}.${errText ? ` ${errText.slice(0, 150)}` : ''}`,
      };
    } catch (e: any) {
      const isCorsOrOffline = e.name === 'TypeError' || e.message?.includes('fetch') || e.message?.includes('Failed');
      return {
        success: false,
        message: isCorsOrOffline
          ? `Error jaringan / CORS (${baseUrl}): Pastikan server aktif dan mengizinkan CORS.`
          : `Error koneksi: ${e.message}`,
      };
    }
  }
}
