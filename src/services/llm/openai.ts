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

  async *sendMessage(req: ChatRequest, config: ProviderConfig): AsyncIterable<StreamEvent> {
    if (!config.apiKey && this.id !== 'ollama') {
      yield { type: 'error', error: `API Key ${this.name} belum diisi.` };
      return;
    }

    const baseUrl = (config.baseUrl || this.defaultBaseUrl).replace(/\/+$/, '');
    const url = `${baseUrl}/chat/completions`;

    const messages = [];
    if (req.systemPrompt) {
      messages.push({ role: 'system', content: req.systemPrompt });
    }
    for (const m of req.messages) {
      messages.push({ role: m.role, content: m.content });
    }

    const body: Record<string, any> = {
      model: config.selectedModel,
      messages,
      stream: true,
      temperature: req.temperature ?? 0.7,
    };

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
      'Authorization': `Bearer ${config.apiKey}`,
    };

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
    if (!config.apiKey && this.id !== 'ollama') {
      return { success: false, message: 'API Key belum diisi.' };
    }

    const baseUrl = (config.baseUrl || this.defaultBaseUrl).replace(/\/+$/, '');
    try {
      const headers: Record<string, string> = {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${config.apiKey}`,
      };

      if (this.id === 'openrouter') {
        headers['HTTP-Referer'] = 'https://github.com/Sam-Office-Agent';
        headers['X-Title'] = 'Sam Office Agent';
      }

      const res = await fetch(`${baseUrl}/chat/completions`, {
        method: 'POST',
        headers,
        body: JSON.stringify({
          model: config.selectedModel,
          messages: [{ role: 'user', content: 'Ping' }],
          max_tokens: 5,
        }),
      });

      if (res.ok) {
        return { success: true, message: `Koneksi berhasil ke ${this.name} (${config.selectedModel})!` };
      }
      const err = await res.text();
      return { success: false, message: `Gagal (HTTP ${res.status}): ${err}` };
    } catch (e: any) {
      return { success: false, message: `Error koneksi: ${e.message}` };
    }
  }
}
