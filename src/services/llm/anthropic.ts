import { ProviderConfig } from '../../types';
import { ChatRequest, ILLMProvider, StreamEvent } from './types';

export class AnthropicProvider implements ILLMProvider {
  id = 'claude';
  name = 'Anthropic Claude';

  async *sendMessage(req: ChatRequest, config: ProviderConfig): AsyncIterable<StreamEvent> {
    if (!config.apiKey) {
      yield { type: 'error', error: 'API Key Anthropic belum diisi.' };
      return;
    }

    const url = 'https://api.anthropic.com/v1/messages';
    const messages = req.messages.map(m => ({
      role: m.role === 'assistant' ? 'assistant' : 'user',
      content: m.content,
    }));

    const body: Record<string, any> = {
      model: config.selectedModel || 'claude-3-5-sonnet-20241022',
      messages,
      max_tokens: 4096,
      stream: true,
    };

    if (req.systemPrompt) {
      body.system = req.systemPrompt;
    }

    if (req.tools && req.tools.length > 0) {
      body.tools = req.tools.map(t => ({
        name: t.name,
        description: t.description,
        input_schema: t.parameters,
      }));
    }

    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': config.apiKey,
          'anthropic-version': '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true',
        },
        body: JSON.stringify(body),
      });

      if (!response.ok) {
        const errorText = await response.text();
        yield { type: 'error', error: `Claude HTTP ${response.status}: ${errorText}` };
        return;
      }

      const reader = response.body?.getReader();
      if (!reader) return;

      const decoder = new TextDecoder('utf-8');
      let buffer = '';
      const activeToolBlocks: Record<number, { id: string; name: string; json: string }> = {};

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop() || '';

        for (const line of lines) {
          const clean = line.trim();
          if (!clean.startsWith('data: ')) continue;
          try {
            const data = JSON.parse(clean.slice(6));

            if (data.type === 'content_block_start' && data.content_block?.type === 'tool_use') {
              activeToolBlocks[data.index] = {
                id: data.content_block.id || `call_claude_${Date.now()}_${data.index}`,
                name: data.content_block.name || '',
                json: '',
              };
            }

            if (data.type === 'content_block_delta') {
              if (data.delta?.type === 'text_delta' && data.delta?.text) {
                yield { type: 'content_delta', delta: data.delta.text };
              }
              if (data.delta?.type === 'input_json_delta' && data.delta?.partial_json) {
                if (activeToolBlocks[data.index]) {
                  activeToolBlocks[data.index].json += data.delta.partial_json;
                }
              }
            }

            if (data.type === 'content_block_stop') {
              if (activeToolBlocks[data.index]) {
                const tb = activeToolBlocks[data.index];
                delete activeToolBlocks[data.index];
                let args = {};
                try {
                  args = JSON.parse(tb.json);
                } catch {
                  args = {};
                }
                yield {
                  type: 'tool_call',
                  toolCall: {
                    id: tb.id,
                    name: tb.name,
                    arguments: args,
                    status: 'pending',
                  },
                };
              }
            }
          } catch {}
        }
      }

      for (const idx in activeToolBlocks) {
        const tb = activeToolBlocks[idx];
        let args = {};
        try {
          args = JSON.parse(tb.json);
        } catch {
          args = {};
        }
        yield {
          type: 'tool_call',
          toolCall: {
            id: tb.id,
            name: tb.name,
            arguments: args,
            status: 'pending',
          },
        };
      }

      yield { type: 'done' };
    } catch (e: any) {
      yield { type: 'error', error: e.message || 'Gagal memanggil Claude API.' };
    }
  }

  async testConnection(config: ProviderConfig): Promise<{ success: boolean; message: string }> {
    if (!config.apiKey) {
      return { success: false, message: 'API Key Anthropic belum diisi.' };
    }

    try {
      const res = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': config.apiKey,
          'anthropic-version': '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true',
        },
        body: JSON.stringify({
          model: config.selectedModel || 'claude-3-5-haiku-20241022',
          messages: [{ role: 'user', content: 'Ping' }],
          max_tokens: 5,
        }),
      });

      if (res.ok) {
        return { success: true, message: 'Koneksi Anthropic Claude berhasil!' };
      }
      const err = await res.text();
      return { success: false, message: `Claude Error (${res.status}): ${err}` };
    } catch (e: any) {
      return { success: false, message: `Koneksi gagal: ${e.message}` };
    }
  }
}
