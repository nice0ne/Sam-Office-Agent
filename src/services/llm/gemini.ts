import { ProviderConfig } from '../../types';
import { ChatRequest, ILLMProvider, StreamEvent } from './types';

export class GeminiProvider implements ILLMProvider {
  id = 'gemini';
  name = 'Google Gemini';

  async *sendMessage(req: ChatRequest, config: ProviderConfig): AsyncIterable<StreamEvent> {
    if (!config.apiKey) {
      yield { type: 'error', error: 'API Key Google Gemini belum diisi.' };
      return;
    }

    const model = config.selectedModel || 'gemini-2.0-flash';
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:streamGenerateContent?key=${config.apiKey}`;

    const contents = req.messages.map(m => ({
      role: m.role === 'assistant' ? 'model' : 'user',
      parts: [{ text: m.content }],
    }));

    const body: Record<string, any> = { contents };

    if (req.systemPrompt) {
      body.systemInstruction = {
        parts: [{ text: req.systemPrompt }],
      };
    }

    if (req.tools && req.tools.length > 0) {
      body.tools = [
        {
          functionDeclarations: req.tools.map(t => ({
            name: t.name,
            description: t.description,
            parameters: t.parameters,
          })),
        },
      ];
    }

    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      });

      if (!response.ok) {
        const errorText = await response.text();
        yield { type: 'error', error: `Gemini HTTP ${response.status}: ${errorText}` };
        return;
      }

      const reader = response.body?.getReader();
      if (!reader) {
        yield { type: 'error', error: 'Tidak ada stream dari Gemini API.' };
        return;
      }

      const decoder = new TextDecoder('utf-8');
      let buffer = '';

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        // Gemini returns chunks of JSON array format
        // Simple extraction of text parts and functionCalls
        const textMatches = [...buffer.matchAll(/"text":\s*"((?:[^"\\]|\\.)*)"/g)];
        if (textMatches.length > 0) {
          for (const match of textMatches) {
            try {
              const text = JSON.parse(`"${match[1]}"`);
              yield { type: 'content_delta', delta: text };
            } catch {}
          }
          buffer = '';
        }
      }

      yield { type: 'done' };
    } catch (e: any) {
      yield { type: 'error', error: e.message || 'Gagal memanggil Gemini API.' };
    }
  }

  async testConnection(config: ProviderConfig): Promise<{ success: boolean; message: string }> {
    if (!config.apiKey) {
      return { success: false, message: 'API Key Gemini belum diisi.' };
    }
    const model = config.selectedModel || 'gemini-2.0-flash';
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${config.apiKey}`;

    try {
      const res = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: 'Ping' }] }],
        }),
      });
      if (res.ok) {
        return { success: true, message: `Koneksi Google Gemini berhasil (${model})!` };
      }
      const err = await res.text();
      return { success: false, message: `Gemini Error (${res.status}): ${err}` };
    } catch (e: any) {
      return { success: false, message: `Koneksi gagal: ${e.message}` };
    }
  }
}
