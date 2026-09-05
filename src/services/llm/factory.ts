import { ProviderConfig, ProviderId } from '../../types';
import { AnthropicProvider } from './anthropic';
import { GeminiProvider } from './gemini';
import { OpenAICompatibleProvider } from './openai';
import { GLMProvider } from './glm';
import { ILLMProvider } from './types';

const providers: Record<ProviderId, ILLMProvider> = {
  gemini: new GeminiProvider(),
  openai: new OpenAICompatibleProvider('openai', 'OpenAI', 'https://api.openai.com/v1'),
  claude: new AnthropicProvider(),
  glm: new GLMProvider(),
  openrouter: new OpenAICompatibleProvider('openrouter', 'OpenRouter', 'https://openrouter.ai/api/v1'),
  ollama: new OpenAICompatibleProvider('ollama', 'Local Ollama', 'http://localhost:11434/v1'),
};

export function getLLMProvider(id: ProviderId): ILLMProvider {
  return providers[id] || providers.gemini;
}

export async function testProviderConnection(config: ProviderConfig): Promise<{ success: boolean; message: string }> {
  const provider = getLLMProvider(config.id);
  return provider.testConnection(config);
}
