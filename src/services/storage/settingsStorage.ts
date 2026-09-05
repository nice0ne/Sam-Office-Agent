import { ExecutionMode, ProviderConfig, ProviderId } from '../../types';

const STORAGE_KEY = 'sam_office_settings_v1';

export interface AppSettings {
  activeProviderId: ProviderId;
  executionMode: ExecutionMode;
  theme: 'light' | 'dark' | 'system';
  providers: Record<ProviderId, ProviderConfig>;
}

export const DEFAULT_PROVIDERS: Record<ProviderId, ProviderConfig> = {
  gemini: {
    id: 'gemini',
    name: 'Google Gemini',
    apiKey: '',
    selectedModel: 'gemini-2.0-flash',
    enabled: true,
  },
  openai: {
    id: 'openai',
    name: 'OpenAI',
    apiKey: '',
    selectedModel: 'gpt-4o',
    enabled: true,
  },
  claude: {
    id: 'claude',
    name: 'Anthropic Claude',
    apiKey: '',
    selectedModel: 'claude-3-5-sonnet-20241022',
    enabled: true,
  },
  glm: {
    id: 'glm',
    name: 'GLM Coding Global (Zhipu)',
    apiKey: '',
    baseUrl: 'https://open.bigmodel.cn/api/paas/v4',
    selectedModel: 'glm-4-plus',
    enabled: true,
  },
  openrouter: {
    id: 'openrouter',
    name: 'OpenRouter / DeepSeek',
    apiKey: '',
    baseUrl: 'https://openrouter.ai/api/v1',
    selectedModel: 'deepseek/deepseek-chat',
    enabled: true,
  },
  ollama: {
    id: 'ollama',
    name: 'Local Ollama',
    apiKey: 'ollama',
    baseUrl: 'http://localhost:11434/v1',
    selectedModel: 'qwen2.5-coder',
    enabled: false,
  },
};

export function getSettings(): AppSettings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return {
        activeProviderId: 'gemini',
        executionMode: 'copilot',
        theme: 'system',
        providers: { ...DEFAULT_PROVIDERS },
      };
    }
    const parsed = JSON.parse(raw);
    return {
      activeProviderId: parsed.activeProviderId || 'gemini',
      executionMode: parsed.executionMode || 'copilot',
      theme: parsed.theme || 'system',
      providers: { ...DEFAULT_PROVIDERS, ...(parsed.providers || {}) },
    };
  } catch {
    return {
      activeProviderId: 'gemini',
      executionMode: 'copilot',
      theme: 'system',
      providers: { ...DEFAULT_PROVIDERS },
    };
  }
}

export function saveSettings(settings: AppSettings): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}

export function saveProviderConfig(config: ProviderConfig): void {
  const current = getSettings();
  current.providers[config.id] = config;
  saveSettings(current);
}

export function getActiveProvider(): ProviderConfig {
  const settings = getSettings();
  return settings.providers[settings.activeProviderId] || settings.providers.gemini;
}

export function setActiveProvider(id: ProviderId): void {
  const settings = getSettings();
  settings.activeProviderId = id;
  saveSettings(settings);
}

export function getExecutionMode(): ExecutionMode {
  return getSettings().executionMode;
}

export function setExecutionMode(mode: ExecutionMode): void {
  const settings = getSettings();
  settings.executionMode = mode;
  saveSettings(settings);
}
