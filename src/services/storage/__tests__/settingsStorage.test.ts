import { describe, it, expect, beforeEach } from 'vitest';
import {
  getSettings,
  saveSettings,
  saveProviderConfig,
  getActiveProvider,
  setActiveProvider,
  getExecutionMode,
  setExecutionMode,
} from '../settingsStorage';

describe('Settings Storage Service', () => {
  beforeEach(() => {
    localStorage.clear();
  });

  it('initializes default provider configurations', () => {
    const settings = getSettings();
    expect(settings.providers.gemini).toBeDefined();
    expect(settings.providers.openai).toBeDefined();
    expect(settings.providers.claude).toBeDefined();
    expect(settings.providers.glm).toBeDefined();
    expect(settings.providers.openrouter).toBeDefined();
    expect(settings.providers.ollama).toBeDefined();
    expect(settings.activeProviderId).toBe('gemini');
    expect(settings.executionMode).toBe('copilot');
    expect(settings.theme).toBe('system');
  });

  it('saves and retrieves updated provider config', () => {
    saveProviderConfig({
      id: 'openai',
      name: 'OpenAI',
      apiKey: 'sk-test-key-12345',
      selectedModel: 'gpt-4o',
      enabled: true,
    });

    const active = getActiveProvider();
    expect(active.apiKey).toBe(''); // Gemini is still active
    setActiveProvider('openai');
    expect(getActiveProvider().apiKey).toBe('sk-test-key-12345');
  });

  it('toggles execution mode between copilot and autopilot', () => {
    expect(getExecutionMode()).toBe('copilot');
    setExecutionMode('autopilot');
    expect(getExecutionMode()).toBe('autopilot');
  });

  it('handles corrupt JSON in localStorage gracefully', () => {
    localStorage.setItem('sam_office_settings_v1', 'invalid-json{{{');
    const settings = getSettings();
    expect(settings.activeProviderId).toBe('gemini');
    expect(settings.executionMode).toBe('copilot');
    expect(settings.providers.gemini).toBeDefined();
  });

  it('merges missing providers if stored settings are partial', () => {
    localStorage.setItem(
      'sam_office_settings_v1',
      JSON.stringify({
        activeProviderId: 'claude',
        providers: {
          claude: {
            id: 'claude',
            name: 'Anthropic Claude',
            apiKey: 'sk-ant-test',
            selectedModel: 'claude-3-5-sonnet-20241022',
            enabled: true,
          },
        },
      })
    );
    const settings = getSettings();
    expect(settings.activeProviderId).toBe('claude');
    expect(settings.providers.claude.apiKey).toBe('sk-ant-test');
    expect(settings.providers.gemini).toBeDefined();
    expect(settings.executionMode).toBe('copilot');
  });

  it('saves entire settings object correctly', () => {
    const current = getSettings();
    current.theme = 'dark';
    saveSettings(current);

    const reloaded = getSettings();
    expect(reloaded.theme).toBe('dark');
  });
});
