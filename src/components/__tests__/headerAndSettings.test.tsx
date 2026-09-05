import React from 'react';
import { renderToString } from 'react-dom/server';
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { Header } from '../Common/Header';
import { SettingsModal } from '../Settings/SettingsModal';
import {
  getSettings,
  saveProviderConfig,
  setActiveProvider,
} from '../../services/storage/settingsStorage';
import * as llmFactory from '../../services/llm/factory';

// Helper to strip React SSR comment separators for clean string assertions
function stripComments(html: string): string {
  return html.replace(/<!--.*?-->/g, '');
}

// Helper to simulate component render with React hooks in node environment
function createHookHarness<P>(Component: React.FC<P>, props: P) {
  const states: any[] = [];
  let stateIdx = 0;
  const effectSlots: { effect: () => void | (() => void); deps?: any[] }[] = [];
  let effectIdx = 0;

  const dispatcher = {
    useState: (initial: any) => {
      const i = stateIdx++;
      if (!(i in states)) {
        states[i] = typeof initial === 'function' ? initial() : initial;
      }
      const setState = (next: any) => {
        states[i] = typeof next === 'function' ? next(states[i]) : next;
      };
      return [states[i], setState];
    },
    useEffect: (effect: () => void | (() => void), deps?: any[]) => {
      const slot = effectSlots[effectIdx];
      let shouldRun = false;

      if (!slot) {
        shouldRun = true;
      } else if (!deps) {
        shouldRun = true;
      } else if (
        deps.length !== slot.deps?.length ||
        deps.some((d, idx) => !Object.is(d, slot.deps?.[idx]))
      ) {
        shouldRun = true;
      }

      effectSlots[effectIdx] = { effect, deps };
      effectIdx++;

      if (shouldRun) {
        effect();
      }
    },
    useRef: (val: any) => ({ current: val }),
    useCallback: (fn: any) => fn,
    useMemo: (fn: any) => fn(),
  };

  const render = (currentProps: P = props): any => {
    stateIdx = 0;
    effectIdx = 0;
    const prev = (React as any).__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED.ReactCurrentDispatcher.current;
    try {
      (React as any).__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED.ReactCurrentDispatcher.current = dispatcher;
      return Component(currentProps);
    } finally {
      (React as any).__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED.ReactCurrentDispatcher.current = prev;
    }
  };

  return { render, getStates: () => states };
}

describe('Header Component', () => {
  it('renders host badge and icon for Excel', () => {
    const html = renderToString(
      <Header
        host="Excel"
        activeProviderId="gemini"
        executionMode="copilot"
        onToggleMode={() => {}}
        onOpenSettings={() => {}}
      />
    );
    expect(html).toContain('Excel');
    expect(html).toContain('Sam Office');
    expect(html).toContain('gemini');
    expect(html).toContain('text-office-excel');
  });

  it('renders host badge and icon for Word', () => {
    const html = renderToString(
      <Header
        host="Word"
        activeProviderId="openai"
        executionMode="copilot"
        onToggleMode={() => {}}
        onOpenSettings={() => {}}
      />
    );
    expect(html).toContain('Word');
    expect(html).toContain('openai');
    expect(html).toContain('text-office-word');
  });

  it('renders host badge and icon for PowerPoint', () => {
    const html = renderToString(
      <Header
        host="PowerPoint"
        activeProviderId="claude"
        executionMode="copilot"
        onToggleMode={() => {}}
        onOpenSettings={() => {}}
      />
    );
    expect(html).toContain('PowerPoint');
    expect(html).toContain('claude');
    expect(html).toContain('text-office-ppt');
  });

  it('renders fallback icon for BrowserDev host', () => {
    const html = renderToString(
      <Header
        host="BrowserDev"
        activeProviderId="ollama"
        executionMode="copilot"
        onToggleMode={() => {}}
        onOpenSettings={() => {}}
      />
    );
    expect(html).toContain('BrowserDev');
    expect(html).toContain('ollama');
    expect(html).toContain('text-office-agent');
  });

  it('renders copilot mode controls and button state', () => {
    const html = renderToString(
      <Header
        host="Excel"
        activeProviderId="gemini"
        executionMode="copilot"
        onToggleMode={() => {}}
        onOpenSettings={() => {}}
      />
    );
    expect(html).toContain('Copilot');
    expect(html).toContain('Mode Copilot (Preview dulu)');
  });

  it('renders autopilot mode controls and button state', () => {
    const html = renderToString(
      <Header
        host="Excel"
        activeProviderId="gemini"
        executionMode="autopilot"
        onToggleMode={() => {}}
        onOpenSettings={() => {}}
      />
    );
    expect(html).toContain('Auto');
    expect(html).toContain('Mode Autopilot (Langsung eksekusi)');
    expect(html).toContain('bg-amber-100');
  });

  it('triggers onToggleMode callback when mode button is clicked', () => {
    const onToggleMode = vi.fn();
    const onOpenSettings = vi.fn();
    const vdom = Header({
      host: 'Excel',
      activeProviderId: 'gemini',
      executionMode: 'copilot',
      onToggleMode,
      onOpenSettings,
    }) as React.ReactElement;

    const rightDiv = vdom.props.children[1];
    const modeButton = rightDiv.props.children[0];
    modeButton.props.onClick();

    expect(onToggleMode).toHaveBeenCalledTimes(1);
    expect(onOpenSettings).not.toHaveBeenCalled();
  });

  it('triggers onOpenSettings callback when settings button is clicked', () => {
    const onToggleMode = vi.fn();
    const onOpenSettings = vi.fn();
    const vdom = Header({
      host: 'Excel',
      activeProviderId: 'gemini',
      executionMode: 'copilot',
      onToggleMode,
      onOpenSettings,
    }) as React.ReactElement;

    const rightDiv = vdom.props.children[1];
    const settingsButton = rightDiv.props.children[1];
    settingsButton.props.onClick();

    expect(onOpenSettings).toHaveBeenCalledTimes(1);
    expect(onToggleMode).not.toHaveBeenCalled();
  });
});

describe('SettingsModal Component', () => {
  beforeEach(() => {
    localStorage.clear();
  });

  it('returns null when isOpen is false', () => {
    const html = renderToString(
      <SettingsModal isOpen={false} onClose={() => {}} onSaved={() => {}} />
    );
    expect(html).toBe('');
  });

  it('renders modal content when isOpen is true', () => {
    const html = renderToString(
      <SettingsModal isOpen={true} onClose={() => {}} onSaved={() => {}} />
    );
    expect(html).toContain('Pengaturan BYOK AI');
    expect(html).toContain('Google');
    expect(html).toContain('OpenAI');
    expect(html).toContain('Anthropic');
    expect(html).toContain('GLM');
    expect(html).toContain('OpenRouter');
    expect(html).toContain('Local');
    expect(html).toContain('Masukkan API Key Anda...');
    expect(html).toContain('Nama Model');
    expect(html).toContain('Custom Base URL / Proxy');
    expect(html).toContain('Test Koneksi');
    expect(html).toContain('Simpan &amp; Gunakan');
  });

  it('renders with existing stored active provider and config', () => {
    saveProviderConfig({
      id: 'openai',
      name: 'OpenAI',
      apiKey: 'sk-existing-test-key',
      selectedModel: 'gpt-4o-mini',
      baseUrl: 'https://custom.openai.proxy/v1',
      enabled: true,
    });
    setActiveProvider('openai');

    const html = renderToString(
      <SettingsModal isOpen={true} onClose={() => {}} onSaved={() => {}} />
    );
    const cleaned = stripComments(html);
    expect(cleaned).toContain('API Key (OpenAI)');
    expect(cleaned).toContain('sk-existing-test-key');
    expect(cleaned).toContain('gpt-4o-mini');
    expect(cleaned).toContain('https://custom.openai.proxy/v1');
  });

  it('allows switching provider tabs', () => {
    const harness = createHookHarness(SettingsModal, {
      isOpen: true,
      onClose: () => {},
      onSaved: () => {},
    });

    let vdom = harness.render();
    const modalContainer = vdom.props.children;
    const tabsContainer = modalContainer.props.children[1];
    const tabButtons = tabsContainer.props.children;

    // Initially tab 0 is Gemini
    const bodyContainer = modalContainer.props.children[2];
    const keyLabel = bodyContainer.props.children[0].props.children[0];
    expect(stripComments(renderToString(keyLabel))).toContain('API Key (Google Gemini)');

    // Click tab 1 (OpenAI)
    const openAiTab = tabButtons[1];
    openAiTab.props.onClick();

    // Re-render
    vdom = harness.render();
    const updatedBody = vdom.props.children.props.children[2];
    const updatedKeyLabel = updatedBody.props.children[0].props.children[0];
    expect(stripComments(renderToString(updatedKeyLabel))).toContain('API Key (OpenAI)');
  });

  it('handles API key, model, and base URL input changes', () => {
    const harness = createHookHarness(SettingsModal, {
      isOpen: true,
      onClose: () => {},
      onSaved: () => {},
    });

    let vdom = harness.render();
    let body = vdom.props.children.props.children[2];
    const keyInput = body.props.children[0].props.children[1];
    const modelInput = body.props.children[1].props.children[1];
    const baseUrlInput = body.props.children[2].props.children[1];

    // Change API key
    keyInput.props.onChange({ target: { value: 'sk-new-gemini-key' } });
    // Change Model
    modelInput.props.onChange({ target: { value: 'gemini-1.5-flash' } });
    // Change Base URL
    baseUrlInput.props.onChange({ target: { value: 'https://custom.gemini.proxy/v1' } });

    vdom = harness.render();
    body = vdom.props.children.props.children[2];
    const updatedKeyInput = body.props.children[0].props.children[1];
    const updatedModelInput = body.props.children[1].props.children[1];
    const updatedBaseUrlInput = body.props.children[2].props.children[1];

    expect(updatedKeyInput.props.value).toBe('sk-new-gemini-key');
    expect(updatedModelInput.props.value).toBe('gemini-1.5-flash');
    expect(updatedBaseUrlInput.props.value).toBe('https://custom.gemini.proxy/v1');
  });

  it('tests connection and displays success feedback', async () => {
    const testSpy = vi.spyOn(llmFactory, 'testProviderConnection').mockResolvedValue({
      success: true,
      message: 'Koneksi Google Gemini berhasil (gemini-2.0-flash)!',
    });

    const harness = createHookHarness(SettingsModal, {
      isOpen: true,
      onClose: () => {},
      onSaved: () => {},
    });

    let vdom = harness.render();
    const footer = vdom.props.children.props.children[3];
    const testBtn = footer.props.children[0];

    await testBtn.props.onClick();

    expect(testSpy).toHaveBeenCalled();

    vdom = harness.render();
    const body = vdom.props.children.props.children[2];
    const feedback = body.props.children[3];
    expect(feedback).toBeTruthy();
    expect(renderToString(feedback)).toContain('Koneksi Google Gemini berhasil');

    testSpy.mockRestore();
  });

  it('tests connection and displays error feedback on failure', async () => {
    const testSpy = vi.spyOn(llmFactory, 'testProviderConnection').mockResolvedValue({
      success: false,
      message: 'API Key Gemini belum diisi.',
    });

    const harness = createHookHarness(SettingsModal, {
      isOpen: true,
      onClose: () => {},
      onSaved: () => {},
    });

    let vdom = harness.render();
    const footer = vdom.props.children.props.children[3];
    const testBtn = footer.props.children[0];

    await testBtn.props.onClick();

    vdom = harness.render();
    const body = vdom.props.children.props.children[2];
    const feedback = body.props.children[3];
    expect(feedback).toBeTruthy();
    expect(renderToString(feedback)).toContain('API Key Gemini belum diisi.');

    testSpy.mockRestore();
  });

  it('saves all edited provider configurations across tabs on "Simpan & Gunakan"', () => {
    const onClose = vi.fn();
    const onSaved = vi.fn();

    const harness = createHookHarness(SettingsModal, {
      isOpen: true,
      onClose,
      onSaved,
    });

    // 1. Edit Gemini (activeTab = gemini)
    let vdom = harness.render();
    let body = vdom.props.children.props.children[2];
    const geminiKeyInput = body.props.children[0].props.children[1];
    geminiKeyInput.props.onChange({ target: { value: 'gemini-multi-tab-key' } });

    // 2. Switch tab to OpenAI
    const tabsContainer = vdom.props.children.props.children[1];
    const tabButtons = tabsContainer.props.children;
    const openAiTab = tabButtons[1];
    openAiTab.props.onClick();

    // 3. Edit OpenAI config
    vdom = harness.render();
    body = vdom.props.children.props.children[2];
    const openAiKeyInput = body.props.children[0].props.children[1];
    const openAiModelInput = body.props.children[1].props.children[1];
    const openAiBaseUrlInput = body.props.children[2].props.children[1];

    openAiKeyInput.props.onChange({ target: { value: 'openai-multi-tab-key' } });
    openAiModelInput.props.onChange({ target: { value: 'gpt-4o-latest' } });
    openAiBaseUrlInput.props.onChange({ target: { value: 'https://proxy.openai.corp/v1' } });

    // 4. Click "Simpan & Gunakan"
    vdom = harness.render();
    const footer = vdom.props.children.props.children[3];
    const saveBtn = footer.props.children[1];
    saveBtn.props.onClick();

    expect(onSaved).toHaveBeenCalledTimes(1);
    expect(onClose).toHaveBeenCalledTimes(1);

    // 5. Verify ALL tab changes were persisted via saveSettings
    const savedSettings = getSettings();
    expect(savedSettings.activeProviderId).toBe('openai');
    expect(savedSettings.providers.gemini.apiKey).toBe('gemini-multi-tab-key');
    expect(savedSettings.providers.openai.apiKey).toBe('openai-multi-tab-key');
    expect(savedSettings.providers.openai.selectedModel).toBe('gpt-4o-latest');
    expect(savedSettings.providers.openai.baseUrl).toBe('https://proxy.openai.corp/v1');
  });

  it('calls onClose when close button (X) is clicked', () => {
    const onClose = vi.fn();
    const onSaved = vi.fn();

    const harness = createHookHarness(SettingsModal, {
      isOpen: true,
      onClose,
      onSaved,
    });

    const vdom = harness.render();
    const headerDiv = vdom.props.children.props.children[0];
    const closeBtn = headerDiv.props.children[1];
    closeBtn.props.onClick();

    expect(onClose).toHaveBeenCalledTimes(1);
  });
});
