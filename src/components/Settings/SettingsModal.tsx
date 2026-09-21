import React, { useState, useEffect } from 'react';
import { ProviderConfig, ProviderId } from '../../types';
import { getSettings, saveSettings, setActiveProvider } from '../../services/storage/settingsStorage';
import { testProviderConnection } from '../../services/llm/factory';
import { ThemeMode, getStoredThemeMode, setStoredThemeMode, applyTheme } from '../../utils/theme';
import { X, CheckCircle2, AlertCircle, Loader2, Sparkles, ShieldCheck, Sun, Moon, Monitor } from 'lucide-react';

export const PROVIDER_LABELS: Record<ProviderId, string> = {
  gemini: 'Gemini',
  openai: 'OpenAI',
  claude: 'Claude',
  glm: 'GLM',
  openrouter: 'OpenRouter',
  ollama: 'Ollama',
  'openai-compatible': 'Compatible',
};

export const OPENAI_COMPATIBLE_PRESETS = [
  { name: 'Groq', baseUrl: 'https://api.groq.com/openai/v1', model: 'llama-3.3-70b-versatile' },
  { name: 'DeepSeek', baseUrl: 'https://api.deepseek.com/v1', model: 'deepseek-chat' },
  { name: 'LM Studio (Local)', baseUrl: 'http://localhost:1234/v1', model: 'qwen2.5-coder-7b-instruct' },
  { name: 'vLLM / LocalAI', baseUrl: 'http://localhost:8000/v1', model: 'default' },
  { name: 'Mistral', baseUrl: 'https://api.mistral.ai/v1', model: 'mistral-large-latest' },
  { name: 'Together AI', baseUrl: 'https://api.together.xyz/v1', model: 'meta-llama/Llama-3.3-70B-Instruct-Turbo' },
];

interface SettingsModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSaved: () => void;
  currentTheme?: ThemeMode;
  onThemeChange?: (mode: ThemeMode) => void;
}

export const SettingsModal: React.FC<SettingsModalProps> = ({
  isOpen,
  onClose,
  onSaved,
  currentTheme,
  onThemeChange,
}) => {
  const [settings, setSettings] = useState(getSettings());
  const [activeTab, setActiveTab] = useState<ProviderId>(settings.activeProviderId);
  const [testing, setTesting] = useState(false);
  const [testResult, setTestResult] = useState<{ success: boolean; message: string } | null>(null);
  const [selectedTheme, setSelectedTheme] = useState<ThemeMode>(() => currentTheme || getStoredThemeMode());

  useEffect(() => {
    if (isOpen) {
      const currentSettings = getSettings();
      setSettings(currentSettings);
      setActiveTab(currentSettings.activeProviderId);
      setTestResult(null);
      setSelectedTheme(currentTheme || getStoredThemeMode());
    }
  }, [isOpen, currentTheme]);

  if (!isOpen) return null;

  const currentConfig: ProviderConfig = settings.providers[activeTab] || {
    id: activeTab,
    name: activeTab,
    apiKey: '',
    selectedModel: '',
    enabled: true,
  };

  const handleUpdateKey = (key: string) => {
    setSettings(prev => {
      const existing = prev.providers[activeTab] || currentConfig;
      return {
        ...prev,
        providers: {
          ...prev.providers,
          [activeTab]: { ...existing, apiKey: key },
        },
      };
    });
  };

  const handleUpdateModel = (model: string) => {
    setSettings(prev => {
      const existing = prev.providers[activeTab] || currentConfig;
      return {
        ...prev,
        providers: {
          ...prev.providers,
          [activeTab]: { ...existing, selectedModel: model },
        },
      };
    });
  };

  const handleUpdateBaseUrl = (baseUrl: string) => {
    setSettings(prev => {
      const existing = prev.providers[activeTab] || currentConfig;
      return {
        ...prev,
        providers: {
          ...prev.providers,
          [activeTab]: { ...existing, baseUrl },
        },
      };
    });
  };

  const handleApplyPreset = (preset: { baseUrl: string; model: string }) => {
    setSettings(prev => {
      const existing = prev.providers[activeTab] || currentConfig;
      return {
        ...prev,
        providers: {
          ...prev.providers,
          [activeTab]: {
            ...existing,
            baseUrl: preset.baseUrl,
            selectedModel: preset.model,
          },
        },
      };
    });
    setTestResult(null);
  };

  const handleTest = async () => {
    setTesting(true);
    setTestResult(null);
    try {
      const configToTest = settings.providers[activeTab] || currentConfig;
      const res = await testProviderConnection(configToTest);
      setTestResult(res);
    } catch (err) {
      setTestResult({
        success: false,
        message: err instanceof Error ? err.message : 'Koneksi gagal',
      });
    } finally {
      setTesting(false);
    }
  };

  const handleSelectTheme = (mode: ThemeMode) => {
    setSelectedTheme(mode);
    setStoredThemeMode(mode);
    applyTheme(mode);
    onThemeChange?.(mode);
  };

  const handleSaveAndApply = () => {
    const updatedSettings = {
      ...settings,
      activeProviderId: activeTab,
    };
    saveSettings(updatedSettings);
    setActiveProvider(activeTab);
    onSaved();
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-3">
      <div className="bg-white dark:bg-gray-800 rounded-xl max-w-md w-full shadow-xl border border-gray-200 dark:border-gray-700 overflow-hidden flex flex-col max-h-[90vh]">
        <div className="flex items-center justify-between p-4 border-b border-gray-100 dark:border-gray-700">
          <div>
            <h2 className="text-base font-semibold text-gray-900 dark:text-gray-100">Pengaturan BYOK AI</h2>
            <p className="text-[11px] text-gray-500 dark:text-gray-400">Bring Your Own Key — 100% Client-Side Privacy</p>
          </div>
          <button onClick={onClose} aria-label="Tutup" className="text-gray-400 hover:text-gray-600">
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="flex border-b border-gray-200 dark:border-gray-700 overflow-x-auto text-xs">
          {(Object.keys(settings.providers) as ProviderId[]).map(pid => (
            <button
              key={pid}
              onClick={() => {
                setActiveTab(pid);
                setTestResult(null);
              }}
              className={`px-3 py-2 whitespace-nowrap font-medium border-b-2 transition ${
                activeTab === pid
                  ? 'border-blue-600 text-blue-600 dark:text-blue-400'
                  : 'border-transparent text-gray-500 hover:text-gray-800 dark:hover:text-gray-200'
              }`}
            >
              {settings.providers[pid].name.split(' ')[0]}
            </button>
          ))}
        </div>

        <div className="p-4 space-y-3 overflow-y-auto flex-1">
          <div>
            <label className="block text-xs font-medium text-gray-700 dark:text-gray-300 mb-1">
              API Key ({currentConfig.name})
            </label>
            <input
              type="password"
              value={currentConfig.apiKey}
              onChange={e => handleUpdateKey(e.target.value)}
              placeholder="Masukkan API Key Anda..."
              className="w-full text-xs px-3 py-2 rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-900 text-gray-900 dark:text-gray-100 focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-gray-700 dark:text-gray-300 mb-1">
              Nama Model
            </label>
            <input
              type="text"
              value={currentConfig.selectedModel}
              onChange={e => handleUpdateModel(e.target.value)}
              placeholder="Contoh: gpt-4o, llama-3.3-70b-versatile, deepseek-chat..."
              className="w-full text-xs px-3 py-2 rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-900 text-gray-900 dark:text-gray-100 focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-gray-700 dark:text-gray-300 mb-1">
              Custom Base URL / Proxy (Opsional)
            </label>
            <input
              type="text"
              value={currentConfig.baseUrl || ''}
              onChange={e => handleUpdateBaseUrl(e.target.value)}
              placeholder="https://..."
              className="w-full text-xs px-3 py-2 rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-900 text-gray-900 dark:text-gray-100 focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
            {activeTab === 'openai-compatible' && (
              <div className="mt-2.5 p-2 bg-blue-50 dark:bg-blue-950/40 rounded-lg border border-blue-100 dark:border-blue-900/50 space-y-1.5">
                <div className="flex items-center gap-1.5 text-[11px] font-semibold text-blue-800 dark:text-blue-300">
                  <Sparkles className="w-3.5 h-3.5" />
                  <span>Preset Cepat (Base URL &amp; Model):</span>
                </div>
                <div className="flex flex-wrap gap-1">
                  {OPENAI_COMPATIBLE_PRESETS.map(p => (
                    <button
                      key={p.name}
                      type="button"
                      onClick={() => handleApplyPreset(p)}
                      className="text-[10px] px-2 py-0.5 rounded bg-white dark:bg-gray-800 border border-blue-200 dark:border-blue-800 text-blue-700 dark:text-blue-300 hover:bg-blue-100 dark:hover:bg-blue-900/50 transition font-medium"
                    >
                      {p.name}
                    </button>
                  ))}
                </div>
              </div>
            )}
          </div>

          {testResult && (
            <div
              className={`p-2.5 rounded-lg text-xs flex items-start gap-2 ${
                testResult.success
                  ? 'bg-green-50 dark:bg-green-950/40 text-green-700 dark:text-green-300'
                  : 'bg-red-50 dark:bg-red-950/40 text-red-700 dark:text-red-300'
              }`}
            >
              {testResult.success ? (
                <CheckCircle2 className="w-4 h-4 shrink-0 text-green-600 mt-0.5" />
              ) : (
                <AlertCircle className="w-4 h-4 shrink-0 text-red-600 mt-0.5" />
              )}
              <span className="break-all">{testResult.message}</span>
            </div>
          )}

          {/* Pilihan Tema Tampilan */}
          <div className="p-2.5 bg-gray-50 dark:bg-gray-900/60 rounded-lg border border-gray-200 dark:border-gray-700/80 space-y-1.5">
            <div className="flex items-center justify-between">
              <span className="text-xs font-semibold text-gray-700 dark:text-gray-300">
                Tema Antarmuka
              </span>
              <span className="text-[10px] text-gray-400 capitalize">
                {selectedTheme === 'auto' ? 'Otomatis' : selectedTheme === 'dark' ? 'Mode Gelap' : 'Mode Terang'}
              </span>
            </div>
            <div className="grid grid-cols-3 gap-1.5">
              <button
                type="button"
                onClick={() => handleSelectTheme('auto')}
                className={`flex items-center justify-center gap-1.5 py-1.5 px-2 rounded-md text-[11px] font-medium border transition-all ${
                  selectedTheme === 'auto'
                    ? 'bg-blue-50 dark:bg-blue-950/60 border-blue-500 text-blue-700 dark:text-blue-300 shadow-2xs'
                    : 'bg-white dark:bg-gray-800 border-gray-200 dark:border-gray-700 text-gray-600 dark:text-gray-400 hover:bg-gray-100 dark:hover:bg-gray-750'
                }`}
              >
                <Monitor className="w-3.5 h-3.5" />
                <span>Otomatis</span>
              </button>
              <button
                type="button"
                onClick={() => handleSelectTheme('light')}
                className={`flex items-center justify-center gap-1.5 py-1.5 px-2 rounded-md text-[11px] font-medium border transition-all ${
                  selectedTheme === 'light'
                    ? 'bg-amber-50 border-amber-500 text-amber-700 shadow-2xs'
                    : 'bg-white dark:bg-gray-800 border-gray-200 dark:border-gray-700 text-gray-600 dark:text-gray-400 hover:bg-gray-100 dark:hover:bg-gray-750'
                }`}
              >
                <Sun className="w-3.5 h-3.5 text-amber-500" />
                <span>Terang</span>
              </button>
              <button
                type="button"
                onClick={() => handleSelectTheme('dark')}
                className={`flex items-center justify-center gap-1.5 py-1.5 px-2 rounded-md text-[11px] font-medium border transition-all ${
                  selectedTheme === 'dark'
                    ? 'bg-indigo-50 dark:bg-indigo-950/60 border-indigo-500 text-indigo-700 dark:text-indigo-300 shadow-2xs'
                    : 'bg-white dark:bg-gray-800 border-gray-200 dark:border-gray-700 text-gray-600 dark:text-gray-400 hover:bg-gray-100 dark:hover:bg-gray-750'
                }`}
              >
                <Moon className="w-3.5 h-3.5 text-indigo-400" />
                <span>Gelap</span>
              </button>
            </div>
          </div>

          <div className="p-2 bg-gray-50 dark:bg-gray-900/40 rounded border border-gray-200 dark:border-gray-700/60 flex items-start gap-2 text-[11px] text-gray-500 dark:text-gray-400">
            <ShieldCheck className="w-4 h-4 shrink-0 text-emerald-600 mt-0.5" />
            <span>Kunci API Anda disimpan eksklusif di peramban lokal perangkat ini (zero-backend).</span>
          </div>
        </div>

        <div className="p-3 bg-gray-50 dark:bg-gray-900/50 border-t border-gray-200 dark:border-gray-700 flex justify-between gap-2">
          <button
            onClick={handleTest}
            disabled={testing}
            className="px-3.5 py-1.5 text-xs font-medium rounded-lg border border-gray-300 dark:border-gray-600 hover:bg-gray-100 dark:hover:bg-gray-800 text-gray-700 dark:text-gray-300 flex items-center gap-1.5"
          >
            {testing && <Loader2 className="w-3.5 h-3.5 animate-spin" />}
            Test Koneksi
          </button>
          <button
            onClick={handleSaveAndApply}
            className="px-4 py-1.5 text-xs font-medium rounded-lg bg-blue-600 hover:bg-blue-700 text-white shadow-sm"
          >
            Simpan & Gunakan
          </button>
        </div>
      </div>
    </div>
  );
};
