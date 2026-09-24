import React, { useState, useEffect, useRef } from 'react';
import { ProviderConfig, ProviderId, SoulConfig } from '../../types';
import {
  getSettings,
  saveSettings,
  setActiveProvider,
  getSearchSettings,
  setSearchSettings,
  SearchSettings,
} from '../../services/storage/settingsStorage';
import {
  getSoulConfig,
  saveSoulConfig,
  removeLearnedDirective,
  exportSoulToMarkdown,
  importSoulFromMarkdown,
  BrandVoicePreset,
  SOUL_PRESETS,
} from '../../services/storage/soulStorage';
import { testProviderConnection } from '../../services/llm/factory';
import { ThemeMode, getStoredThemeMode, setStoredThemeMode, applyTheme } from '../../utils/theme';
import { createBackup, restoreBackup } from '../../services/storage/backupManager';
import {
  X,
  CheckCircle2,
  AlertCircle,
  Loader2,
  Sparkles,
  ShieldCheck,
  Sun,
  Moon,
  Monitor,
  ChevronDown,
  Globe,
  BookOpen,
  Download,
  Upload,
  Database,
  Eye,
  EyeOff,
} from 'lucide-react';

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
  onSaved?: () => void;
  currentTheme?: ThemeMode;
  onThemeChange?: (mode: ThemeMode) => void;
  providers?: any;
  activeProvider?: string;
  onUpdateProviders?: (providers: any) => void;
  onSelectProvider?: (provider: any) => void;
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
  const [searchSettings, setLocalSearchSettings] = useState<SearchSettings>(() => getSearchSettings());
  const [soulConfig, setSoulConfig] = useState<SoulConfig>(() => getSoulConfig());
  const soulFileInputRef = useRef<HTMLInputElement>(null);

  // Backup & Restore states
  const [exportPassword, setExportPassword] = useState('');
  const [exportPlainJson, setExportPlainJson] = useState(false);
  const [showExportPassword, setShowExportPassword] = useState(false);
  const [isExporting, setIsExporting] = useState(false);

  const [restoreFileContent, setRestoreFileContent] = useState<string | null>(null);
  const [restoreFileName, setRestoreFileName] = useState('');
  const [isRestoreEncrypted, setIsRestoreEncrypted] = useState(false);
  const [restorePassword, setRestorePassword] = useState('');
  const [showRestorePassword, setShowRestorePassword] = useState(false);
  const [cleanRestore, setCleanRestore] = useState(false);
  const [isRestoring, setIsRestoring] = useState(false);
  const [backupStatus, setBackupStatus] = useState<{
    type: 'success' | 'error';
    message: string;
    details?: string;
  } | null>(null);
  const backupFileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (isOpen) {
      const currentSettings = getSettings();
      setSettings(currentSettings);
      setActiveTab(currentSettings.activeProviderId);
      setTestResult(null);
      setSelectedTheme(currentTheme || getStoredThemeMode());
      setLocalSearchSettings(getSearchSettings());
      setSoulConfig(getSoulConfig());
      setBackupStatus(null);
      setRestoreFileContent(null);
      setRestoreFileName('');
      setRestorePassword('');
      setIsRestoreEncrypted(false);
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

  const handleUpdateSearchProvider = (provider: 'duckduckgo' | 'tavily') => {
    const updated: SearchSettings = { ...searchSettings, searchProvider: provider };
    setLocalSearchSettings(updated);
    setSearchSettings(updated);
  };

  const handleUpdateTavilyKey = (key: string) => {
    const updated: SearchSettings = { ...searchSettings, tavilyApiKey: key };
    setLocalSearchSettings(updated);
    setSearchSettings(updated);
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

  const handleToggleSoul = (enabled: boolean) => {
    const updated = { ...soulConfig, enabled };
    setSoulConfig(updated);
    saveSoulConfig(updated);
  };

  const handleCorporateNameChange = (corporateName: string) => {
    const updated = { ...soulConfig, corporateName };
    setSoulConfig(updated);
    saveSoulConfig(updated);
  };

  const handleSelectPreset = (preset: BrandVoicePreset) => {
    const updatedMarkdown = preset === 'custom' ? soulConfig.rawSoulMarkdown : (SOUL_PRESETS[preset] || soulConfig.rawSoulMarkdown);
    const updated = {
      ...soulConfig,
      brandVoicePreset: preset,
      rawSoulMarkdown: updatedMarkdown,
    };
    setSoulConfig(updated);
    saveSoulConfig(updated);
  };

  const handleRawSoulChange = (text: string) => {
    const updated = {
      ...soulConfig,
      rawSoulMarkdown: text,
      brandVoicePreset: 'custom' as BrandVoicePreset,
    };
    setSoulConfig(updated);
    saveSoulConfig(updated);
  };

  const handleDeleteDirective = (id: string) => {
    removeLearnedDirective(id);
    const updatedDirectives = soulConfig.learnedDirectives.filter(d => d.id !== id);
    const updated = { ...soulConfig, learnedDirectives: updatedDirectives };
    setSoulConfig(updated);
  };

  const handleExportSoul = () => {
    try {
      const mdContent = exportSoulToMarkdown();
      if (typeof window !== 'undefined' && typeof document !== 'undefined') {
        const blob = new Blob([mdContent], { type: 'text/markdown;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = 'SOUL.md';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      }
    } catch (err) {
      console.error('Gagal mengekspor SOUL.md:', err);
    }
  };

  const handleImportSoul = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const reader = new FileReader();
      reader.onload = (event) => {
        const content = event.target?.result;
        if (typeof content === 'string') {
          const updated = importSoulFromMarkdown(content);
          setSoulConfig(updated);
        }
      };
      reader.readAsText(file);
    } catch (err) {
      console.error('Gagal mengimpor SOUL.md:', err);
    } finally {
      e.target.value = '';
    }
  };

  const handleDownloadBackup = async () => {
    try {
      setIsExporting(true);
      setBackupStatus(null);
      const password = exportPlainJson ? undefined : exportPassword.trim() || undefined;
      const backupJson = await createBackup({ password });

      if (typeof window !== 'undefined' && typeof document !== 'undefined') {
        const blob = new Blob([backupJson], { type: 'application/json;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        const dateStr = new Date().toISOString().split('T')[0];
        const ext = exportPlainJson ? 'json' : 'sam-backup';
        link.download = `sam-office-backup-${dateStr}.${ext}`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      }

      setBackupStatus({
        type: 'success',
        message: exportPlainJson
          ? 'Cadangan Plain JSON berhasil diunduh!'
          : 'Cadangan terenkripsi AES-256-GCM berhasil diunduh!',
      });
    } catch (err: any) {
      setBackupStatus({
        type: 'error',
        message: `Gagal membuat cadangan: ${err?.message || err}`,
      });
    } finally {
      setIsExporting(false);
    }
  };

  const handleSelectBackupFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setBackupStatus(null);
    setRestoreFileName(file.name);
    try {
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const content = event.target?.result as string;
          setRestoreFileContent(content);
          const parsed = JSON.parse(content);
          setIsRestoreEncrypted(Boolean(parsed.encrypted));
        } catch {
          setBackupStatus({
            type: 'error',
            message: 'Berkas tidak valid: format JSON tidak terbaca.',
          });
          setRestoreFileContent(null);
          setIsRestoreEncrypted(false);
        }
      };
      reader.onerror = () => {
        setBackupStatus({
          type: 'error',
          message: 'Gagal membaca berkas cadangan.',
        });
      };
      reader.readAsText(file);
    } catch (err: any) {
      setBackupStatus({
        type: 'error',
        message: `Gagal membuka berkas: ${err?.message || err}`,
      });
    } finally {
      e.target.value = '';
    }
  };

  const handleRestoreBackup = async () => {
    if (!restoreFileContent) {
      setBackupStatus({
        type: 'error',
        message: 'Pilih berkas cadangan terlebih dahulu.',
      });
      return;
    }

    try {
      setIsRestoring(true);
      setBackupStatus(null);

      const res = await restoreBackup(restoreFileContent, {
        password: isRestoreEncrypted ? restorePassword : undefined,
        cleanRestore,
      });

      if (res.success) {
        const currentSettings = getSettings();
        setSettings(currentSettings);
        setActiveTab(currentSettings.activeProviderId);
        setLocalSearchSettings(getSearchSettings());
        setSoulConfig(getSoulConfig());

        const summaryParts = [
          `${res.summary.providersCount} konfigurasi provider AI`,
          res.summary.soulRestored ? 'pedoman SOUL.md' : null,
          `${res.summary.customToolsCount} custom tools`,
          `${res.summary.snapshotsCount} snapshot lintas aplikasi`,
        ].filter(Boolean);

        setBackupStatus({
          type: 'success',
          message: 'Pemulihan data berhasil!',
          details: `Berhasil memulihkan: ${summaryParts.join(', ')}.`,
        });

        setRestoreFileContent(null);
        setRestoreFileName('');
        setRestorePassword('');
        setIsRestoreEncrypted(false);
      }
    } catch (err: any) {
      setBackupStatus({
        type: 'error',
        message: `Gagal memulihkan cadangan: ${err?.message || err}`,
      });
    } finally {
      setIsRestoring(false);
    }
  };

  const handleSaveAndApply = () => {
    const updatedSettings = {
      ...settings,
      activeProviderId: activeTab,
    };
    saveSettings(updatedSettings);
    setActiveProvider(activeTab);
    setSearchSettings(searchSettings);
    saveSoulConfig(soulConfig);
    onSaved?.();
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

        {/* Provider Menu Dropdown */}
        <div className="px-4 py-2.5 bg-gray-50/80 dark:bg-gray-900/50 border-b border-gray-200 dark:border-gray-700 flex items-center justify-between gap-2">
          <label htmlFor="byok-provider-dropdown" className="text-[11px] font-semibold text-gray-600 dark:text-gray-300 uppercase tracking-wider shrink-0 flex items-center gap-1">
            <span>Provider AI:</span>
          </label>
          <div className="relative flex-1 max-w-[240px]">
            <select
              id="byok-provider-dropdown"
              value={activeTab}
              aria-label="Pilih Provider BYOK"
              onChange={e => {
                setActiveTab(e.target.value as ProviderId);
                setTestResult(null);
              }}
              className="w-full text-xs font-semibold py-1.5 pl-3 pr-8 rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-800 text-gray-800 dark:text-gray-200 focus:outline-none focus:ring-2 focus:ring-blue-500 shadow-2xs appearance-none cursor-pointer truncate"
            >
              {(Object.keys(settings.providers) as ProviderId[]).map(pid => (
                <option key={pid} value={pid}>
                  {settings.providers[pid].name}
                </option>
              ))}
            </select>
            <ChevronDown className="w-3.5 h-3.5 text-gray-400 pointer-events-none absolute right-2.5 top-1/2 -translate-y-1/2" />
          </div>
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

          {/* Konfigurasi Pencarian Web (Live Web Search) */}
          <div className="p-2.5 bg-gray-50 dark:bg-gray-900/60 rounded-lg border border-gray-200 dark:border-gray-700/80 space-y-2">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-1.5">
                <Globe className="w-3.5 h-3.5 text-blue-500" />
                <span className="text-xs font-semibold text-gray-700 dark:text-gray-300">
                  Pencarian Web (Live Research)
                </span>
              </div>
              <span className="text-[10px] text-gray-400 capitalize">
                {searchSettings.searchProvider === 'duckduckgo' ? 'Zero-Config' : 'AI Search'}
              </span>
            </div>

            <div>
              <label className="block text-[11px] font-medium text-gray-600 dark:text-gray-400 mb-1">
                Penyedia Pencarian
              </label>
              <div className="grid grid-cols-2 gap-1.5">
                <button
                  type="button"
                  onClick={() => handleUpdateSearchProvider('duckduckgo')}
                  className={`py-1.5 px-2.5 rounded-md text-left flex flex-col transition-all border ${
                    searchSettings.searchProvider === 'duckduckgo'
                      ? 'bg-blue-50 dark:bg-blue-950/60 border-blue-500 text-blue-700 dark:text-blue-300 shadow-2xs'
                      : 'bg-white dark:bg-gray-800 border-gray-200 dark:border-gray-700 text-gray-600 dark:text-gray-400 hover:bg-gray-100 dark:hover:bg-gray-750'
                  }`}
                >
                  <span className="text-xs font-semibold">DuckDuckGo</span>
                  <span className="text-[10px] opacity-75">Zero-Config Gratis</span>
                </button>

                <button
                  type="button"
                  onClick={() => handleUpdateSearchProvider('tavily')}
                  className={`py-1.5 px-2.5 rounded-md text-left flex flex-col transition-all border ${
                    searchSettings.searchProvider === 'tavily'
                      ? 'bg-blue-50 dark:bg-blue-950/60 border-blue-500 text-blue-700 dark:text-blue-300 shadow-2xs'
                      : 'bg-white dark:bg-gray-800 border-gray-200 dark:border-gray-700 text-gray-600 dark:text-gray-400 hover:bg-gray-100 dark:hover:bg-gray-750'
                  }`}
                >
                  <span className="text-xs font-semibold">Tavily AI Search</span>
                  <span className="text-[10px] opacity-75">Optimasi Agen AI</span>
                </button>
              </div>
            </div>

            {(searchSettings.searchProvider === 'tavily' || Boolean(searchSettings.tavilyApiKey)) && (
              <div>
                <label className="block text-[11px] font-medium text-gray-600 dark:text-gray-400 mb-1">
                  Tavily API Key
                </label>
                <input
                  type="password"
                  value={searchSettings.tavilyApiKey || ''}
                  onChange={e => handleUpdateTavilyKey(e.target.value)}
                  placeholder="tvly-..."
                  className="w-full text-xs px-3 py-1.5 rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-900 text-gray-900 dark:text-gray-100 focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
                <p className="text-[10px] text-gray-400 mt-1">
                  Dapatkan API Key di tavily.com. Jika kosong atau gagal, sistem otomatis menggunakan DuckDuckGo.
                </p>
              </div>
            )}
          </div>

          {/* Brand Voice & Pedoman Korporat (SOUL.md) */}
          <div className="p-2.5 bg-gray-50 dark:bg-gray-900/60 rounded-lg border border-gray-200 dark:border-gray-700/80 space-y-2.5">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-1.5">
                <BookOpen className="w-3.5 h-3.5 text-blue-600 dark:text-blue-400" />
                <span className="text-xs font-semibold text-gray-700 dark:text-gray-300">
                  Brand Voice &amp; Pedoman Korporat (SOUL.md)
                </span>
              </div>
              <label className="flex items-center gap-1.5 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={soulConfig.enabled}
                  onChange={e => handleToggleSoul(e.target.checked)}
                  className="w-3.5 h-3.5 text-blue-600 rounded border-gray-300 dark:border-gray-600 focus:ring-blue-500"
                />
                <span className="text-[10px] font-medium text-gray-600 dark:text-gray-400">
                  Aktifkan Corporate Directives (SOUL.md)
                </span>
              </label>
            </div>

            <div>
              <label className="block text-[11px] font-medium text-gray-600 dark:text-gray-400 mb-1">
                Nama Entitas / Perusahaan
              </label>
              <input
                type="text"
                value={soulConfig.corporateName}
                onChange={e => handleCorporateNameChange(e.target.value)}
                placeholder="Nama Entitas / Perusahaan (contoh: PT Maju Bersama)"
                className="w-full text-xs px-3 py-1.5 rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-900 text-gray-900 dark:text-gray-100 focus:outline-none focus:ring-2 focus:ring-blue-500"
              />
            </div>

            <div>
              <label className="block text-[11px] font-medium text-gray-600 dark:text-gray-400 mb-1">
                Preset Karakter Brand Voice
              </label>
              <div className="grid grid-cols-2 gap-1.5">
                {[
                  { id: 'formal_executive', label: 'Formal Eksekutif', desc: 'Baku, data-driven, wibawa' },
                  { id: 'modern_professional', label: 'Modern & Ringkas', desc: 'Agile, lugas, aktif' },
                  { id: 'financial_compliance', label: 'Finansial & Kepatuhan', desc: 'Audit, mitigasi risiko' },
                  { id: 'custom', label: 'Kustom', desc: 'Aturan disesuaikan sendiri' },
                ].map(preset => (
                  <button
                    key={preset.id}
                    type="button"
                    onClick={() => handleSelectPreset(preset.id as BrandVoicePreset)}
                    className={`py-1.5 px-2.5 rounded-md text-left flex flex-col transition-all border ${
                      soulConfig.brandVoicePreset === preset.id
                        ? 'bg-blue-50 dark:bg-blue-950/60 border-blue-500 text-blue-700 dark:text-blue-300 shadow-2xs'
                        : 'bg-white dark:bg-gray-800 border-gray-200 dark:border-gray-700 text-gray-600 dark:text-gray-400 hover:bg-gray-100 dark:hover:bg-gray-750'
                    }`}
                  >
                    <span className="text-xs font-semibold">{preset.label}</span>
                    <span className="text-[10px] opacity-75">{preset.desc}</span>
                  </button>
                ))}
              </div>
            </div>

            <div>
              <label className="block text-[11px] font-medium text-gray-600 dark:text-gray-400 mb-1">
                Pedoman SOUL Markdown (SOUL.md)
              </label>
              <textarea
                rows={5}
                value={soulConfig.rawSoulMarkdown}
                onChange={e => handleRawSoulChange(e.target.value)}
                placeholder="# SOUL & CORPORATE BRAND DIRECTIVES..."
                className="w-full text-[11px] font-mono p-2 rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-900 text-gray-900 dark:text-gray-100 focus:outline-none focus:ring-2 focus:ring-blue-500 leading-relaxed resize-y"
              />
            </div>

            {/* Direktif yang Dipelajari Mandiri */}
            <div className="space-y-1.5">
              <div className="flex items-center justify-between">
                <span className="text-[11px] font-medium text-gray-600 dark:text-gray-400">
                  Direktif yang Dipelajari Mandiri ({soulConfig.learnedDirectives?.length || 0})
                </span>
              </div>
              {soulConfig.learnedDirectives && soulConfig.learnedDirectives.length > 0 ? (
                <div className="flex flex-wrap gap-1 max-h-32 overflow-y-auto p-1 bg-white/60 dark:bg-gray-855/60 rounded border border-gray-200/80 dark:border-gray-700/60">
                  {soulConfig.learnedDirectives.map(d => (
                    <div
                      key={d.id}
                      className="flex items-center gap-1 px-1.5 py-0.5 rounded text-[10px] bg-blue-50 dark:bg-blue-950/40 text-blue-800 dark:text-blue-300 border border-blue-200 dark:border-blue-900/50"
                    >
                      <span className="font-semibold uppercase text-[9px] opacity-80">[{d.category}]</span>
                      <span className="max-w-[200px] truncate" title={d.rule}>{d.rule}</span>
                      <button
                        type="button"
                        onClick={() => handleDeleteDirective(d.id)}
                        className="text-gray-400 hover:text-red-500 font-bold ml-0.5 leading-none"
                        title="Hapus direktif ini"
                      >
                        ✕
                      </button>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="text-[10px] text-gray-400 italic">
                  Belum ada aturan yang dipelajari mandiri. Sam akan otomatis mencatat preferensi gaya Anda.
                </p>
              )}
            </div>

            {/* Ekspor & Impor SOUL.md */}
            <div className="flex items-center gap-2 pt-1 border-t border-gray-200/60 dark:border-gray-700/60">
              <input
                type="file"
                ref={soulFileInputRef}
                accept=".md,text/markdown"
                className="hidden"
                onChange={handleImportSoul}
              />
              <button
                type="button"
                onClick={handleExportSoul}
                className="flex-1 flex items-center justify-center gap-1.5 py-1.5 px-2 rounded-md text-[11px] font-medium border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-800 text-gray-700 dark:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-750 transition shadow-2xs"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Ekspor SOUL.md</span>
              </button>
              <button
                type="button"
                onClick={() => soulFileInputRef.current?.click()}
                className="flex-1 flex items-center justify-center gap-1.5 py-1.5 px-2 rounded-md text-[11px] font-medium border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-800 text-gray-700 dark:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-750 transition shadow-2xs"
              >
                <Upload className="w-3.5 h-3.5" />
                <span>Impor SOUL.md</span>
              </button>
            </div>
          </div>

          {/* Cadangan & Pemulihan Data (Backup & Restore) */}
          <div className="p-2.5 bg-gray-50 dark:bg-gray-900/60 rounded-lg border border-gray-200 dark:border-gray-700/80 space-y-2.5">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-1.5">
                <Database className="w-3.5 h-3.5 text-blue-600 dark:text-blue-400" />
                <span className="text-xs font-semibold text-gray-700 dark:text-gray-300">
                  Cadangan &amp; Pemulihan Data (Backup &amp; Restore)
                </span>
              </div>
              <span className="text-[10px] text-gray-400">AES-256-GCM</span>
            </div>

            {/* Status Message / Banner */}
            {backupStatus && (
              <div
                className={`p-2 rounded-lg text-xs flex flex-col gap-0.5 ${
                  backupStatus.type === 'success'
                    ? 'bg-green-50 dark:bg-green-950/40 text-green-700 dark:text-green-300 border border-green-200 dark:border-green-800'
                    : 'bg-red-50 dark:bg-red-950/40 text-red-700 dark:text-red-300 border border-red-200 dark:border-red-800'
                }`}
              >
                <div className="flex items-center gap-1.5 font-medium">
                  {backupStatus.type === 'success' ? (
                    <CheckCircle2 className="w-3.5 h-3.5 shrink-0 text-green-600" />
                  ) : (
                    <AlertCircle className="w-3.5 h-3.5 shrink-0 text-red-600" />
                  )}
                  <span>{backupStatus.message}</span>
                </div>
                {backupStatus.details && (
                  <p className="text-[11px] opacity-90 pl-5">{backupStatus.details}</p>
                )}
              </div>
            )}

            {/* Export Card */}
            <div className="p-2 bg-white dark:bg-gray-800/80 rounded-md border border-gray-200 dark:border-gray-700 space-y-2">
              <div className="flex items-center justify-between">
                <span className="text-[11px] font-semibold text-gray-700 dark:text-gray-300">
                  Ekspor Cadangan
                </span>
                <label className="flex items-center gap-1.5 cursor-pointer select-none">
                  <input
                    type="checkbox"
                    checked={exportPlainJson}
                    onChange={e => setExportPlainJson(e.target.checked)}
                    className="w-3.5 h-3.5 text-blue-600 rounded border-gray-300 dark:border-gray-600 focus:ring-blue-500"
                  />
                  <span className="text-[10px] text-gray-600 dark:text-gray-400">
                    Ekspor tanpa enkripsi (Plain JSON)
                  </span>
                </label>
              </div>

              {!exportPlainJson && (
                <div className="relative">
                  <input
                    type={showExportPassword ? 'text' : 'password'}
                    value={exportPassword}
                    onChange={e => setExportPassword(e.target.value)}
                    placeholder="Kata sandi enkripsi cadangan..."
                    className="w-full text-xs px-3 py-1.5 pr-8 rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-900 text-gray-900 dark:text-gray-100 focus:outline-none focus:ring-2 focus:ring-blue-500"
                  />
                  <button
                    type="button"
                    onClick={() => setShowExportPassword(!showExportPassword)}
                    aria-label={showExportPassword ? 'Sembunyikan Kata Sandi' : 'Tampilkan Kata Sandi'}
                    className="absolute right-2.5 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600 dark:hover:text-gray-300"
                  >
                    {showExportPassword ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
                  </button>
                </div>
              )}

              <button
                type="button"
                onClick={handleDownloadBackup}
                disabled={isExporting}
                className="w-full flex items-center justify-center gap-1.5 py-1.5 px-3 rounded-md text-[11px] font-medium bg-blue-50 hover:bg-blue-100 dark:bg-blue-950/60 dark:hover:bg-blue-900/80 text-blue-700 dark:text-blue-300 border border-blue-200 dark:border-blue-800 transition shadow-2xs"
              >
                {isExporting ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : null}
                <span>📥 Unduh Cadangan (.sam-backup)</span>
              </button>
            </div>

            {/* Restore Card */}
            <div className="p-2 bg-white dark:bg-gray-800/80 rounded-md border border-gray-200 dark:border-gray-700 space-y-2">
              <span className="text-[11px] font-semibold text-gray-700 dark:text-gray-300 block">
                Pemulihan Cadangan
              </span>

              <input
                type="file"
                ref={backupFileInputRef}
                accept=".sam-backup,.json"
                className="hidden"
                onChange={handleSelectBackupFile}
              />

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => backupFileInputRef.current?.click()}
                  className="flex-1 flex items-center justify-center gap-1.5 py-1.5 px-3 rounded-md text-[11px] font-medium border border-gray-300 dark:border-gray-600 bg-gray-50 dark:bg-gray-700 text-gray-700 dark:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-600 transition shadow-2xs"
                >
                  <Upload className="w-3.5 h-3.5" />
                  <span>📤 Pilih Berkas Cadangan</span>
                </button>
              </div>

              {restoreFileName && (
                <div className="text-[10px] text-gray-500 dark:text-gray-400 bg-gray-50 dark:bg-gray-900/50 px-2 py-1 rounded border border-gray-200 dark:border-gray-700 truncate">
                  Berkas dipilih: <span className="font-semibold text-gray-700 dark:text-gray-300">{restoreFileName}</span>
                  {isRestoreEncrypted && <span className="ml-1 text-amber-600 dark:text-amber-400 font-medium">(Terenkripsi)</span>}
                </div>
              )}

              {isRestoreEncrypted && (
                <div className="relative">
                  <input
                    type={showRestorePassword ? 'text' : 'password'}
                    value={restorePassword}
                    onChange={e => setRestorePassword(e.target.value)}
                    placeholder="Masukkan kata sandi pembuka cadangan..."
                    className="w-full text-xs px-3 py-1.5 pr-8 rounded-lg border border-amber-300 dark:border-amber-600 bg-white dark:bg-gray-900 text-gray-900 dark:text-gray-100 focus:outline-none focus:ring-2 focus:ring-amber-500"
                  />
                  <button
                    type="button"
                    onClick={() => setShowRestorePassword(!showRestorePassword)}
                    aria-label={showRestorePassword ? 'Sembunyikan Kata Sandi' : 'Tampilkan Kata Sandi'}
                    className="absolute right-2.5 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600 dark:hover:text-gray-300"
                  >
                    {showRestorePassword ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
                  </button>
                </div>
              )}

              <label className="flex items-center gap-1.5 cursor-pointer select-none pt-0.5">
                <input
                  type="checkbox"
                  checked={cleanRestore}
                  onChange={e => setCleanRestore(e.target.checked)}
                  className="w-3.5 h-3.5 text-blue-600 rounded border-gray-300 dark:border-gray-600 focus:ring-blue-500"
                />
                <span className="text-[10px] text-gray-600 dark:text-gray-400">
                  Bersihkan data lokal saat ini sebelum restore (Clean Restore)
                </span>
              </label>

              <button
                type="button"
                onClick={handleRestoreBackup}
                disabled={!restoreFileContent || isRestoring}
                className={`w-full flex items-center justify-center gap-1.5 py-1.5 px-3 rounded-md text-[11px] font-medium transition ${
                  !restoreFileContent || isRestoring
                    ? 'opacity-50 cursor-not-allowed bg-gray-100 dark:bg-gray-800 text-gray-400 border border-gray-200 dark:border-gray-700'
                    : 'bg-emerald-600 hover:bg-emerald-700 text-white shadow-2xs'
                }`}
              >
                {isRestoring ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : null}
                <span>Pulihkan Data (Restore)</span>
              </button>
            </div>
          </div>

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
