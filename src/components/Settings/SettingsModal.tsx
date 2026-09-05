import React, { useState, useEffect } from 'react';
import { ProviderConfig, ProviderId } from '../../types';
import { getSettings, saveProviderConfig, setActiveProvider } from '../../services/storage/settingsStorage';
import { testProviderConnection } from '../../services/llm/factory';
import { X, CheckCircle2, AlertCircle, Loader2 } from 'lucide-react';

interface SettingsModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSaved: () => void;
}

export const SettingsModal: React.FC<SettingsModalProps> = ({ isOpen, onClose, onSaved }) => {
  const [settings, setSettings] = useState(getSettings());
  const [activeTab, setActiveTab] = useState<ProviderId>(settings.activeProviderId);
  const [testing, setTesting] = useState(false);
  const [testResult, setTestResult] = useState<{ success: boolean; message: string } | null>(null);

  useEffect(() => {
    if (isOpen) {
      const currentSettings = getSettings();
      setSettings(currentSettings);
      setActiveTab(currentSettings.activeProviderId);
      setTestResult(null);
    }
  }, [isOpen]);

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

  const handleSaveAndApply = () => {
    const configToSave = settings.providers[activeTab] || currentConfig;
    saveProviderConfig(configToSave);
    setActiveProvider(activeTab);
    onSaved();
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-3">
      <div className="bg-white dark:bg-gray-800 rounded-xl max-w-sm w-full shadow-xl border border-gray-200 dark:border-gray-700 overflow-hidden flex flex-col max-h-[90vh]">
        <div className="flex items-center justify-between p-4 border-b border-gray-100 dark:border-gray-700">
          <h2 className="text-base font-semibold text-gray-900 dark:text-gray-100">Pengaturan BYOK AI</h2>
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
              className="w-full text-xs px-3 py-2 rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-900 text-gray-900 dark:text-gray-100 focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
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
        </div>

        <div className="p-3 bg-gray-50 dark:bg-gray-900/50 border-t border-gray-200 dark:border-gray-700 flex justify-between gap-2">
          <button
            onClick={handleTest}
            disabled={testing}
            className="px-3 py-1.5 text-xs font-medium rounded-lg border border-gray-300 dark:border-gray-600 hover:bg-gray-100 dark:hover:bg-gray-800 text-gray-700 dark:text-gray-300 flex items-center gap-1.5"
          >
            {testing && <Loader2 className="w-3.5 h-3.5 animate-spin" />}
            Test Koneksi
          </button>
          <button
            onClick={handleSaveAndApply}
            className="px-3 py-1.5 text-xs font-medium rounded-lg bg-blue-600 hover:bg-blue-700 text-white shadow-sm"
          >
            Simpan & Gunakan
          </button>
        </div>
      </div>
    </div>
  );
};
