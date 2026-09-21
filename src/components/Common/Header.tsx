import React, { useState, useRef, useEffect } from 'react';
import { HostType, ProviderId, ExecutionMode } from '../../types';
import { ThemeMode } from '../../utils/theme';
import {
  Settings,
  Zap,
  Play,
  FileSpreadsheet,
  FileText,
  Presentation,
  Sun,
  Moon,
  Monitor,
  ChevronDown,
  Check,
} from 'lucide-react';

export const AVAILABLE_PROVIDERS: Array<{ id: ProviderId; label: string; badge: string }> = [
  { id: 'gemini', label: 'Google Gemini', badge: 'Gemini 2.0' },
  { id: 'openai', label: 'OpenAI (ChatGPT)', badge: 'GPT-4o' },
  { id: 'openai-compatible', label: 'Compatible (API / Local)', badge: 'Groq / LM Studio / vLLM' },
  { id: 'claude', label: 'Anthropic Claude', badge: 'Claude 3.5' },
  { id: 'glm', label: 'GLM Coding (Zhipu)', badge: 'GLM-4' },
  { id: 'openrouter', label: 'OpenRouter', badge: 'Multi-model' },
  { id: 'ollama', label: 'Local Ollama', badge: 'Offline' },
];

export interface HeaderProps {
  host: HostType;
  activeProviderId: ProviderId;
  executionMode: ExecutionMode;
  onToggleMode: () => void;
  onOpenSettings: () => void;
  themeMode?: ThemeMode;
  onToggleTheme?: () => void;
  onSelectProvider?: (providerId: ProviderId) => void;
}

export const Header: React.FC<HeaderProps> = ({
  host,
  activeProviderId,
  executionMode,
  onToggleMode,
  onOpenSettings,
  themeMode = 'auto',
  onToggleTheme,
  onSelectProvider,
}) => {
  let isProviderOpen = false;
  let setIsProviderOpen: React.Dispatch<React.SetStateAction<boolean>> = () => {};
  let dropdownRef: React.RefObject<HTMLDivElement> = { current: null };

  const hasDispatcher = Boolean(
    (React as any)?.__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED?.ReactCurrentDispatcher?.current
  );

  if (hasDispatcher) {
    // eslint-disable-next-line react-hooks/rules-of-hooks
    const [open, setOpen] = useState(false);
    isProviderOpen = open;
    setIsProviderOpen = setOpen;
    // eslint-disable-next-line react-hooks/rules-of-hooks
    dropdownRef = useRef<HTMLDivElement>(null);

    // eslint-disable-next-line react-hooks/rules-of-hooks
    useEffect(() => {
      if (!isProviderOpen || typeof document === 'undefined') return;
      const handleClickOutside = (e: MouseEvent) => {
        if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) {
          setIsProviderOpen(false);
        }
      };
      const handleKeyDown = (e: KeyboardEvent) => {
        if (e.key === 'Escape') {
          setIsProviderOpen(false);
        }
      };
      document.addEventListener('mousedown', handleClickOutside);
      document.addEventListener('keydown', handleKeyDown);
      return () => {
        document.removeEventListener('mousedown', handleClickOutside);
        document.removeEventListener('keydown', handleKeyDown);
      };
    }, [isProviderOpen]);
  }
  const getHostIcon = () => {
    switch (host) {
      case 'Excel':
        return <FileSpreadsheet className="w-4 h-4 text-office-excel" />;
      case 'Word':
        return <FileText className="w-4 h-4 text-office-word" />;
      case 'PowerPoint':
        return <Presentation className="w-4 h-4 text-office-ppt" />;
      default:
        return <Zap className="w-4 h-4 text-office-agent" />;
    }
  };

  const getHostBadgeStyle = () => {
    switch (host) {
      case 'Excel':
        return 'bg-emerald-100 dark:bg-emerald-950/60 text-emerald-800 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-800/60';
      case 'Word':
        return 'bg-blue-100 dark:bg-blue-950/60 text-blue-800 dark:text-blue-300 border border-blue-200 dark:border-blue-800/60';
      case 'PowerPoint':
        return 'bg-orange-100 dark:bg-orange-950/60 text-orange-800 dark:text-orange-300 border border-orange-200 dark:border-orange-800/60';
      default:
        return 'bg-blue-100 dark:bg-blue-900/60 text-blue-700 dark:text-blue-300';
    }
  };

  return (
    <header className="flex items-center justify-between px-3 py-1.5 bg-white/95 dark:bg-gray-800/95 backdrop-blur-xs border-b border-gray-200/80 dark:border-gray-700/80 shadow-2xs transition-colors shrink-0">
      <div className="flex items-center gap-2 min-w-0">
        <div className="p-1 rounded-md bg-gray-50 dark:bg-gray-700/60 border border-gray-200/60 dark:border-gray-700 shrink-0">
          {getHostIcon()}
        </div>
        <div className="flex items-center gap-1.5 min-w-0">
          <h1 className="text-xs font-bold text-gray-900 dark:text-gray-100 flex items-center gap-1 shrink-0">
            Sam Office
            <span className={`text-[9px] px-1.5 py-0.5 rounded-full font-semibold tracking-tight ${getHostBadgeStyle()}`}>
              {host}
            </span>
          </h1>
          <span className="text-gray-300 dark:text-gray-600 text-xs select-none">·</span>

          {/* Provider Dropdown Trigger & Menu */}
          <div className="relative" ref={dropdownRef}>
            <button
              type="button"
              onClick={() => setIsProviderOpen(prev => !prev)}
              aria-haspopup="true"
              aria-expanded={isProviderOpen}
              aria-label="Pilih Provider AI"
              title="Ganti Provider AI"
              className="flex items-center gap-1 text-[10px] text-gray-600 dark:text-gray-300 hover:text-gray-900 dark:hover:text-gray-100 font-medium capitalize px-1.5 py-0.5 rounded-md hover:bg-gray-100 dark:hover:bg-gray-700/60 border border-transparent hover:border-gray-200 dark:hover:border-gray-700 transition select-none"
            >
              <span className="truncate max-w-[80px]">{activeProviderId}</span>
              <ChevronDown className={`w-2.5 h-2.5 opacity-60 transition-transform ${isProviderOpen ? 'rotate-180' : ''}`} />
            </button>

            {isProviderOpen && (
              <div
                role="menu"
                aria-label="Daftar Provider AI"
                className="absolute left-0 mt-1.5 w-52 bg-white dark:bg-gray-850 rounded-lg shadow-lg border border-gray-200 dark:border-gray-700 py-1 z-50 select-none animate-in fade-in-0 zoom-in-95 duration-100"
              >
                <div className="px-2.5 py-1 text-[10px] font-semibold text-gray-400 dark:text-gray-500 uppercase tracking-wider border-b border-gray-100 dark:border-gray-800">
                  Pilih Provider AI
                </div>
                <div className="max-h-56 overflow-y-auto py-0.5">
                  {AVAILABLE_PROVIDERS.map(p => {
                    const isSelected = p.id === activeProviderId;
                    return (
                      <button
                        key={p.id}
                        role="menuitem"
                        type="button"
                        onClick={() => {
                          onSelectProvider?.(p.id);
                          setIsProviderOpen(false);
                        }}
                        className={`w-full flex items-center justify-between px-2.5 py-1.5 text-left text-xs transition-colors ${
                          isSelected
                            ? 'bg-blue-50 dark:bg-blue-950/50 text-blue-700 dark:text-blue-300 font-semibold'
                            : 'text-gray-700 dark:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-800'
                        }`}
                      >
                        <div className="flex flex-col min-w-0 pr-1">
                          <span className="truncate">{p.label}</span>
                          <span className="text-[9px] text-gray-400 dark:text-gray-500 font-normal truncate">
                            {p.badge}
                          </span>
                        </div>
                        {isSelected && <Check className="w-3.5 h-3.5 text-blue-600 dark:text-blue-400 shrink-0" />}
                      </button>
                    );
                  })}
                </div>
                <div className="border-t border-gray-100 dark:border-gray-800 pt-1 px-1">
                  <button
                    type="button"
                    onClick={() => {
                      setIsProviderOpen(false);
                      onOpenSettings();
                    }}
                    className="w-full flex items-center gap-1.5 px-2 py-1 text-[11px] text-gray-600 dark:text-gray-400 hover:text-blue-600 dark:hover:text-blue-400 hover:bg-gray-50 dark:hover:bg-gray-800/50 rounded transition"
                  >
                    <Settings className="w-3 h-3" />
                    <span>Pengaturan API Key...</span>
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>

      <div className="flex items-center gap-1 shrink-0">
        {/* Copilot vs Autopilot Toggle */}
        <button
          onClick={onToggleMode}
          title={executionMode === 'copilot' ? 'Mode Copilot (Preview dulu)' : 'Mode Autopilot (Langsung eksekusi)'}
          className={`flex items-center gap-1 text-[10px] px-2 py-0.5 rounded-full font-medium transition-all shadow-2xs select-none ${
            executionMode === 'autopilot'
              ? 'bg-amber-100 dark:bg-amber-900/50 text-amber-800 dark:text-amber-200 border border-amber-300 dark:border-amber-700'
              : 'bg-gray-100 dark:bg-gray-700 text-gray-600 dark:text-gray-300 hover:bg-gray-200 dark:hover:bg-gray-600 border border-gray-200/60 dark:border-gray-600/60'
          }`}
        >
          {executionMode === 'autopilot' ? (
            <>
              <Play className="w-2.5 h-2.5 fill-current" /> Auto
            </>
          ) : (
            <>
              <Zap className="w-2.5 h-2.5" /> Copilot
            </>
          )}
        </button>

        {/* Theme Mode Toggle Button */}
        {onToggleTheme && (
          <button
            onClick={onToggleTheme}
            className={`p-1 rounded-md transition-all border ${
              themeMode === 'dark'
                ? 'text-indigo-400 bg-indigo-950/40 border-indigo-800/50 hover:bg-indigo-900/50'
                : themeMode === 'light'
                ? 'text-amber-500 bg-amber-500/10 border-amber-400/40 hover:bg-amber-500/20'
                : 'text-gray-500 dark:text-gray-400 bg-transparent border-transparent hover:bg-gray-100 dark:hover:bg-gray-700'
            }`}
            title={
              themeMode === 'dark'
                ? 'Tema: Gelap (Klik untuk mode Otomatis)'
                : themeMode === 'light'
                ? 'Tema: Terang (Klik untuk mode Gelap)'
                : 'Tema: Otomatis Office/Sistem (Klik untuk mode Terang)'
            }
            aria-label="Ganti Tema"
          >
            {themeMode === 'dark' ? (
              <Moon className="w-3.5 h-3.5 text-indigo-400" />
            ) : themeMode === 'light' ? (
              <Sun className="w-3.5 h-3.5 text-amber-500" />
            ) : (
              <Monitor className="w-3.5 h-3.5 text-gray-500 dark:text-gray-400" />
            )}
          </button>
        )}

        {/* Settings Button */}
        <button
          onClick={onOpenSettings}
          className="p-1 text-gray-500 hover:text-gray-800 dark:hover:text-gray-200 rounded-md hover:bg-gray-100 dark:hover:bg-gray-700 transition"
          title="Pengaturan BYOK API Key"
        >
          <Settings className="w-3.5 h-3.5" />
        </button>
      </div>
    </header>
  );
};
