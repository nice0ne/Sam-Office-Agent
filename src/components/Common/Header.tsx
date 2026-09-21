import React from 'react';
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
} from 'lucide-react';

export interface HeaderProps {
  host: HostType;
  activeProviderId: ProviderId;
  executionMode: ExecutionMode;
  onToggleMode: () => void;
  onOpenSettings: () => void;
  themeMode?: ThemeMode;
  onToggleTheme?: () => void;
}

export const Header: React.FC<HeaderProps> = ({
  host,
  activeProviderId,
  executionMode,
  onToggleMode,
  onOpenSettings,
  themeMode = 'auto',
  onToggleTheme,
}) => {
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
          <span
            className="text-[10px] text-gray-500 dark:text-gray-400 font-medium capitalize truncate max-w-[85px]"
            title={activeProviderId}
          >
            {activeProviderId}
          </span>
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
            className="p-1 text-gray-500 hover:text-gray-800 dark:hover:text-gray-200 rounded-md hover:bg-gray-100 dark:hover:bg-gray-700 transition"
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
