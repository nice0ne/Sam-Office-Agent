import React from 'react';
import { HostType, ProviderId, ExecutionMode } from '../../types';
import { Settings, Zap, Play, FileSpreadsheet, FileText, Presentation } from 'lucide-react';

interface HeaderProps {
  host: HostType;
  activeProviderId: ProviderId;
  executionMode: ExecutionMode;
  onToggleMode: () => void;
  onOpenSettings: () => void;
}

export const Header: React.FC<HeaderProps> = ({
  host,
  activeProviderId,
  executionMode,
  onToggleMode,
  onOpenSettings,
}) => {
  const getHostIcon = () => {
    switch (host) {
      case 'Excel':
        return <FileSpreadsheet className="w-5 h-5 text-office-excel" />;
      case 'Word':
        return <FileText className="w-5 h-5 text-office-word" />;
      case 'PowerPoint':
        return <Presentation className="w-5 h-5 text-office-ppt" />;
      default:
        return <Zap className="w-5 h-5 text-office-agent" />;
    }
  };

  return (
    <header className="flex items-center justify-between px-4 py-3 bg-white dark:bg-gray-800 border-b border-gray-200 dark:border-gray-700 shadow-sm">
      <div className="flex items-center space-x-2">
        {getHostIcon()}
        <div>
          <h1 className="text-sm font-bold text-gray-900 dark:text-gray-100 flex items-center gap-1">
            Sam Office
            <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-blue-100 dark:bg-blue-900/60 text-blue-700 dark:text-blue-300 font-medium">
              {host}
            </span>
          </h1>
          <p className="text-[11px] text-gray-500 capitalize">{activeProviderId}</p>
        </div>
      </div>

      <div className="flex items-center space-x-2">
        {/* Copilot vs Autopilot Toggle */}
        <button
          onClick={onToggleMode}
          title={executionMode === 'copilot' ? 'Mode Copilot (Preview dulu)' : 'Mode Autopilot (Langsung eksekusi)'}
          className={`flex items-center gap-1 text-xs px-2.5 py-1 rounded-full font-medium transition-colors ${
            executionMode === 'autopilot'
              ? 'bg-amber-100 dark:bg-amber-900/50 text-amber-800 dark:text-amber-200 border border-amber-300'
              : 'bg-gray-100 dark:bg-gray-700 text-gray-700 dark:text-gray-300'
          }`}
        >
          {executionMode === 'autopilot' ? (
            <>
              <Play className="w-3 h-3 fill-current" /> Auto
            </>
          ) : (
            <>
              <Zap className="w-3 h-3" /> Copilot
            </>
          )}
        </button>

        <button
          onClick={onOpenSettings}
          className="p-1.5 text-gray-500 hover:text-gray-800 dark:hover:text-gray-200 rounded-lg hover:bg-gray-100 dark:hover:bg-gray-700 transition"
          title="Pengaturan BYOK API Key"
        >
          <Settings className="w-4 h-4" />
        </button>
      </div>
    </header>
  );
};
