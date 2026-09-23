import React, { useState, useEffect } from 'react';
import { HostType, CrossAppSnapshot } from '../../types';
import {
  getLatestCrossAppSnapshot,
  dismissCrossAppSnapshot,
} from '../../services/storage/crossAppBridge';

export interface CrossAppSnapshotCardProps {
  host: HostType;
  onSendMessage?: (prompt: string) => void;
  onDismiss?: () => void;
}

export const CrossAppSnapshotCard: React.FC<CrossAppSnapshotCardProps> = ({
  host,
  onSendMessage,
  onDismiss,
}) => {
  const [snapshot, setSnapshot] = useState<CrossAppSnapshot | null>(() =>
    getLatestCrossAppSnapshot(host)
  );

  useEffect(() => {
    setSnapshot(getLatestCrossAppSnapshot(host));
  }, [host]);

  if (!snapshot) {
    return null;
  }

  const handleDismiss = () => {
    dismissCrossAppSnapshot(snapshot.id);
    setSnapshot(null);
    onDismiss?.();
  };

  const getActionPrompt = (): string => {
    if (host === 'Word') {
      return `Ambil snapshot data terbaru dari ${snapshot.sourceHost} di Universal Hub, lalu buatkan tabel dan draf naskah laporannya di dokumen ini.`;
    }
    if (host === 'PowerPoint') {
      return 'Ambil data atau ringkasan terbaru dari Universal Hub dan buatkan 5 slide presentasi eksekutif lengkap dengan speaker notes.';
    }
    if (host === 'Excel') {
      return 'Ambil data snapshot dari Universal Hub dan tuliskan datanya ke lembar kerja aktif.';
    }
    return `Ambil snapshot data terbaru dari ${snapshot.sourceHost} di Universal Hub.`;
  };

  const getActionButtonLabel = (): string => {
    if (host === 'Word') {
      return '📄 Buat Laporan & Sisipkan Tabel';
    }
    if (host === 'PowerPoint') {
      return '📊 Buat 5 Slide Presentasi';
    }
    if (host === 'Excel') {
      return '📋 Tulis ke Lembar Kerja';
    }
    return '📥 Impor dari Hub';
  };

  const getHostBadgeStyle = (sourceHost: HostType): string => {
    switch (sourceHost) {
      case 'Excel':
        return 'bg-emerald-100 text-emerald-800 dark:bg-emerald-950/60 dark:text-emerald-300 border-emerald-300 dark:border-emerald-800';
      case 'Word':
        return 'bg-blue-100 text-blue-800 dark:bg-blue-950/60 dark:text-blue-300 border-blue-300 dark:border-blue-800';
      case 'PowerPoint':
        return 'bg-amber-100 text-amber-800 dark:bg-amber-950/60 dark:text-amber-300 border-amber-300 dark:border-amber-800';
      default:
        return 'bg-purple-100 text-purple-800 dark:bg-purple-950/60 dark:text-purple-300 border-purple-300 dark:border-purple-800';
    }
  };

  const rowCount = snapshot.tableData
    ? snapshot.tableData.totalRows || snapshot.tableData.rows?.length || 0
    : 0;

  const detailsText = snapshot.tableData
    ? `Tabel ${rowCount} baris${snapshot.tableData.headers?.length ? `, ${snapshot.tableData.headers.length} kolom` : ''}${
        snapshot.summaryText ? ` • ${snapshot.summaryText}` : ''
      }`
    : snapshot.summaryText || 'Data siap diimpor ke dokumen.';

  return (
    <div
      data-testid="cross-app-snapshot-card"
      className="p-3 mb-2 rounded-lg border border-blue-200/80 dark:border-blue-800/60 bg-blue-50/70 dark:bg-blue-950/30 shadow-xs transition-all animate-fadeIn"
    >
      <div className="flex items-start justify-between gap-2">
        <div className="flex items-center gap-1.5 min-w-0">
          <span
            className={`px-1.5 py-0.5 rounded text-[10px] font-semibold border uppercase tracking-wider shrink-0 ${getHostBadgeStyle(
              snapshot.sourceHost
            )}`}
          >
            {snapshot.sourceHost}
          </span>
          <span className="text-xs font-semibold text-gray-800 dark:text-gray-100 truncate">
            {snapshot.title}
          </span>
        </div>
        <button
          type="button"
          aria-label="Abaikan snapshot"
          title="Abaikan"
          onClick={handleDismiss}
          className="text-gray-400 hover:text-gray-600 dark:text-gray-500 dark:hover:text-gray-300 text-xs px-1 rounded transition-colors"
        >
          ✕
        </button>
      </div>

      <div className="mt-1 text-[11px] text-gray-600 dark:text-gray-300">
        {`Ditemukan data dari ${snapshot.sourceHost}: ${detailsText}`}
      </div>

      <div className="mt-2 flex items-center justify-end">
        <button
          type="button"
          onClick={() => onSendMessage?.(getActionPrompt())}
          className="px-2.5 py-1 text-xs font-medium bg-blue-600 hover:bg-blue-700 active:bg-blue-800 text-white rounded-md shadow-2xs transition-all active:scale-[0.98]"
        >
          {getActionButtonLabel()}
        </button>
      </div>
    </div>
  );
};
