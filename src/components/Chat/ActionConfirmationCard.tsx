import React from 'react';
import { ShieldCheck, AlertTriangle, Layers, XCircle, Check } from 'lucide-react';
import { PendingActionProposal } from '../../types';

export interface ActionConfirmationCardProps {
  proposal: PendingActionProposal;
  onConfirm: (proposalId: string) => void;
  onCancel: (proposalId: string) => void;
  disabled?: boolean;
}

export const ActionConfirmationCard: React.FC<ActionConfirmationCardProps> = ({
  proposal,
  onConfirm,
  onCancel,
  disabled = false,
}) => {
  const getActionBadgeColor = () => {
    switch (proposal.actionType) {
      case 'modify_cells':
        return 'bg-amber-100 text-amber-800 dark:bg-amber-900/40 dark:text-amber-300 border-amber-200 dark:border-amber-800';
      case 'clean_data':
        return 'bg-blue-100 text-blue-800 dark:bg-blue-900/40 dark:text-blue-300 border-blue-200 dark:border-blue-800';
      case 'format_cells':
        return 'bg-purple-100 text-purple-800 dark:bg-purple-900/40 dark:text-purple-300 border-purple-200 dark:border-purple-800';
      default:
        return 'bg-orange-100 text-orange-800 dark:bg-orange-900/40 dark:text-orange-300 border-orange-200 dark:border-orange-800';
    }
  };

  const getActionTypeLabel = () => {
    switch (proposal.actionType) {
      case 'modify_cells':
        return 'Modifikasi Sel';
      case 'clean_data':
        return 'Pembersihan Data';
      case 'format_cells':
        return 'Format Tampilan';
      default:
        return 'Aksi Excel';
    }
  };

  return (
    <div className="mt-2.5 p-3 rounded-lg border border-amber-300/80 dark:border-amber-700/80 bg-gradient-to-br from-amber-50/60 to-orange-50/40 dark:from-amber-950/20 dark:to-orange-950/20 shadow-sm text-xs">
      {/* Header */}
      <div className="flex items-center justify-between gap-2 mb-2">
        <div className="flex items-center gap-1.5 font-semibold text-amber-900 dark:text-amber-200 text-[11px]">
          <ShieldCheck className="w-4 h-4 text-amber-600 dark:text-amber-400 shrink-0" />
          <span>Konfirmasi Eksekusi Aman</span>
        </div>
        <div className="flex items-center gap-1.5">
          <span className={`inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium border ${getActionBadgeColor()}`}>
            {getActionTypeLabel()}
          </span>
          <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-[10px] font-medium bg-orange-100 dark:bg-orange-900/40 text-orange-800 dark:text-orange-300 border border-orange-200 dark:border-orange-800">
            <AlertTriangle className="w-3 h-3 text-orange-600 dark:text-orange-400" />
            <span>{`${proposal.affectedCellsCount} sel`}</span>
          </span>
        </div>
      </div>

      {/* Description */}
      <div className="text-gray-700 dark:text-gray-200 font-medium mb-2 leading-relaxed">
        {proposal.description}
      </div>

      {/* Detail / Target Range */}
      {proposal.targetRange && (
        <div className="flex items-center gap-1.5 mb-2.5 text-[11px] text-gray-600 dark:text-gray-300 font-mono bg-white/70 dark:bg-gray-800/70 px-2 py-1 rounded border border-gray-200/80 dark:border-gray-700">
          <Layers className="w-3.5 h-3.5 text-gray-500 shrink-0" />
          <span className="text-gray-500 dark:text-gray-400">Target Range:</span>
          <span className="font-semibold text-gray-800 dark:text-gray-100">{proposal.targetRange}</span>
        </div>
      )}

      {/* Action Buttons */}
      <div className="flex items-center justify-end gap-2 pt-1 border-t border-amber-200/60 dark:border-amber-800/40">
        <button
          type="button"
          onClick={() => onCancel(proposal.id)}
          disabled={disabled}
          className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-[11px] font-medium text-gray-700 dark:text-gray-300 bg-white dark:bg-gray-800 border border-gray-300 dark:border-gray-600 hover:bg-gray-100 dark:hover:bg-gray-700 disabled:opacity-50 disabled:cursor-not-allowed transition-colors shadow-xs"
        >
          <XCircle className="w-3.5 h-3.5 text-gray-500" />
          <span>Batalkan</span>
        </button>
        <button
          type="button"
          onClick={() => onConfirm(proposal.id)}
          disabled={disabled}
          className="inline-flex items-center gap-1 px-3 py-1 rounded-md text-[11px] font-medium text-white bg-emerald-600 hover:bg-emerald-700 dark:bg-emerald-600 dark:hover:bg-emerald-500 disabled:opacity-50 disabled:cursor-not-allowed transition-colors shadow-xs"
        >
          <Check className="w-3.5 h-3.5" />
          <span>Terapkan Perubahan</span>
        </button>
      </div>
    </div>
  );
};
