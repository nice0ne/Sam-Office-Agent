import React from 'react';
import { ToolCall } from '../../types';
import { CheckCircle2, Play, AlertCircle } from 'lucide-react';

export interface ActionCardProps {
  toolCall: ToolCall;
  onApply: (toolCall: ToolCall) => void;
  isExecuting?: boolean;
}

export const ActionCard: React.FC<ActionCardProps> = ({ toolCall, onApply, isExecuting }) => {
  const isApplied = toolCall.status === 'applied';
  const isFailed = toolCall.status === 'failed';
  const args = toolCall.arguments || {};

  return (
    <div className="mt-2 p-3 bg-gray-50 dark:bg-gray-900 border border-gray-200 dark:border-gray-700 rounded-lg text-xs space-y-2">
      <div className="flex items-center justify-between font-semibold text-gray-800 dark:text-gray-200">
        <span className="capitalize">{toolCall.name.replace(/_/g, ' ')}</span>
        {isApplied && (
          <span className="flex items-center gap-1 text-green-600 dark:text-green-400 font-medium">
            <CheckCircle2 className="w-3.5 h-3.5" /> Diterapkan
          </span>
        )}
        {isFailed && (
          <span className="flex items-center gap-1 text-red-600 dark:text-red-400 font-medium">
            <AlertCircle className="w-3.5 h-3.5" /> Gagal
          </span>
        )}
      </div>

      <div className="bg-white dark:bg-gray-800 p-2 rounded border border-gray-100 dark:border-gray-700 font-mono text-[11px] text-gray-700 dark:text-gray-300 max-h-24 overflow-y-auto select-text">
        {toolCall.name === 'write_cells' && (
          <div>
            <span className="text-blue-500">Range:</span> {args.range}
            {args.formula && (
              <div><span className="text-green-500">Formula:</span> {args.formula}</div>
            )}
            {args.values && (
              <div><span className="text-purple-500">Data:</span> {Array.isArray(args.values) ? `${args.values.length} baris data` : 'Tabel data'}</div>
            )}
          </div>
        )}
        {toolCall.name === 'format_range' && (
          <div>
            <span className="text-blue-500">Range:</span> {args.range}
            {args.numberFormat && (
              <div><span className="text-purple-500">Format:</span> {args.numberFormat}</div>
            )}
          </div>
        )}
        {toolCall.name === 'create_chart' && (
          <div>
            <span className="text-amber-500">Chart:</span> {args.chartType} ({args.dataRange})
          </div>
        )}
        {toolCall.name !== 'write_cells' && toolCall.name !== 'format_range' && toolCall.name !== 'create_chart' && (
          <pre className="whitespace-pre-wrap">{JSON.stringify(args, null, 2)}</pre>
        )}
      </div>

      {isFailed && toolCall.error && (
        <div className="text-red-500 text-[10px]">{toolCall.error}</div>
      )}

      {!isApplied && (
        <button
          onClick={() => onApply(toolCall)}
          disabled={isExecuting}
          className="w-full py-1.5 px-3 bg-blue-600 hover:bg-blue-700 text-white rounded font-medium flex items-center justify-center gap-1.5 transition disabled:opacity-50"
        >
          {isExecuting ? 'Menerapkan...' : (
            <>
              <Play className="w-3 h-3 fill-current" /> Terapkan ke Dokumen
            </>
          )}
        </button>
      )}
    </div>
  );
};
