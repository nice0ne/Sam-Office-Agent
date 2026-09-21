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
    <div className="mt-1.5 p-2 bg-gray-50/95 dark:bg-gray-900/95 border border-gray-200/80 dark:border-gray-700/80 rounded-lg text-xs space-y-1.5 shadow-2xs">
      <div className="flex items-center justify-between font-semibold text-gray-800 dark:text-gray-200 text-[11px]">
        <span className="capitalize">{toolCall.name.replace(/_/g, ' ')}</span>
        {isApplied && (
          <span className="flex items-center gap-1 text-green-600 dark:text-green-400 font-medium text-[10px]">
            <CheckCircle2 className="w-3 h-3" /> Diterapkan
          </span>
        )}
        {isFailed && (
          <span className="flex items-center gap-1 text-red-600 dark:text-red-400 font-medium text-[10px]">
            <AlertCircle className="w-3 h-3" /> Gagal
          </span>
        )}
      </div>

      <div className="bg-white dark:bg-gray-800 p-1.5 rounded border border-gray-100 dark:border-gray-750 font-mono text-[10px] text-gray-700 dark:text-gray-300 max-h-24 overflow-y-auto select-text leading-tight">
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
        {toolCall.name === 'insert_chart_image' && (
          <div>
            <span className="text-amber-500">Grafik Word:</span> {args.chartType?.toUpperCase()} {args.title ? `- "${args.title}"` : ''}
            {Array.isArray(args.labels) && (
              <div><span className="text-purple-500">Data:</span> {args.labels.length} item ({args.labels.slice(0, 4).join(', ')}{args.labels.length > 4 ? '...' : ''})</div>
            )}
          </div>
        )}
        {toolCall.name === 'insert_page_break' && (
          <div>
            <span className="text-indigo-500">Pemisah:</span> {args.breakType === 'section' ? 'Section Break (Halaman Berikutnya)' : 'Page Break (Halaman Baru)'}
          </div>
        )}
        {toolCall.name === 'find_and_replace' && (
          <div>
            <div><span className="text-blue-500">Cari:</span> &quot;{args.findText}&quot;</div>
            <div><span className="text-green-500">Ganti dengan:</span> &quot;{args.replaceText}&quot;</div>
          </div>
        )}
        {toolCall.name === 'create_slide' && (
          <div>
            <div className="font-semibold text-orange-500">Slide: {args.title}</div>
            {Array.isArray(args.bullets) && args.bullets.length > 0 && (
              <ul className="mt-1 space-y-0.5 text-[10px] text-gray-600 dark:text-gray-300">
                {args.bullets.slice(0, 3).map((b: string, i: number) => (
                  <li key={i} className="truncate">• {b}</li>
                ))}
                {args.bullets.length > 3 && <li className="text-gray-400">...dan {args.bullets.length - 3} poin lainnya</li>}
              </ul>
            )}
            {args.notes && (
              <div className="mt-1 text-[10px] text-gray-400 italic truncate">Notes: {args.notes}</div>
            )}
          </div>
        )}
        {toolCall.name === 'create_presentation_deck' && (
          <div>
            <div className="font-semibold text-orange-500">Presentasi ({Array.isArray(args.slides) ? args.slides.length : 0} Slide):</div>
            {Array.isArray(args.slides) && (
              <ol className="mt-1 space-y-0.5 text-[10px] text-gray-600 dark:text-gray-300">
                {args.slides.slice(0, 4).map((s: any, i: number) => (
                  <li key={i} className="truncate font-medium">{i + 1}. {s.title || `Slide ${i + 1}`}</li>
                ))}
                {args.slides.length > 4 && <li className="text-gray-400">...dan {args.slides.length - 4} slide lainnya</li>}
              </ol>
            )}
          </div>
        )}
        {toolCall.name === 'read_slides' && (
          <div>
            <div className="font-semibold text-blue-500">Membaca Slide PowerPoint</div>
            <div className="text-[11px] text-gray-600 dark:text-gray-300">
              {args.slideNumber ? `Membaca Slide #${args.slideNumber}` : (args.allSlides ? 'Membaca seluruh slide presentasi' : 'Membaca slide aktif')}
            </div>
          </div>
        )}
        {toolCall.name === 'read_sheet' && (
          <div>
            <div className="font-semibold text-blue-500">Membaca Lembar Kerja Excel</div>
            <div className="text-[11px] text-gray-600 dark:text-gray-300">
              {args.range ? `Range: ${args.range}` : 'Membaca area data aktif'}
            </div>
          </div>
        )}
        {toolCall.name !== 'write_cells' &&
          toolCall.name !== 'format_range' &&
          toolCall.name !== 'create_chart' &&
          toolCall.name !== 'insert_chart_image' &&
          toolCall.name !== 'insert_page_break' &&
          toolCall.name !== 'find_and_replace' &&
          toolCall.name !== 'create_slide' &&
          toolCall.name !== 'create_presentation_deck' &&
          toolCall.name !== 'read_slides' &&
          toolCall.name !== 'read_sheet' && (
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
          className="w-full py-1 px-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded text-[11px] font-medium flex items-center justify-center gap-1.5 transition disabled:opacity-50"
        >
          {isExecuting ? (toolCall.name.startsWith('read_') ? 'Membaca...' : 'Menerapkan...') : (
            <>
              <Play className="w-2.5 h-2.5 fill-current" /> {toolCall.name.startsWith('read_') ? 'Baca Data Sekarang' : 'Terapkan ke Dokumen'}
            </>
          )}
        </button>
      )}
    </div>
  );
};
