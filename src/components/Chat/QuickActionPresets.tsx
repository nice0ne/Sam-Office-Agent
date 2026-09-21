import React from 'react';
import { HostType } from '../../types';

export interface QuickPresetItem {
  id: string;
  label: string;
  prompt: string;
}

export interface QuickActionPresetsProps {
  host: HostType;
  onSelectPreset: (prompt: string) => void;
  disabled?: boolean;
}

export const PRESETS_BY_HOST: Record<HostType, QuickPresetItem[]> = {
  Excel: [
    {
      id: 'excel_clean_data',
      label: '✨ Bersihkan Data',
      prompt: 'Bersihkan data di tabel aktif: hapus duplikat, rapikan spasi berlebih, dan isi sel yang kosong.',
    },
    {
      id: 'excel_rekap_chart',
      label: '📊 Rekap & Chart',
      prompt: 'Buatkan ringkasan data tabel ini ke kolom kosong di sebelah kanan dan buatkan grafik visualisasinya.',
    },
    {
      id: 'excel_format_color',
      label: '🎨 Format & Color Scale',
      prompt: 'Terapkan format tebal pada baris header dan terapkan conditional formatting color scale pada kolom angka.',
    },
  ],
  Word: [
    {
      id: 'word_mom',
      label: '📝 Notulen Rapat (MoM)',
      prompt: 'Buatkan dokumen terstruktur Notulen Rapat (Minutes of Meeting) lengkap dengan agenda, pembahasan, dan action items.',
    },
    {
      id: 'word_sop',
      label: '📑 Buat SOP',
      prompt: 'Buatkan dokumen Standar Operasional Prosedur (SOP) terstruktur dengan tujuan, ruang lingkup, dan langkah-langkah alur kerja.',
    },
    {
      id: 'word_polish',
      label: '✍️ Poles Bahasa & EYD',
      prompt: 'Poles dan perbaiki tata bahasa dokumen/seleksi aktif agar sesuai dengan standar EYD formal bisnis profesional.',
    },
  ],
  PowerPoint: [
    {
      id: 'ppt_deck_3',
      label: '📑 Buat Deck 3 Slide',
      prompt: 'Buatkan 3 slide presentasi terstruktur: Slide 1 Judul Cover, Slide 2 Poin Materi Utama, Slide 3 Key Metrics Capaian.',
    },
    {
      id: 'ppt_summarize',
      label: '📋 Ringkas Seluruh Slide',
      prompt: 'Baca seluruh slide pada presentasi ini lalu buatkan ringkasan eksekutif komprehensif.',
    },
    {
      id: 'ppt_speaker_notes',
      label: '🎤 Catatan Pemateri',
      prompt: 'Buatkan naskah catatan pemateri (speaker notes) yang elegan dan siap pakai untuk materi presentasi ini.',
    },
  ],
  BrowserDev: [
    {
      id: 'dev_clean_data',
      label: '✨ Bersihkan Data',
      prompt: 'Bersihkan data di tabel aktif: hapus duplikat, rapikan spasi berlebih, dan isi sel yang kosong.',
    },
    {
      id: 'dev_rekap_chart',
      label: '📊 Rekap & Chart',
      prompt: 'Buatkan ringkasan data tabel ini ke kolom kosong di sebelah kanan dan buatkan grafik visualisasinya.',
    },
    {
      id: 'dev_format_color',
      label: '🎨 Format & Color Scale',
      prompt: 'Terapkan format tebal pada baris header dan terapkan conditional formatting color scale pada kolom angka.',
    },
  ],
};

export const QuickActionPresets: React.FC<QuickActionPresetsProps> = ({
  host,
  onSelectPreset,
  disabled = false,
}) => {
  const presets = PRESETS_BY_HOST[host] || PRESETS_BY_HOST.Excel;

  return (
    <div
      aria-label="Quick Action Presets"
      className="flex items-center gap-1.5 overflow-x-auto pb-2 mb-1 scrollbar-none no-scrollbar text-xs"
    >
      {presets.map(preset => (
        <button
          key={preset.id}
          type="button"
          disabled={Boolean(disabled)}
          onClick={() => onSelectPreset(preset.prompt)}
          className="px-2.5 py-1 rounded-full text-[11px] font-medium bg-gray-100 dark:bg-gray-750 hover:bg-blue-50 hover:text-blue-600 dark:hover:bg-blue-900/30 dark:hover:text-blue-400 text-gray-700 dark:text-gray-300 transition-colors border border-gray-200/80 dark:border-gray-700/80 shrink-0 whitespace-nowrap disabled:opacity-50 disabled:cursor-not-allowed select-none shadow-2xs"
        >
          {preset.label}
        </button>
      ))}
    </div>
  );
};
