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

export const UNIVERSAL_LEARN_SOUL_PRESET: QuickPresetItem = {
  id: 'universal_learn_soul',
  label: '🧠 Pelajari Preferensi Ini',
  prompt: 'Catat gaya bahasa, aturan format angka, dan preferensi istilah dari dokumen ini ke dalam direktif korporat (SOUL.md) agar selalu diterapkan pada pekerjaan berikutnya.',
};

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
    {
      id: 'excel_audit_formula',
      label: '🔍 Audit Formula & Error',
      prompt: 'Audit seluruh formula pada sheet ini: cari error #REF!, #VALUE!, #DIV/0!, atau sel dengan rumus tidak konsisten dan berikan rekomendasinya.',
    },
    {
      id: 'excel_data_story',
      label: '📈 Buat Analisis Tren',
      prompt: 'Analisis data pada tabel aktif ini: buatkan ringkasan eksekutif, tren pertumbuhan, dan insight temuan utama.',
    },
    {
      id: 'excel_share_hub',
      label: '📤 Bagikan ke Hub (Word/PPT)',
      prompt: 'Kirim ringkasan data dan tabel aktif ini ke Universal Hub agar bisa langsung dipakai di Word atau PowerPoint.',
    },
    {
      id: 'excel_web_search',
      label: '🌐 Riset Web & Buat Tabel',
      prompt: 'Cari data kurs valuta asing terkini (USD, EUR, SGD, JPY ke Rupiah) di internet lalu buatkan tabel komparasinya di lembar kerja aktif.',
    },
    UNIVERSAL_LEARN_SOUL_PRESET,
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
    {
      id: 'word_review_contract',
      label: '⚖️ Review Kontrak & Risiko',
      prompt: 'Audit dan review kepatuhan klausul pada dokumen/kontrak ini: periksa SLA, denda, termin pembayaran, klausul risiko tinggi, dan berikan rekomendasi perbaikan.',
    },
    {
      id: 'word_corporate_style',
      label: '🎨 Format Brand Korporat',
      prompt: 'Terapkan standarisasi gaya dan format korporat profesional (Corporate Navy) pada seluruh dokumen ini: tata hierarki heading, font, dan spasi yang rapi.',
    },
    {
      id: 'word_import_hub',
      label: '📥 Impor Data dari Excel Hub',
      prompt: 'Ambil snapshot data terbaru dari Excel di Universal Hub, lalu buatkan tabel dan draf naskah laporannya di dokumen ini.',
    },
    {
      id: 'word_web_search',
      label: '🌐 Riset Web & Tulis Laporan',
      prompt: 'Cari tren dan perkembangan industri terkini di internet, lalu buatkan naskah laporan eksekutif lengkap dengan tabel dan kesimpulannya.',
    },
    {
      id: 'word_insert_flowchart',
      label: '📐 Buat Diagram Alur (Flowchart)',
      prompt: 'Buat diagram alur proses visual (flowchart) profesional berdasarkan SOP atau proses kerja dalam dokumen ini, lalu sisipkan langsung ke naskah.',
    },
    UNIVERSAL_LEARN_SOUL_PRESET,
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
    {
      id: 'ppt_doc_to_deck',
      label: '📊 Dokumen ke Slide (Doc-to-Deck)',
      prompt: 'Ubah teks/dokumen laporan ini menjadi 5 slide presentasi eksekutif terstruktur lengkap dengan naskah pembicara (speaker notes).',
    },
    {
      id: 'ppt_import_hub',
      label: '📥 Buat Slide dari Hub (Excel/Word)',
      prompt: 'Ambil data atau ringkasan terbaru dari Universal Hub dan buatkan 5 slide presentasi eksekutif lengkap dengan speaker notes.',
    },
    {
      id: 'ppt_web_search',
      label: '🌐 Riset Data & Buat Slide',
      prompt: 'Cari data statistik dan fakta terkini di internet, lalu ubah menjadi 5 slide presentasi eksekutif lengkap dengan naskah pembicara.',
    },
    {
      id: 'ppt_insert_flowchart',
      label: '📐 Sisipkan Slide Flowchart Alur',
      prompt: 'Buatkan 1 slide visual khusus diagram alur proses (flowchart) bisnis dengan tema profesional dan naskah penjelasan pemateri.',
    },
    UNIVERSAL_LEARN_SOUL_PRESET,
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
    {
      id: 'excel_audit_formula',
      label: '🔍 Audit Formula & Error',
      prompt: 'Audit seluruh formula pada sheet ini: cari error #REF!, #VALUE!, #DIV/0!, atau sel dengan rumus tidak konsisten dan berikan rekomendasinya.',
    },
    {
      id: 'excel_data_story',
      label: '📈 Buat Analisis Tren',
      prompt: 'Analisis data pada tabel aktif ini: buatkan ringkasan eksekutif, tren pertumbuhan, dan insight temuan utama.',
    },
    {
      id: 'dev_share_hub',
      label: '📤 Bagikan ke Hub (Word/PPT)',
      prompt: 'Kirim ringkasan data dan tabel aktif ini ke Universal Hub agar bisa langsung dipakai di Word atau PowerPoint.',
    },
    {
      id: 'dev_web_search',
      label: '🌐 Riset Web & Buat Tabel',
      prompt: 'Cari data kurs valuta asing terkini (USD, EUR, SGD, JPY ke Rupiah) di internet lalu buatkan tabel komparasinya di lembar kerja aktif.',
    },
    UNIVERSAL_LEARN_SOUL_PRESET,
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
      className="flex items-center gap-1 overflow-x-auto pb-1 mb-1 scrollbar-none no-scrollbar"
    >
      {presets.map(preset => (
        <button
          key={preset.id}
          type="button"
          disabled={Boolean(disabled)}
          onClick={() => onSelectPreset(preset.prompt)}
          className="px-2 py-0.5 rounded-md text-[10px] font-medium bg-gray-100/90 dark:bg-gray-750/90 hover:bg-blue-50 hover:text-blue-600 dark:hover:bg-blue-950/40 dark:hover:text-blue-400 text-gray-700 dark:text-gray-300 transition-all border border-gray-200/70 dark:border-gray-700/60 shrink-0 whitespace-nowrap disabled:opacity-50 disabled:cursor-not-allowed select-none shadow-2xs hover:shadow-xs active:scale-[0.98]"
        >
          {preset.label}
        </button>
      ))}
    </div>
  );
};
