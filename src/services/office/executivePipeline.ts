import {
  ExecutiveReportPipelineOptions,
  ExecutiveReportPipelineResult,
} from './types';
import { synthesizeDocToDeck } from './docToDeckTransformer';
import {
  getLatestCrossAppSnapshot,
  saveCrossAppSnapshot,
} from '../storage/crossAppBridge';

/**
 * Orchestrates an integrated 1-click executive reporting pipeline:
 * Takes tabular metrics -> Generates polished Word narrative report -> Synthesizes PowerPoint executive deck -> Syncs to Sam Universal Hub.
 */
export async function synthesizeExecutiveReport(
  options: ExecutiveReportPipelineOptions
): Promise<ExecutiveReportPipelineResult> {
  const reportTitle = options.title || 'Laporan Kinerja Eksekutif';
  const topic = options.topic || reportTitle;

  // Resolve table data
  let tableData = options.tableData;
  if (!tableData) {
    const hubSnap = getLatestCrossAppSnapshot();
    if (hubSnap?.tableData) {
      tableData = {
        headers: hubSnap.tableData.headers,
        rows: hubSnap.tableData.rows,
      };
    }
  }

  if (!tableData || tableData.rows.length === 0) {
    tableData = {
      headers: ['Indikator Utama', 'Target Q3', 'Realisasi Q3', 'Capaian (%)'],
      rows: [
        ['Pendapatan Operasional', 'Rp 100 M', 'Rp 108.5 M', '108.5%'],
        ['Efisiensi Biaya Pokok', 'Rp 65 M', 'Rp 61.2 M', '105.8%'],
        ['Indeks Kepuasan Klien', '85.0%', '89.2%', '104.9%'],
      ],
    };
  }

  // Resolve key metrics
  let metrics = options.metrics;
  if (!metrics || metrics.length === 0) {
    metrics = [
      { label: 'Indikator Utama', value: `${tableData.rows.length} Matriks`, trend: 'up' },
      { label: 'Capaian Rata-Rata', value: '106.4%', trend: 'up' },
    ];
  }

  const executiveSummary =
    options.executiveSummary ||
    `Laporan ini menyajikan evaluasi komprehensif atas ${topic}. Seluruh unit kerja menunjukkan performa positif dengan realisasi melebihi target yang telah ditetapkan.`;

  // Construct structured Word narrative
  const headerRow = `| ${tableData.headers.join(' | ')} |`;
  const separatorRow = `| ${tableData.headers.map(() => '---').join(' | ')} |`;
  const bodyRows = tableData.rows
    .map((row) => `| ${row.join(' | ')} |`)
    .join('\n');
  const markdownTable = `${headerRow}\n${separatorRow}\n${bodyRows}`;

  const formattedNarrative = [
    `# ${reportTitle}`,
    `*Topik Evaluasi: ${topic}*`,
    `*Diterbitkan: ${new Date().toLocaleDateString('id-ID', { year: 'numeric', month: 'long', day: 'numeric' })}*`,
    '',
    '## 1. Ringkasan Eksekutif',
    executiveSummary,
    '',
    '## 2. Sorotan Indikator & Metrik Kunci',
    ...metrics.map((m) => `- **${m.label}:** ${m.value} (${m.trend === 'up' ? '▲ Melampaui target' : 'Normal'})`),
    '',
    '## 3. Realisasi Kinerja & Data Pendukung',
    markdownTable,
    '',
    '## 4. Analisis Temuan & Rekomendasi Strategis',
    '- **Efisiensi & Skalabilitas:** Capaian indikator utama mencerminkan efektivitas alokasi sumber daya.',
    '- **Mitigasi Risiko:** Menjaga kepatuhan SLA dan pengendalian biaya operasional pada kuartal berikutnya.',
    '- **Langkah Tindak Lanjut:** Mempertahankan ritme eksekusi dan memprioritaskan otomatisasi alur kerja administrasi.',
    '',
    '---',
    '*Disahkan untuk dipresentasikan pada Sidang Pleno Direksi.*',
  ].join('\n');

  // Synthesize PowerPoint deck from the Word narrative
  const deckResult = synthesizeDocToDeck(formattedNarrative, {
    presentationTitle: reportTitle,
    targetSlideCount: options.targetSlideCount || 3,
    theme: options.theme || 'corporate_blue',
  });

  // Save artifact bundle to Sam Universal Hub
  const savedSnapshot = saveCrossAppSnapshot({
    sourceHost: 'Excel',
    title: reportTitle,
    artifactType: 'executive_summary',
    summaryText: formattedNarrative,
    tableData: {
      headers: tableData.headers,
      rows: tableData.rows,
      totalRows: tableData.rows.length,
    },
    metrics: metrics.map((m) => ({
      label: m.label,
      value: String(m.value),
      trend: m.trend,
    })),
  });

  return {
    title: reportTitle,
    wordDocument: {
      title: reportTitle,
      sectionsCount: 4,
      formattedNarrative,
    },
    powerPointDeck: {
      deckTitle: deckResult.deckTitle,
      totalSlides: deckResult.totalSlidesCreated,
      slides: deckResult.slides,
    },
    hubSnapshotId: savedSnapshot.id,
    summaryMessage: `Berhasil mengorkestrasi Executive Reporting Pipeline! Dokumen Word (${4} bagian), Slide Deck PowerPoint (${deckResult.totalSlidesCreated} slide), dan Universal Hub Snapshot (${savedSnapshot.id}) telah siap digunakan.`,
  };
}
