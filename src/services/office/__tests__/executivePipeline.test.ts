import { describe, it, expect } from 'vitest';
import { synthesizeExecutiveReport } from '../executivePipeline';
import { listCrossAppSnapshots } from '../../storage/crossAppBridge';

describe('Executive Reporting Pipeline Engine', () => {
  it('synthesizes integrated Word executive narrative and PowerPoint slides from tabular data', async () => {
    const tableData = {
      headers: ['Divisi', 'Target (M)', 'Realisasi (M)', 'Persentase'],
      rows: [
        ['Pemasaran & Penjualan', 50, 54.2, '108.4%'],
        ['Operasional & Logistik', 35, 33.8, '96.5%'],
        ['Teknologi Informasi', 20, 21.0, '105.0%'],
      ],
    };

    const metrics = [
      { label: 'Total Realisasi', value: 'Rp 109,0 M', trend: 'up' as const },
      { label: 'Pencapaian Target', value: '103.8%', trend: 'up' as const },
    ];

    const result = await synthesizeExecutiveReport({
      title: 'Laporan Kinerja Eksekutif Triwulan III',
      topic: 'Evaluasi Kinerja dan Realisasi Anggaran Q3',
      tableData,
      metrics,
      targetSlideCount: 3,
      theme: 'corporate_blue',
      executiveSummary: 'Kinerja operasional dan finansial konsisten melampaui target tahunan dengan realisasi rata-rata 103.8%.',
    });

    expect(result.title).toBe('Laporan Kinerja Eksekutif Triwulan III');
    expect(result.wordDocument.formattedNarrative).toContain('Laporan Kinerja Eksekutif Triwulan III');
    expect(result.wordDocument.formattedNarrative).toContain('Pemasaran & Penjualan');
    expect(result.wordDocument.sectionsCount).toBeGreaterThanOrEqual(3);

    expect(result.powerPointDeck.totalSlides).toBeGreaterThanOrEqual(3);
    expect(result.powerPointDeck.slides[0].title).toContain('Laporan Kinerja Eksekutif');
    expect(result.powerPointDeck.slides[0].speakerScript).toBeDefined();

    expect(result.hubSnapshotId).toBeDefined();
    const snapshots = listCrossAppSnapshots();
    expect(snapshots.some((s) => s.id === result.hubSnapshotId)).toBe(true);
    expect(result.summaryMessage).toContain('Berhasil');
  });

  it('handles report generation gracefully when tableData is omitted with fallback', async () => {
    const result = await synthesizeExecutiveReport({
      title: 'Ringkasan Rapat Dewan Direksi',
      executiveSummary: 'Seluruh inisiatif strategis disetujui untuk dieksekusi pada kuartal mendatang.',
    });

    expect(result.wordDocument.formattedNarrative).toContain('Ringkasan Rapat Dewan Direksi');
    expect(result.powerPointDeck.totalSlides).toBeGreaterThanOrEqual(2);
  });
});
