import { describe, it, expect } from 'vitest';
import { synthesizeDocToDeck } from '../src/services/office/docToDeckTransformer';

describe('synthesizeDocToDeck', () => {
  it('partitions raw report text into 5-slide Executive Storyline Arc', () => {
    const rawText = `Laporan Inisiatif Transformasi Digital 2026
Latar Belakang: Proses pelaporan manual membutuhkan waktu 4 hari per minggu dan rawan kesalahan manusia.
Solusi: Penerapan Sam Office Agent untuk otomatisasi dokumen Excel, Word, dan PowerPoint.
Capaian & Metrik: Peningkatan efisiensi waktu hingga 75%, akurasi data mencapai 99.8%, kepuasan pengguna 95%.
Rencana Aksi: Sosialisasi seluruh divisi pada Oktober 2026 dan evaluasi triwulanan.`;

    const result = synthesizeDocToDeck(rawText, { theme: 'corporate_blue' });

    expect(result.totalSlidesCreated).toBe(5);
    expect(result.deckTitle).toContain('Transformasi Digital');
    expect(result.appliedTheme).toBe('corporate_blue');

    const categories = result.slides.map(s => s.category);
    expect(categories).toEqual(['cover', 'context', 'strategy', 'metrics', 'roadmap']);

    for (const slide of result.slides) {
      expect(slide.title.length).toBeGreaterThan(0);
      expect(slide.speakerScript.hook.length).toBeGreaterThan(0);
      expect(slide.speakerScript.keyTalkingPoints.length).toBeGreaterThan(0);
      expect(slide.speakerScript.transition.length).toBeGreaterThan(0);
    }
  });

  it('handles empty or brief text gracefully with default briefing structure', () => {
    const result = synthesizeDocToDeck('', { presentationTitle: 'Briefing Eksekutif' });
    expect(result.totalSlidesCreated).toBeGreaterThanOrEqual(4);
    expect(result.deckTitle).toBe('Briefing Eksekutif');
    expect(result.slides[0].category).toBe('cover');
  });

  it('correctly parses inline sentences with metrics and assigns themes', () => {
    const inlineDoc = 'Proposal Modernisasi TI. Masalah: Sistem warisan lambat. Solusi: Migrasi cloud modern. Metrik: Uptime 99.9%. Rencana Aksi: Rollout Q4.';
    const result = synthesizeDocToDeck(inlineDoc, { theme: 'emerald_executive' });

    expect(result.deckTitle).toBe('Proposal Modernisasi TI');
    expect(result.appliedTheme).toBe('emerald_executive');
    expect(result.totalSlidesCreated).toBe(5);

    const metricsSlide = result.slides.find((s) => s.category === 'metrics');
    expect(metricsSlide).toBeDefined();
    expect(metricsSlide?.metrics?.length).toBeGreaterThan(0);
    expect(metricsSlide?.metrics?.[0].value).toContain('99.9%');
  });
});
