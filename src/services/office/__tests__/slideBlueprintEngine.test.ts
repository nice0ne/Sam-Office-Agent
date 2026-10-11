import { describe, it, expect } from 'vitest';
import {
  SLIDESPEAK_THEME_PALETTES,
  formatActionTitle,
  buildSequoiaPitchDeck,
} from '../slideBlueprintEngine';

describe('Slide Blueprint Engine (SlideSpeak Themes)', () => {
  describe('SLIDESPEAK_THEME_PALETTES', () => {
    it('defines colors and typography tokens for all 10 curated themes', () => {
      const expectedThemes = [
        'corporate_blue',
        'emerald_executive',
        'modern_dark',
        'minimalist_clean',
        'mckinsey_consulting',
        'demo_day',
        'linear_dark',
        'stripe_saas',
        'ted_stage',
        'sequoia_pitch',
      ];

      for (const t of expectedThemes) {
        const theme = SLIDESPEAK_THEME_PALETTES[t];
        expect(theme).toBeDefined();
        expect(theme.primary).toMatch(/^#[0-9A-Fa-f]{6}$/);
        expect(theme.bg).toMatch(/^#[0-9A-Fa-f]{6}$/);
        expect(theme.text).toMatch(/^#[0-9A-Fa-f]{6}$/);
        expect(theme.accent).toMatch(/^#[0-9A-Fa-f]{6}$/);
      }
    });

    it('has high contrast dark canvas for linear_dark and ted_stage', () => {
      expect(SLIDESPEAK_THEME_PALETTES.linear_dark.bg).toBe('#08090A');
      expect(SLIDESPEAK_THEME_PALETTES.ted_stage.bg).toBe('#0A0A0A');
    });

    it('has distinct consulting signature colors for mckinsey_consulting and demo_day', () => {
      expect(SLIDESPEAK_THEME_PALETTES.mckinsey_consulting.primary).toBe('#051C2C');
      expect(SLIDESPEAK_THEME_PALETTES.demo_day.accent).toBe('#FF6B00');
    });
  });

  describe('formatActionTitle', () => {
    it('turns topic and takeaway into McKinsey-style full sentence action title', () => {
      const title = formatActionTitle(
        'Kinerja Operasional Q3',
        'Efisiensi meningkat 35% didorong oleh otomatisasi sistem terintegrasi'
      );
      expect(title).toBe('Kinerja Operasional Q3: Efisiensi Meningkat 35% Didorong Oleh Otomatisasi Sistem Terintegrasi');
    });

    it('handles standalone topic gracefully', () => {
      const title = formatActionTitle('Strategi Penetrasi Pasar Digital');
      expect(title).toBe('Strategi Penetrasi Pasar Digital');
    });
  });

  describe('buildSequoiaPitchDeck', () => {
    it('builds a 10-slide venture capital pitch deck adhering to Sequoia Capital framework', () => {
      const deck = buildSequoiaPitchDeck({
        companyName: 'Nusantara AI Cloud',
        oneLiner: 'Infrastruktur AI berdaulat untuk perusahaan dan instansi Indonesia',
        problem: 'Biaya komputasi luar negeri tinggi dan isu kepatuhan privasi data perbankan lokal',
        solution: 'Penyedia cloud GPU lokal dengan enkripsi berstandar perbankan dan biaya 50% lebih hemat',
        targetMarket: 'Rp 45 Triliun belanja IT enterprise nasional pada tahun 2027',
      });

      expect(deck.length).toBe(10);
      expect(deck[0].title).toContain('Nusantara AI Cloud');
      expect(deck[0].category).toBe('cover');

      // Verify Sequoia canonical steps
      const titles = deck.map((s) => s.title.toLowerCase());
      expect(titles.some((t) => t.includes('tujuan') || t.includes('purpose') || t.includes('visi'))).toBe(true);
      expect(titles.some((t) => t.includes('masalah') || t.includes('problem'))).toBe(true);
      expect(titles.some((t) => t.includes('solusi') || t.includes('solution'))).toBe(true);
      expect(titles.some((t) => t.includes('pasar') || t.includes('market'))).toBe(true);

      // Verify speaker scripts are present on each slide
      for (const slide of deck) {
        expect(slide.speakerScript).toBeDefined();
        expect(slide.speakerScript.hook.length).toBeGreaterThan(0);
        expect(slide.speakerScript.transition.length).toBeGreaterThan(0);
      }
    });
  });
});
