import { describe, it, expect, beforeEach } from 'vitest';
import {
  getSoulConfig,
  saveSoulConfig,
  addLearnedDirective,
  removeLearnedDirective,
  exportSoulToMarkdown,
  importSoulFromMarkdown,
  clearSoulConfig,
} from '../src/services/storage/soulStorage';

describe('soulStorage', () => {
  beforeEach(() => {
    clearSoulConfig();
  });

  it('provides default corporate SoulConfig', () => {
    const config = getSoulConfig();
    expect(config.enabled).toBe(true);
    expect(config.brandVoicePreset).toBe('formal_executive');
    expect(config.corporateName).toBeDefined();
    expect(config.rawSoulMarkdown).toContain('SOUL & CORPORATE BRAND DIRECTIVES');
    expect(config.learnedDirectives).toEqual([]);
  });

  it('updates SoulConfig partially and persists changes', () => {
    const updated = saveSoulConfig({
      corporateName: 'PT Samudra Mega Corpora',
      brandVoicePreset: 'modern_professional',
    });

    expect(updated.corporateName).toBe('PT Samudra Mega Corpora');
    expect(updated.brandVoicePreset).toBe('modern_professional');

    const fresh = getSoulConfig();
    expect(fresh.corporateName).toBe('PT Samudra Mega Corpora');
  });

  it('adds and removes learned directives correctly', () => {
    const directive = addLearnedDirective(
      'Gunakan istilah inisiatif alih-alih skema',
      'terminology',
      'Koreksi dari pimpinan rapat'
    );

    expect(directive.id).toBeDefined();
    expect(directive.rule).toContain('inisiatif');
    expect(directive.category).toBe('terminology');

    let config = getSoulConfig();
    expect(config.learnedDirectives.length).toBe(1);

    removeLearnedDirective(directive.id);
    config = getSoulConfig();
    expect(config.learnedDirectives.length).toBe(0);
  });

  it('exports to markdown and imports back correctly', () => {
    addLearnedDirective('Format mata uang wajib Rp', 'formatting');
    const md = exportSoulToMarkdown();

    expect(md).toContain('SOUL & CORPORATE BRAND DIRECTIVES');
    expect(md).toContain('Format mata uang wajib Rp');

    clearSoulConfig();
    expect(getSoulConfig().learnedDirectives.length).toBe(0);

    const imported = importSoulFromMarkdown(md);
    expect(imported.rawSoulMarkdown).toContain('SOUL & CORPORATE BRAND DIRECTIVES');
  });
});
