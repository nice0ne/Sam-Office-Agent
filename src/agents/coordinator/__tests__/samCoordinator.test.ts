import { describe, it, expect } from 'vitest';
import { SamCoordinator } from '../samCoordinator';
import { saveSoulConfig, addLearnedDirective, clearSoulConfig } from '../../../services/storage/soulStorage';

describe('SamCoordinator', () => {
  it('selects correct specialist agent based on host context', () => {
    const coordinator = new SamCoordinator();
    expect(coordinator.getSpecialist('Excel').id).toBe('excel-specialist');
    expect(coordinator.getSpecialist('Word').id).toBe('word-specialist');
    expect(coordinator.getSpecialist('PowerPoint').id).toBe('ppt-specialist');
  });

  it('falls back to Excel specialist for BrowserDev and unknown hosts', () => {
    const coordinator = new SamCoordinator();
    expect(coordinator.getSpecialist('BrowserDev').id).toBe('excel-specialist');
    expect(coordinator.getSpecialist('Unknown' as any).id).toBe('excel-specialist');
  });

  it('builds system prompt tailored to active host', () => {
    const coordinator = new SamCoordinator();
    const prompt = coordinator.buildSystemPrompt('Excel', { host: 'Excel' });
    expect(prompt).toContain('Sam');
    expect(prompt).toContain('Excel');
    expect(prompt).toContain('Aturan Penting:');
  });

  it('builds system prompt tailored for Word and PowerPoint hosts', () => {
    const coordinator = new SamCoordinator();
    const wordPrompt = coordinator.buildSystemPrompt('Word', { host: 'Word' });
    expect(wordPrompt).toContain('Sam');
    expect(wordPrompt).toContain('Word');

    const pptPrompt = coordinator.buildSystemPrompt('PowerPoint', { host: 'PowerPoint' });
    expect(pptPrompt).toContain('Sam');
    expect(pptPrompt).toContain('PowerPoint');
  });

  it('dynamically injects active SOUL.md directives and learned rules into system prompt', () => {
    clearSoulConfig();
    const coordinator = new SamCoordinator();
    saveSoulConfig({
      enabled: true,
      corporateName: 'PT Telko Digital Nusantara',
      brandVoicePreset: 'formal_executive',
    });
    addLearnedDirective('Format seluruh angka dengan pemisah titik', 'formatting');

    const prompt = coordinator.buildSystemPrompt({ host: 'Word' });
    expect(prompt).toContain('CORPORATE BRAND VOICE & SOUL DIRECTIVES');
    expect(prompt).toContain('PT Telko Digital Nusantara');
    expect(prompt).toContain('Format seluruh angka dengan pemisah titik');

    // When disabled, SOUL directives should be excluded
    saveSoulConfig({ enabled: false });
    const disabledPrompt = coordinator.buildSystemPrompt({ host: 'Word' });
    expect(disabledPrompt).not.toContain('CORPORATE BRAND VOICE & SOUL DIRECTIVES');
  });
});
