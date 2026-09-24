import { describe, it, expect, beforeEach } from 'vitest';
import { createBackup, restoreBackup } from '../src/services/storage/backupManager';
import { saveProviders, saveActiveProvider, getProviders, getActiveProvider } from '../src/services/storage/settingsStorage';
import { saveSoulConfig, getSoulConfig, clearSoulConfig } from '../src/services/storage/soulStorage';
import { saveCustomTool, getCustomTools, clearCustomTools } from '../src/services/storage/customToolStorage';
import { saveCrossAppSnapshot, listCrossAppSnapshots, clearCrossAppSnapshots } from '../src/services/storage/crossAppBridge';

describe('backupManager', () => {
  beforeEach(() => {
    clearSoulConfig();
    clearCustomTools();
    clearCrossAppSnapshots();
  });

  it('creates and restores unencrypted backup cleanly', async () => {
    saveSoulConfig({ corporateName: 'PT Nusantara Jaya', brandVoicePreset: 'financial_compliance' });
    saveCustomTool({ name: 'calc_discount', description: 'Hitung diskon', hostType: 'Excel', code: 'return 10;' });
    saveCrossAppSnapshot({ sourceHost: 'Excel', title: 'Snapshot Keuangan Q3', artifactType: 'table_data' });
    saveProviders([
      { id: 'openai', name: 'OpenAI', apiKey: 'sk-test-openai', selectedModel: 'gpt-4o', enabled: true },
    ]);
    saveActiveProvider('openai');

    const backupJson = await createBackup();
    const parsed = JSON.parse(backupJson);
    expect(parsed.format).toBe('sam-office-agent-backup');
    expect(parsed.encrypted).toBe(false);

    // Clear local state
    clearSoulConfig();
    clearCustomTools();
    clearCrossAppSnapshots();
    saveActiveProvider('gemini');
    expect(getSoulConfig().corporateName).not.toBe('PT Nusantara Jaya');

    // Restore
    const res = await restoreBackup(backupJson);
    expect(res.success).toBe(true);
    expect(getSoulConfig().corporateName).toBe('PT Nusantara Jaya');
    expect(getCustomTools().some(t => t.name === 'calc_discount')).toBe(true);
    expect(listCrossAppSnapshots().length).toBeGreaterThan(0);
    expect(getActiveProvider().id).toBe('openai');
    expect(getProviders().some(p => p.id === 'openai' && p.apiKey === 'sk-test-openai')).toBe(true);
  });

  it('creates and restores password-encrypted backup', async () => {
    saveSoulConfig({ corporateName: 'PT Enkripsi Indonesia' });
    const password = 'SuperSecretPassword2026';

    const backupJson = await createBackup({ password });
    const parsed = JSON.parse(backupJson);
    expect(parsed.encrypted).toBe(true);
    expect(parsed.salt).toBeDefined();
    expect(parsed.iv).toBeDefined();

    clearSoulConfig();

    // Wrong password fails
    await expect(restoreBackup(backupJson, { password: 'WrongPassword' })).rejects.toThrow();

    // Correct password succeeds
    const res = await restoreBackup(backupJson, { password });
    expect(res.success).toBe(true);
    expect(getSoulConfig().corporateName).toBe('PT Enkripsi Indonesia');
  });

  it('performs clean restore when cleanRestore option is true', async () => {
    saveSoulConfig({ corporateName: 'PT Awal' });
    const backupJson = await createBackup();

    saveCustomTool({ name: 'temporary_tool', description: 'Temp', hostType: 'Word', code: 'return 1;' });
    expect(getCustomTools().length).toBe(1);

    await restoreBackup(backupJson, { cleanRestore: true });
    // Temporary tool added after backup should be wiped by cleanRestore
    expect(getCustomTools().length).toBe(0);
  });
});
