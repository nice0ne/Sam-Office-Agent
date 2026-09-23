import { describe, it, expect, beforeEach } from 'vitest';
import { ExcelAgent } from '../src/agents/excel/excelAgent';
import { WordAgent } from '../src/agents/word/wordAgent';
import { PPTAgent } from '../src/agents/powerpoint/pptAgent';
import { MockOfficeDriver } from '../src/services/office/mockDriver';
import { clearSoulConfig, getSoulConfig } from '../src/services/storage/soulStorage';

describe('Self-Learning SOUL Directives Tooling', () => {
  beforeEach(() => {
    clearSoulConfig();
  });

  it('provides learn_corporate_directive tool in all specialist agents', () => {
    const mockDriver = new MockOfficeDriver();
    const excelAgent = new ExcelAgent(mockDriver);
    const wordAgent = new WordAgent(mockDriver);
    const pptAgent = new PPTAgent(mockDriver);

    expect(excelAgent.getTools().some(t => t.name === 'learn_corporate_directive')).toBe(true);
    expect(wordAgent.getTools().some(t => t.name === 'learn_corporate_directive')).toBe(true);
    expect(pptAgent.getTools().some(t => t.name === 'learn_corporate_directive')).toBe(true);
  });

  it('executes learn_corporate_directive and persists learned rule', async () => {
    const mockDriver = new MockOfficeDriver();
    const wordAgent = new WordAgent(mockDriver);

    const res = await wordAgent.executeTool(
      {
        id: 'learn-1',
        name: 'learn_corporate_directive',
        arguments: {
          rule: 'Gunakan istilah mitra strategis untuk klien prioritas',
          category: 'terminology',
          explanation: 'Koreksi penulisan dokumen kontrak',
        },
        status: 'pending',
      },
      { host: 'Word' }
    );

    expect(res.success).toBe(true);
    expect(res.result).toContain('Berhasil mempelajari direktif korporat');
    expect(res.result).toContain('mitra strategis');

    const config = getSoulConfig();
    expect(config.learnedDirectives.length).toBe(1);
    expect(config.learnedDirectives[0].rule).toContain('mitra strategis');
    expect(config.learnedDirectives[0].category).toBe('terminology');
  });
});
