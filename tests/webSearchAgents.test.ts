import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ExcelAgent } from '../src/agents/excel/excelAgent';
import { WordAgent } from '../src/agents/word/wordAgent';
import { PPTAgent } from '../src/agents/powerpoint/pptAgent';
import { MockOfficeDriver } from '../src/services/office/mockDriver';

describe('Universal web_search tool across specialists', () => {
  beforeEach(() => {
    vi.restoreAllMocks();
    const mockHtml = `
      <div class="result">
        <h2 class="result__title"><a class="result__a" href="https://bi.go.id">Bank Indonesia Kurs</a></h2>
        <a class="result__snippet">Kurs USD adalah Rp16.250.</a>
      </div>
    `;
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: true,
      text: async () => mockHtml,
    } as any);
  });

  it('provides web_search in ExcelAgent and returns search results with citations', async () => {
    const driver = new MockOfficeDriver();
    const agent = new ExcelAgent(driver);
    expect(agent.getTools().some(t => t.name === 'web_search')).toBe(true);

    const res = await agent.executeTool(
      { id: 'w1', name: 'web_search', arguments: { query: 'kurs dollar hari ini' }, status: 'pending' },
      { host: 'Excel' }
    );
    expect(res.success).toBe(true);
    expect(res.result).toContain('Bank Indonesia Kurs');
    expect(res.result).toContain('Rp16.250');
    expect(res.result).toContain('Sumber Referensi');
  });

  it('provides web_search in WordAgent', async () => {
    const driver = new MockOfficeDriver();
    const agent = new WordAgent(driver);
    expect(agent.getTools().some(t => t.name === 'web_search')).toBe(true);

    const res = await agent.executeTool(
      { id: 'w2', name: 'web_search', arguments: { query: 'tren HR 2026' }, status: 'pending' },
      { host: 'Word' }
    );
    expect(res.success).toBe(true);
    expect(res.result).toContain('Sumber Referensi');
  });

  it('provides web_search in PPTAgent', async () => {
    const driver = new MockOfficeDriver();
    const agent = new PPTAgent(driver);
    expect(agent.getTools().some(t => t.name === 'web_search')).toBe(true);

    const res = await agent.executeTool(
      { id: 'w3', name: 'web_search', arguments: { query: 'statistik e-commerce' }, status: 'pending' },
      { host: 'PowerPoint' }
    );
    expect(res.success).toBe(true);
    expect(res.result).toContain('Sumber Referensi');
  });
});
