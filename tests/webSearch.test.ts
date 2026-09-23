import { describe, it, expect, vi, beforeEach } from 'vitest';
import { searchWeb } from '../src/services/search';
import { getSearchSettings, setSearchSettings } from '../src/services/storage/settingsStorage';

describe('Web Search Service', () => {
  beforeEach(() => {
    localStorage.clear();
    vi.restoreAllMocks();
  });

  it('manages search settings with DuckDuckGo as default', () => {
    const settings = getSearchSettings();
    expect(settings.searchProvider).toBe('duckduckgo');

    setSearchSettings({ searchProvider: 'tavily', tavilyApiKey: 'tvly-test-123' });
    const updated = getSearchSettings();
    expect(updated.searchProvider).toBe('tavily');
    expect(updated.tavilyApiKey).toBe('tvly-test-123');
  });

  it('parses DuckDuckGo HTML results and formats clean citations', async () => {
    const mockHtml = `
      <div class="result web-result">
        <a class="result__url" href="//duckduckgo.com/l/?uddg=https%3A%2F%2Fwww.bi.go.id%2Fkurs">Bank Indonesia</a>
        <h2 class="result__title"><a class="result__a" href="//duckduckgo.com/l/?uddg=https%3A%2F%2Fwww.bi.go.id%2Fkurs">Informasi Kurs BI</a></h2>
        <a class="result__snippet">Kurs transaksi USD ke IDR hari ini adalah Rp16.200 per dolar.</a>
      </div>
    `;

    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: true,
      text: async () => mockHtml,
    } as any);

    const res = await searchWeb('kurs USD ke IDR');
    expect(res.provider).toBe('duckduckgo');
    expect(res.results.length).toBe(1);
    expect(res.results[0].title).toBe('Informasi Kurs BI');
    expect(res.results[0].url).toBe('https://www.bi.go.id/kurs');
    expect(res.results[0].snippet).toContain('Rp16.200');
    expect(res.citationsFormatted).toContain('https://www.bi.go.id/kurs');
  });

  it('falls back to DuckDuckGo when Tavily API key is missing or fails', async () => {
    setSearchSettings({ searchProvider: 'tavily', tavilyApiKey: '' });

    const mockHtml = `
      <div class="result">
        <h2 class="result__title"><a class="result__a" href="https://example.com">Fallback Result</a></h2>
        <a class="result__snippet">Contoh snippet fallback.</a>
      </div>
    `;

    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: true,
      text: async () => mockHtml,
    } as any);

    const res = await searchWeb('test query');
    expect(res.provider).toBe('duckduckgo');
    expect(res.results.length).toBeGreaterThan(0);
  });
});
