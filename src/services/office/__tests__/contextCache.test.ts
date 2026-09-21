import { describe, it, expect, beforeEach, vi } from 'vitest';
import { DocumentContextCache, documentContextCache } from '../contextCache';
import { AgentContext } from '../../../agents/types';

describe('DocumentContextCache', () => {
  let cache: DocumentContextCache;

  beforeEach(() => {
    cache = new DocumentContextCache(5000); // 5s TTL for tests
    vi.useRealTimers();
  });

  it('stores and retrieves context within TTL using get and set', () => {
    const excelContext: AgentContext = {
      host: 'Excel',
      documentSummary: 'Sheet1 Data with 10 rows',
      activeCellOrRange: 'A1:B10',
    };

    cache.set('Excel', excelContext);
    const hit = cache.get('Excel');
    expect(hit).not.toBeNull();
    expect(hit?.documentSummary).toBe('Sheet1 Data with 10 rows');
    expect(hit?.activeCellOrRange).toBe('A1:B10');
  });

  it('supports getCached and setCached aliases', () => {
    const wordContext: AgentContext = {
      host: 'Word',
      documentSummary: 'Sample Document',
      selectedText: 'Hello World',
    };

    cache.setCached('Word', wordContext);
    const hit = cache.getCached('Word');
    expect(hit).not.toBeNull();
    expect(hit?.documentSummary).toBe('Sample Document');
  });

  it('returns null for nonexistent keys', () => {
    expect(cache.get('PowerPoint')).toBeNull();
    expect(cache.getCached('PowerPoint')).toBeNull();
  });

  it('invalidates cache for a specific host', () => {
    cache.set('Excel', { host: 'Excel', documentSummary: 'Excel Doc' });
    cache.set('Word', { host: 'Word', documentSummary: 'Word Doc' });

    cache.invalidate('Excel');

    expect(cache.get('Excel')).toBeNull();
    expect(cache.get('Word')).not.toBeNull();
    expect(cache.get('Word')?.documentSummary).toBe('Word Doc');
  });

  it('invalidates cache for all hosts when no host argument is passed', () => {
    cache.set('Excel', { host: 'Excel', documentSummary: 'Excel Doc' });
    cache.set('Word', { host: 'Word', documentSummary: 'Word Doc' });
    cache.set('PowerPoint', { host: 'PowerPoint', documentSummary: 'Deck Doc' });

    cache.invalidate();

    expect(cache.get('Excel')).toBeNull();
    expect(cache.get('Word')).toBeNull();
    expect(cache.get('PowerPoint')).toBeNull();
  });

  it('returns null when cached entry has expired beyond TTL', async () => {
    const shortCache = new DocumentContextCache(20); // 20ms TTL
    shortCache.set('Word', { host: 'Word', documentSummary: 'Doc to expire' });

    expect(shortCache.get('Word')).not.toBeNull();

    await new Promise((resolve) => setTimeout(resolve, 30));

    expect(shortCache.get('Word')).toBeNull();
  });

  it('defaults to 15000ms TTL when not specified', () => {
    const defaultCache = new DocumentContextCache();
    expect(defaultCache.getTtlMs()).toBe(15000);
  });

  it('exports a singleton instance documentContextCache', () => {
    expect(documentContextCache).toBeInstanceOf(DocumentContextCache);
    documentContextCache.invalidate();

    const pptContext: AgentContext = { host: 'PowerPoint', documentSummary: 'Slide 1' };
    documentContextCache.set('PowerPoint', pptContext);
    expect(documentContextCache.get('PowerPoint')?.documentSummary).toBe('Slide 1');

    documentContextCache.invalidate('PowerPoint');
    expect(documentContextCache.get('PowerPoint')).toBeNull();
  });
});
