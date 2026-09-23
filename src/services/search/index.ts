import { getSearchSettings } from '../storage/settingsStorage';
import { searchWithDuckDuckGo } from './duckduckgo';
import { searchWithTavily } from './tavily';
import { SearchOptions, SearchResponse, SearchResultItem } from './types';

export function formatCitations(results: SearchResultItem[]): string {
  if (!results || results.length === 0) return '';
  const lines = ['🌐 **Sumber Referensi:**'];
  results.forEach((r, idx) => {
    lines.push(`${idx + 1}. [${r.title || r.url}](${r.url})`);
  });
  return lines.join('\n');
}

export async function searchWeb(
  query: string,
  options?: SearchOptions
): Promise<SearchResponse> {
  const settings = getSearchSettings();
  const requestedProvider = options?.provider || settings.searchProvider || 'duckduckgo';
  const tavilyApiKey =
    (options?.tavilyApiKey !== undefined ? options.tavilyApiKey : settings.tavilyApiKey) || '';
  const maxResults = options?.maxResults ?? 5;

  let results: SearchResultItem[] = [];
  let resolvedProvider: 'duckduckgo' | 'tavily' = 'duckduckgo';
  let error: string | undefined;

  if (requestedProvider === 'tavily' && tavilyApiKey.trim()) {
    try {
      results = await searchWithTavily(query, tavilyApiKey, maxResults);
      resolvedProvider = 'tavily';
    } catch (err: any) {
      console.warn('Tavily search failed, falling back to DuckDuckGo:', err?.message || err);
      try {
        results = await searchWithDuckDuckGo(query, { ...options, maxResults });
        resolvedProvider = 'duckduckgo';
      } catch (ddgErr: any) {
        error = ddgErr?.message || String(ddgErr);
      }
    }
  } else {
    try {
      results = await searchWithDuckDuckGo(query, { ...options, maxResults });
      resolvedProvider = 'duckduckgo';
    } catch (err: any) {
      error = err?.message || String(err);
    }
  }

  const citationsFormatted = formatCitations(results);

  return {
    query,
    provider: resolvedProvider,
    results,
    citationsFormatted,
    ...(error ? { error } : {}),
  };
}

export * from './types';
export * from './duckduckgo';
export * from './tavily';
