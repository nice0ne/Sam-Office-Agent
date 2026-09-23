import { SearchResultItem } from './types';

interface TavilyRawResult {
  title: string;
  url: string;
  content: string;
  score?: number;
}

interface TavilySearchResponse {
  results?: TavilyRawResult[];
}

export async function searchWithTavily(
  query: string,
  apiKey: string,
  maxResults: number = 5
): Promise<SearchResultItem[]> {
  if (!apiKey || !apiKey.trim()) {
    throw new Error('Tavily API key is missing');
  }

  const response = await fetch('https://api.tavily.com/search', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      api_key: apiKey.trim(),
      query,
      max_results: maxResults,
      search_depth: 'basic',
    }),
  });

  if (!response.ok) {
    const errorText = await response.text().catch(() => '');
    throw new Error(`Tavily search failed (${response.status}): ${errorText}`);
  }

  const data = (await response.json()) as TavilySearchResponse;
  const rawResults = data.results || [];

  return rawResults.map((item) => ({
    title: item.title || '',
    url: item.url || '',
    snippet: item.content || '',
  }));
}
