export interface SearchResultItem {
  title: string;
  url: string;
  snippet: string;
}

export interface SearchOptions {
  maxResults?: number;
  provider?: 'duckduckgo' | 'tavily';
  tavilyApiKey?: string;
}

export interface SearchResponse {
  query: string;
  provider: 'duckduckgo' | 'tavily';
  results: SearchResultItem[];
  citationsFormatted: string;
  error?: string;
}
