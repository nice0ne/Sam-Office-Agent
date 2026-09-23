import { SearchOptions, SearchResultItem } from './types';

export function cleanDdgUrl(rawUrl: string): string {
  if (!rawUrl) return '';
  let url = rawUrl.trim();
  if (url.startsWith('//')) {
    url = 'https:' + url;
  } else if (url.startsWith('/')) {
    url = 'https://duckduckgo.com' + url;
  }

  try {
    const parsed = new URL(url);
    const uddg = parsed.searchParams.get('uddg');
    if (uddg) {
      return decodeURIComponent(uddg);
    }
  } catch {
    const match = url.match(/[?&]uddg=([^&]+)/);
    if (match && match[1]) {
      return decodeURIComponent(match[1]);
    }
  }

  return url;
}

export function parseDuckDuckGoHtml(html: string): SearchResultItem[] {
  if (typeof DOMParser !== 'undefined') {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, 'text/html');
      const elements = doc.querySelectorAll('.result, .web-result');
      const items: SearchResultItem[] = [];
      elements.forEach((el) => {
        const linkEl = (el.querySelector('.result__a') ||
          el.querySelector('.result__title a') ||
          el.querySelector('a')) as HTMLAnchorElement | null;
        const snippetEl = el.querySelector('.result__snippet');
        if (!linkEl) return;
        const rawUrl = linkEl.getAttribute('href') || '';
        const title = linkEl.textContent?.trim() || '';
        const snippet = snippetEl?.textContent?.trim() || '';
        if (title && rawUrl) {
          items.push({
            title,
            url: cleanDdgUrl(rawUrl),
            snippet,
          });
        }
      });
      if (items.length > 0) return items;
    } catch {
      // Fall back to regex parsing
    }
  }

  // Regex fallback for environments without DOMParser (e.g. Node.js vitest environment)
  const items: SearchResultItem[] = [];
  const resultBlocks = html.split(/class=["'][^"']*\bresult\b/i).slice(1);
  for (const block of resultBlocks) {
    const linkMatch =
      block.match(/<a[^>]*class=["'][^"']*result__a[^"']*["'][^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/i) ||
      block.match(/<h2[^>]*class=["'][^"']*result__title[^"']*["'][^>]*>[\s\S]*?<a[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/i) ||
      block.match(/<a[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/i);
    const snippetMatch = block.match(/class=["'][^"']*result__snippet[^"']*["'][^>]*>([\s\S]*?)<\/(?:a|div|span|p)>/i);

    if (linkMatch) {
      const rawUrl = linkMatch[1];
      const title = linkMatch[2].replace(/<[^>]+>/g, '').trim();
      const snippet = snippetMatch ? snippetMatch[1].replace(/<[^>]+>/g, '').trim() : '';
      if (title && rawUrl) {
        items.push({
          title,
          url: cleanDdgUrl(rawUrl),
          snippet,
        });
      }
    }
  }
  return items;
}

export async function searchWithDuckDuckGo(
  query: string,
  options?: SearchOptions
): Promise<SearchResultItem[]> {
  try {
    const body = new URLSearchParams({ q: query, b: '' }).toString();
    const response = await fetch('https://html.duckduckgo.com/html/', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body,
    });

    if (!response.ok) {
      throw new Error(`DuckDuckGo search failed with status ${response.status}`);
    }

    const html = await response.text();
    const results = parseDuckDuckGoHtml(html);
    const maxResults = options?.maxResults ?? 5;
    return results.slice(0, maxResults);
  } catch (error) {
    console.warn('DuckDuckGo search error:', error);
    throw error;
  }
}
