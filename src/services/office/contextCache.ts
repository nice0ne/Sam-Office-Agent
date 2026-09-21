import { AgentContext } from '../../agents/types';
import { HostType } from '../../types';

interface CacheEntry {
  context: AgentContext;
  timestamp: number;
}

/**
 * Cache for storing document context snapshots per Office Host (Excel, Word, PowerPoint).
 * Reduces latency by avoiding redundant Office.js context retrieval within a given TTL window.
 */
export class DocumentContextCache {
  private cache = new Map<HostType, CacheEntry>();
  private ttlMs: number;

  constructor(ttlMs = 15000) {
    this.ttlMs = ttlMs;
  }

  /**
   * Retrieves the cached snapshot for the given host if present and not expired.
   */
  get(host: HostType): AgentContext | null {
    const entry = this.cache.get(host);
    if (!entry) {
      return null;
    }

    if (Date.now() - entry.timestamp > this.ttlMs) {
      this.cache.delete(host);
      return null;
    }

    return entry.context;
  }

  /**
   * Alias for `get`.
   */
  getCached(host: HostType): AgentContext | null {
    return this.get(host);
  }

  /**
   * Stores a context snapshot for the given host.
   */
  set(host: HostType, context: AgentContext): void {
    this.cache.set(host, {
      context,
      timestamp: Date.now(),
    });
  }

  /**
   * Alias for `set`.
   */
  setCached(host: HostType, context: AgentContext): void {
    this.set(host, context);
  }

  /**
   * Checks whether an unexpired context snapshot exists for the given host.
   */
  has(host: HostType): boolean {
    return this.get(host) !== null;
  }

  /**
   * Invalidates cached context for a specific host, or all hosts if omitted.
   */
  invalidate(host?: HostType): void {
    if (host) {
      this.cache.delete(host);
    } else {
      this.cache.clear();
    }
  }

  /**
   * Gets the configured TTL in milliseconds.
   */
  getTtlMs(): number {
    return this.ttlMs;
  }
}

/**
 * Global singleton instance for document context caching.
 */
export const documentContextCache = new DocumentContextCache();
