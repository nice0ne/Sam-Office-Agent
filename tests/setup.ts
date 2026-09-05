// Test setup configuration
import { beforeEach } from 'vitest';

class LocalStorageMock {
  private store: Record<string, string> = {};

  clear() {
    this.store = {};
  }

  getItem(key: string): string | null {
    return this.store[key] ?? null;
  }

  setItem(key: string, value: string) {
    this.store[key] = String(value);
  }

  removeItem(key: string) {
    delete this.store[key];
  }

  get length(): number {
    return Object.keys(this.store).length;
  }

  key(index: number): string | null {
    return Object.keys(this.store)[index] ?? null;
  }
}

if (typeof globalThis.localStorage === 'undefined' || typeof globalThis.localStorage?.clear !== 'function') {
  Object.defineProperty(globalThis, 'localStorage', {
    value: new LocalStorageMock(),
    writable: true,
    configurable: true,
  });
}

beforeEach(() => {
  // Clear localStorage before each test
  if (typeof localStorage !== 'undefined' && typeof localStorage.clear === 'function') {
    localStorage.clear();
  }
});

