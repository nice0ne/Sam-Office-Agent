import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import {
  isColorDark,
  getStoredThemeMode,
  setStoredThemeMode,
  applyTheme,
} from '../theme';

describe('Theme Utility', () => {
  const classListSet = new Set<string>();
  const mockDocument = {
    documentElement: {
      classList: {
        add: (cls: string) => classListSet.add(cls),
        remove: (cls: string) => classListSet.delete(cls),
        contains: (cls: string) => classListSet.has(cls),
      },
    },
  };

  beforeEach(() => {
    localStorage.clear();
    classListSet.clear();
    if (typeof (globalThis as any).window === 'undefined') {
      (globalThis as any).window = globalThis;
    }
    (globalThis as any).document = mockDocument;
    delete (window as any).Office;
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  describe('isColorDark', () => {
    it('correctly identifies dark hex colors', () => {
      expect(isColorDark('#000000')).toBe(true);
      expect(isColorDark('#262626')).toBe(true); // Office Black theme
      expect(isColorDark('#2b2b2b')).toBe(true); // Office Dark Gray
      expect(isColorDark('#1e1e1e')).toBe(true);
      expect(isColorDark('#000')).toBe(true);
    });

    it('correctly identifies light hex colors', () => {
      expect(isColorDark('#ffffff')).toBe(false);
      expect(isColorDark('#f3f4f6')).toBe(false);
      expect(isColorDark('#fff')).toBe(false);
      expect(isColorDark('#ffffff')).toBe(false);
    });

    it('handles empty or invalid color strings safely', () => {
      expect(isColorDark('')).toBe(false);
      expect(isColorDark('invalid')).toBe(false);
    });
  });

  describe('Theme Mode Storage', () => {
    it('defaults to auto when no theme is stored', () => {
      expect(getStoredThemeMode()).toBe('auto');
    });

    it('stores and retrieves theme mode correctly', () => {
      setStoredThemeMode('dark');
      expect(getStoredThemeMode()).toBe('dark');

      setStoredThemeMode('light');
      expect(getStoredThemeMode()).toBe('light');

      setStoredThemeMode('auto');
      expect(getStoredThemeMode()).toBe('auto');
    });
  });

  describe('applyTheme', () => {
    it('forces dark class when mode is dark', () => {
      const isDark = applyTheme('dark');
      expect(isDark).toBe(true);
      expect(document.documentElement.classList.contains('dark')).toBe(true);
    });

    it('removes dark class when mode is light', () => {
      document.documentElement.classList.add('dark');
      const isDark = applyTheme('light');
      expect(isDark).toBe(false);
      expect(document.documentElement.classList.contains('dark')).toBe(false);
    });

    it('applies dark theme in auto mode when Office context theme is dark', () => {
      (window as any).Office = {
        context: {
          officeTheme: {
            bodyBackgroundColor: '#262626',
          },
        },
      };
      const isDark = applyTheme('auto');
      expect(isDark).toBe(true);
      expect(document.documentElement.classList.contains('dark')).toBe(true);
    });

    it('applies light theme in auto mode when Office context theme is light', () => {
      (window as any).Office = {
        context: {
          officeTheme: {
            bodyBackgroundColor: '#ffffff',
          },
        },
      };
      const isDark = applyTheme('auto');
      expect(isDark).toBe(false);
      expect(document.documentElement.classList.contains('dark')).toBe(false);
    });

    it('falls back to prefers-color-scheme when Office theme is not present', () => {
      window.matchMedia = vi.fn().mockImplementation(query => ({
        matches: query.includes('dark'),
        media: query,
        onchange: null,
        addListener: vi.fn(),
        removeListener: vi.fn(),
        addEventListener: vi.fn(),
        removeEventListener: vi.fn(),
        dispatchEvent: vi.fn(),
      }));

      const isDark = applyTheme('auto');
      expect(isDark).toBe(true);
      expect(document.documentElement.classList.contains('dark')).toBe(true);
    });
  });
});
