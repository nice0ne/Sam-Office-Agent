export type ThemeMode = 'auto' | 'light' | 'dark';

const STORAGE_KEY = 'sam_theme_mode';

/**
 * Calculates whether a given hex color code is considered dark
 * based on perceived luminance.
 */
export function isColorDark(hexOrRgb: string): boolean {
  if (!hexOrRgb) return false;
  let hex = hexOrRgb.trim();
  if (hex.startsWith('#')) {
    hex = hex.slice(1);
    if (hex.length === 3) {
      hex = hex.split('').map(c => c + c).join('');
    }
    if (hex.length === 6) {
      const r = parseInt(hex.substring(0, 2), 16);
      const g = parseInt(hex.substring(2, 4), 16);
      const b = parseInt(hex.substring(4, 6), 16);
      if (isNaN(r) || isNaN(g) || isNaN(b)) return false;
      const brightness = (r * 299 + g * 587 + b * 114) / 1000;
      return brightness < 128;
    }
  }
  return false;
}

export function getStoredThemeMode(): ThemeMode {
  try {
    const val = localStorage.getItem(STORAGE_KEY);
    if (val === 'light' || val === 'dark' || val === 'auto') {
      return val;
    }
  } catch {
    // Ignore storage read errors
  }
  return 'auto';
}

export function setStoredThemeMode(mode: ThemeMode): void {
  try {
    localStorage.setItem(STORAGE_KEY, mode);
  } catch (e) {
    console.warn('Gagal menyimpan tema ke localStorage:', e);
  }
}

/**
 * Determines whether dark mode should be active given the mode,
 * checking Office theme and system media queries if mode === 'auto'.
 */
export function shouldBeDark(mode: ThemeMode): boolean {
  if (mode === 'dark') return true;
  if (mode === 'light') return false;

  // Mode is 'auto': first check Office context theme
  if (typeof window !== 'undefined') {
    const office = (window as any).Office;
    const officeBg = office?.context?.officeTheme?.bodyBackgroundColor;
    if (officeBg) {
      return isColorDark(officeBg);
    }

    // Fallback to system prefers-color-scheme
    if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
      return true;
    }
  }

  return false;
}

/**
 * Applies the given theme mode by adding or removing the 'dark' class
 * on document.documentElement. Returns true if dark mode is active.
 */
export function applyTheme(mode: ThemeMode): boolean {
  const isDark = shouldBeDark(mode);
  if (typeof document !== 'undefined') {
    if (isDark) {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
  }
  return isDark;
}
