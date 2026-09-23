import { CrossAppSnapshot, HostType } from '../../types';

const SNAPSHOTS_STORAGE_KEY = 'sam_cross_app_snapshots_v1';
const DISMISSED_STORAGE_KEY = 'sam_cross_app_dismissed_v1';
const DEFAULT_MAX_AGE_MS = 24 * 60 * 60 * 1000; // 24 hours

// Safe in-memory fallback for environments without localStorage (Node / SSR)
const memoryFallback = new Map<string, string>();

function getStorageItem(key: string): string | null {
  try {
    if (typeof localStorage !== 'undefined' && typeof localStorage.getItem === 'function') {
      return localStorage.getItem(key);
    }
  } catch (e) {
    console.warn(`[crossAppBridge] Failed to read localStorage key "${key}":`, e);
  }
  return memoryFallback.get(key) ?? null;
}

function setStorageItem(key: string, value: string): void {
  try {
    if (typeof localStorage !== 'undefined' && typeof localStorage.setItem === 'function') {
      localStorage.setItem(key, value);
      return;
    }
  } catch (e) {
    console.warn(`[crossAppBridge] Failed to write localStorage key "${key}":`, e);
  }
  memoryFallback.set(key, value);
}

function removeStorageItem(key: string): void {
  try {
    if (typeof localStorage !== 'undefined' && typeof localStorage.removeItem === 'function') {
      localStorage.removeItem(key);
      return;
    }
  } catch (e) {
    console.warn(`[crossAppBridge] Failed to remove localStorage key "${key}":`, e);
  }
  memoryFallback.delete(key);
}

function readAllSnapshots(): CrossAppSnapshot[] {
  const raw = getStorageItem(SNAPSHOTS_STORAGE_KEY);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function writeAllSnapshots(snapshots: CrossAppSnapshot[]): void {
  setStorageItem(SNAPSHOTS_STORAGE_KEY, JSON.stringify(snapshots));
}

function readDismissedIds(): string[] {
  const raw = getStorageItem(DISMISSED_STORAGE_KEY);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function writeDismissedIds(ids: string[]): void {
  setStorageItem(DISMISSED_STORAGE_KEY, JSON.stringify(ids));
}

/**
 * Saves a new cross-application snapshot payload to client-side storage.
 */
export function saveCrossAppSnapshot(
  snapshot: Omit<CrossAppSnapshot, 'id' | 'timestamp'>
): CrossAppSnapshot {
  const id = `snap_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`;
  const timestamp = Date.now();

  const newSnapshot: CrossAppSnapshot = {
    ...snapshot,
    id,
    timestamp,
  };

  const existing = readAllSnapshots();
  const updated = [newSnapshot, ...existing];
  writeAllSnapshots(updated);

  return newSnapshot;
}

/**
 * Lists unexpired cross-application snapshots ordered from newest to oldest.
 */
export function listCrossAppSnapshots(
  maxAgeMs: number = DEFAULT_MAX_AGE_MS
): CrossAppSnapshot[] {
  const now = Date.now();
  const all = readAllSnapshots();
  return all
    .filter(snap => now - snap.timestamp <= maxAgeMs)
    .sort((a, b) => b.timestamp - a.timestamp);
}

/**
 * Returns the most recent non-dismissed unexpired snapshot, optionally excluding the active host.
 */
export function getLatestCrossAppSnapshot(
  excludeHost?: HostType,
  maxAgeMs: number = DEFAULT_MAX_AGE_MS
): CrossAppSnapshot | null {
  const unexpired = listCrossAppSnapshots(maxAgeMs);
  const dismissedSet = new Set(readDismissedIds());

  for (const snap of unexpired) {
    if (dismissedSet.has(snap.id)) {
      continue;
    }
    if (excludeHost && snap.sourceHost === excludeHost) {
      continue;
    }
    return snap;
  }

  return null;
}

/**
 * Marks a snapshot ID as dismissed so it won't prompt the user again.
 */
export function dismissCrossAppSnapshot(id: string): void {
  const dismissed = readDismissedIds();
  if (!dismissed.includes(id)) {
    dismissed.push(id);
    writeDismissedIds(dismissed);
  }
}

/**
 * Checks whether a snapshot ID has been dismissed.
 */
export function isCrossAppSnapshotDismissed(id: string): boolean {
  const dismissed = readDismissedIds();
  return dismissed.includes(id);
}

/**
 * Clears all snapshots and dismissal records.
 */
export function clearCrossAppSnapshots(): void {
  removeStorageItem(SNAPSHOTS_STORAGE_KEY);
  removeStorageItem(DISMISSED_STORAGE_KEY);
  memoryFallback.clear();
}
