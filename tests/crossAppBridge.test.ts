import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import {
  saveCrossAppSnapshot,
  getLatestCrossAppSnapshot,
  listCrossAppSnapshots,
  dismissCrossAppSnapshot,
  isCrossAppSnapshotDismissed,
  clearCrossAppSnapshots,
} from '../src/services/storage/crossAppBridge';

describe('CrossAppBridge Storage Engine', () => {
  beforeEach(() => {
    clearCrossAppSnapshots();
    if (typeof localStorage !== 'undefined') {
      localStorage.clear();
    }
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it('saves and retrieves latest cross-app snapshot excluding active host', () => {
    saveCrossAppSnapshot({
      sourceHost: 'Excel',
      title: 'Rekap Penjualan Q3',
      artifactType: 'table_data',
      tableData: {
        headers: ['Wilayah', 'Target', 'Realisasi'],
        rows: [['Barat', 1000, 1200]],
        totalRows: 1,
      },
      summaryText: 'Realisasi wilayah barat melampaui target 120%.',
    });

    // In Excel: excludeHost='Excel' should return null
    expect(getLatestCrossAppSnapshot('Excel')).toBeNull();

    // In Word: excludeHost='Word' should find Excel's snapshot
    const wordView = getLatestCrossAppSnapshot('Word');
    expect(wordView).not.toBeNull();
    expect(wordView?.title).toBe('Rekap Penjualan Q3');
    expect(wordView?.sourceHost).toBe('Excel');
  });

  it('supports dismissing a snapshot', () => {
    const snap = saveCrossAppSnapshot({
      sourceHost: 'Excel',
      title: 'Data Uji',
      artifactType: 'table_data',
    });

    expect(isCrossAppSnapshotDismissed(snap.id)).toBe(false);
    dismissCrossAppSnapshot(snap.id);
    expect(isCrossAppSnapshotDismissed(snap.id)).toBe(true);

    // Dismissed snapshot should no longer be returned as latest
    expect(getLatestCrossAppSnapshot('Word')).toBeNull();
  });

  it('filters out snapshots older than TTL (24 hours)', () => {
    vi.useFakeTimers();
    const initialTime = new Date('2026-09-23T10:00:00.000Z');
    vi.setSystemTime(initialTime);

    const snap = saveCrossAppSnapshot({
      sourceHost: 'Excel',
      title: 'Old Report',
      artifactType: 'table_data',
    });

    // Right after creation, it is retrievable
    expect(getLatestCrossAppSnapshot('Word')?.id).toBe(snap.id);
    expect(listCrossAppSnapshots().some(s => s.id === snap.id)).toBe(true);

    // Advance time by 25 hours (past 24h default TTL)
    vi.advanceTimersByTime(25 * 60 * 60 * 1000);

    // Should now be excluded
    expect(getLatestCrossAppSnapshot('Word')).toBeNull();
    expect(listCrossAppSnapshots().some(s => s.id === snap.id)).toBe(false);
  });

  it('lists all unexpired snapshots ordered by timestamp descending', () => {
    vi.useFakeTimers();
    const t1 = new Date('2026-09-23T10:00:00.000Z');
    vi.setSystemTime(t1);

    const snap1 = saveCrossAppSnapshot({
      sourceHost: 'Excel',
      title: 'Snapshot 1',
      artifactType: 'table_data',
    });

    vi.advanceTimersByTime(1000);
    const snap2 = saveCrossAppSnapshot({
      sourceHost: 'Word',
      title: 'Snapshot 2',
      artifactType: 'executive_summary',
    });

    const list = listCrossAppSnapshots();
    expect(list.length).toBe(2);
    expect(list[0].id).toBe(snap2.id);
    expect(list[1].id).toBe(snap1.id);
  });
});
