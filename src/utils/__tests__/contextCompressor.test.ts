import { describe, it, expect } from 'vitest';
import { compressTableContext } from '../contextCompressor';

describe('compressTableContext', () => {
  it('samples first rows and provides row count summary', () => {
    const mockData = [
      ['ID', 'Nama', 'Sales'],
      ['1', 'Andi', 100],
      ['2', 'Budi', 200],
      ['3', 'Cici', 300],
      ['4', 'Dedi', 400],
      ['5', 'Eka', 500],
      ['6', 'Fani', 600],
    ];

    const result = compressTableContext(mockData, 3);
    expect(result.headers).toEqual(['ID', 'Nama', 'Sales']);
    expect(result.sampledRows.length).toBe(3);
    expect(result.totalRows).toBe(6);
    expect(result.summaryText).toContain('Total baris: 6');
  });

  it('handles empty data correctly', () => {
    const result = compressTableContext([]);
    expect(result.headers).toEqual([]);
    expect(result.sampledRows).toEqual([]);
    expect(result.totalRows).toBe(0);
    expect(result.totalColumns).toBe(0);
    expect(result.summaryText).toBe('Tabel kosong.');
  });

  it('defaults to sampling 5 rows when maxSampleRows is not provided', () => {
    const mockData = [
      ['Col1'],
      ['Row1'],
      ['Row2'],
      ['Row3'],
      ['Row4'],
      ['Row5'],
      ['Row6'],
      ['Row7'],
    ];
    const result = compressTableContext(mockData);
    expect(result.sampledRows.length).toBe(5);
    expect(result.totalRows).toBe(7);
  });
});
