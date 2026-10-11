import { describe, it, expect } from 'vitest';
import {
  parseFlexibleNumber,
  detectKeyAndAmountColumns,
  reconcileDataSets,
  buildReconciliationSheetGrid,
} from '../dataReconciler';

describe('Data Reconciler Engine', () => {
  describe('parseFlexibleNumber', () => {
    it('parses standard numbers and numeric strings', () => {
      expect(parseFlexibleNumber(1250000)).toBe(1250000);
      expect(parseFlexibleNumber('1250000')).toBe(1250000);
      expect(parseFlexibleNumber('1250000.50')).toBe(1250000.5);
    });

    it('parses Indonesian currency format strings with Rp and periods', () => {
      expect(parseFlexibleNumber('Rp 1.500.000,00')).toBe(1500000);
      expect(parseFlexibleNumber('Rp. 2.750.000')).toBe(2750000);
      expect(parseFlexibleNumber('IDR 500.000,50')).toBe(500000.5);
    });

    it('parses accounting negative parentheses format', () => {
      expect(parseFlexibleNumber('(50.000)')).toBe(-50000);
      expect(parseFlexibleNumber('(Rp 100.000,00)')).toBe(-100000);
      expect(parseFlexibleNumber('-25000')).toBe(-25000);
    });

    it('handles dashes, blanks, and null safely', () => {
      expect(parseFlexibleNumber('-')).toBe(0);
      expect(parseFlexibleNumber('')).toBe(0);
      expect(parseFlexibleNumber(null)).toBe(0);
      expect(parseFlexibleNumber(undefined)).toBe(0);
      expect(parseFlexibleNumber('N/A')).toBe(0);
    });
  });

  describe('detectKeyAndAmountColumns', () => {
    it('automatically identifies key and amount column indices from headers', () => {
      const headers = ['Tanggal', 'No. Referensi', 'Keterangan', 'Nominal Masuk'];
      const { keyIndex, amountIndex } = detectKeyAndAmountColumns(headers);
      expect(keyIndex).toBe(1); // 'No. Referensi'
      expect(amountIndex).toBe(3); // 'Nominal Masuk'
    });

    it('falls back to invoice/id and total/jumlah', () => {
      const headers = ['ID Transaksi', 'Pelanggan', 'Jumlah'];
      const { keyIndex, amountIndex } = detectKeyAndAmountColumns(headers);
      expect(keyIndex).toBe(0);
      expect(amountIndex).toBe(2);
    });

    it('defaults to index 0 and index 1 when no matching keywords are found', () => {
      const headers = ['Kolom A', 'Kolom B'];
      const { keyIndex, amountIndex } = detectKeyAndAmountColumns(headers);
      expect(keyIndex).toBe(0);
      expect(amountIndex).toBe(1);
    });
  });

  describe('reconcileDataSets', () => {
    const tableBank = [
      ['No. Ref', 'Keterangan', 'Debit / Mutasi'],
      ['TRX-001', 'Setoran Modal PT ABC', 'Rp 10.000.000'],
      ['TRX-002', 'Pembayaran Biaya Server', 'Rp 2.500.000'],
      ['TRX-003', 'Biaya Administrasi Bank', 'Rp 25.000'],
      ['TRX-004', 'Penerimaan Piutang Client X', 'Rp 5.000.000'],
    ];

    const tableKasInternal = [
      ['Kode Bukti', 'Deskripsi Kas', 'Jumlah Keluar/Masuk'],
      ['TRX-001', 'Setoran Modal', '10000000'],
      ['TRX-002', 'Biaya Server AWS', '2450000'], // Selisih 50.000
      ['TRX-004', 'Pelunasan Client X', '5000000'],
      ['TRX-005', 'Beli Perlengkapan ATK Kantor', '150000'], // Hanya di internal
    ];

    it('performs bidirectional reconciliation detecting matched, discrepancy, and unmatched items', () => {
      const summary = reconcileDataSets(tableBank, tableKasInternal);

      expect(summary.totalRowsA).toBe(4);
      expect(summary.totalRowsB).toBe(4);

      // Matched: TRX-001 (10jt) and TRX-004 (5jt)
      expect(summary.matchedCount).toBe(2);
      expect(summary.matchedTotalAmount).toBe(15000000);

      // Discrepancy: TRX-002 (Bank 2.500.000 vs Internal 2.450.000 -> Diff 50.000)
      expect(summary.discrepancyCount).toBe(1);
      expect(summary.discrepancies.length).toBe(1);
      expect(summary.discrepancies[0].key).toBe('TRX-002');
      expect(summary.discrepancies[0].amountA).toBe(2500000);
      expect(summary.discrepancies[0].amountB).toBe(2450000);
      expect(summary.discrepancies[0].diff).toBe(50000);

      // Unmatched A (Hanya di Bank): TRX-003 (Biaya Admin Bank 25.000)
      expect(summary.unmatchedCountA).toBe(1);
      const unmatchedA = summary.unmatchedItems.filter(i => i.source === 'A');
      expect(unmatchedA.length).toBe(1);
      expect(unmatchedA[0].key).toBe('TRX-003');

      // Unmatched B (Hanya di Kas Internal): TRX-005 (Beli ATK 150.000)
      expect(summary.unmatchedCountB).toBe(1);
      const unmatchedB = summary.unmatchedItems.filter(i => i.source === 'B');
      expect(unmatchedB.length).toBe(1);
      expect(unmatchedB[0].key).toBe('TRX-005');

      // Markdown summary should be generated
      expect(summary.markdownSummary).toContain('Laporan Rekonsiliasi Data');
      expect(summary.markdownSummary).toContain('TRX-002');
    });

    it('supports tolerance for rounding differences', () => {
      const tA = [
        ['ID', 'Nilai'],
        ['INV-01', '1000.50'],
      ];
      const tB = [
        ['ID', 'Nilai'],
        ['INV-01', '1000.00'],
      ];

      // Without tolerance -> discrepancy
      const resStrict = reconcileDataSets(tA, tB, { tolerance: 0 });
      expect(resStrict.discrepancyCount).toBe(1);
      expect(resStrict.matchedCount).toBe(0);

      // With tolerance of 1.0 -> matched
      const resTolerant = reconcileDataSets(tA, tB, { tolerance: 1.0 });
      expect(resTolerant.discrepancyCount).toBe(0);
      expect(resTolerant.matchedCount).toBe(1);
    });
  });

  describe('buildReconciliationSheetGrid', () => {
    it('generates 2D arrays for KPI summary and detail tables', () => {
      const mockSummary = {
        totalRowsA: 10,
        totalRowsB: 10,
        matchedCount: 8,
        matchedTotalAmount: 80000000,
        discrepancyCount: 1,
        netDiscrepancyAmount: 500000,
        unmatchedCountA: 1,
        unmatchedCountB: 0,
        discrepancies: [
          { key: 'INV-10', descriptionA: 'Service', descriptionB: 'Service', amountA: 1500000, amountB: 1000000, diff: 500000 },
        ],
        unmatchedItems: [
          { key: 'INV-11', description: 'Pending Bank', amount: 200000, source: 'A' as const },
        ],
        markdownSummary: 'Summary',
      };

      const grid = buildReconciliationSheetGrid(mockSummary);
      expect(grid.summaryRows.length).toBeGreaterThan(3);
      expect(grid.detailRows.length).toBeGreaterThan(2);

      // Headers of detail table
      const headerRow = grid.detailRows[0];
      expect(headerRow).toContain('Kunci / No. Ref');
      expect(headerRow).toContain('Selisih (Diff)');
      expect(headerRow).toContain('Status');
    });
  });
});
