import { describe, it, expect } from 'vitest';
import {
  extractPlaceholders,
  substitutePlaceholders,
  mapTableRowsToRecords,
  composeMailMergeBatch,
} from '../mailMergeEngine';

describe('Mail Merge Engine', () => {
  describe('extractPlaceholders', () => {
    it('extracts unique placeholder tags from template string', () => {
      const template = 'Kepada Yth. Sdr/i {{Nama}}\ndi {{Alamat}}.\nDengan ini kami mengundang Anda untuk pembayaran sebesar {{ Nominal }}.';
      const placeholders = extractPlaceholders(template);
      expect(placeholders).toEqual(['Nama', 'Alamat', 'Nominal']);
    });

    it('returns empty array if no placeholders present', () => {
      expect(extractPlaceholders('Surat tanpa placeholder.')).toEqual([]);
    });
  });

  describe('substitutePlaceholders', () => {
    it('substitutes placeholders case-insensitively with record data', () => {
      const template = 'Halo {{Nama}}, unit kerja {{Divisi}} dengan NIK {{nik}}.';
      const record = {
        nama: 'Ahmad Dahlan',
        divisi: 'Teknologi Informasi',
        NIK: '19870101',
      };
      const result = substitutePlaceholders(template, record);
      expect(result).toBe('Halo Ahmad Dahlan, unit kerja Teknologi Informasi dengan NIK 19870101.');
    });

    it('handles missing fields gracefully with dash fallback without error', () => {
      const template = 'Nama: {{Nama}}, Telepon: {{NoHP}}';
      const record = { Nama: 'Siti Aminah' };
      const result = substitutePlaceholders(template, record);
      expect(result).toBe('Nama: Siti Aminah, Telepon: -');
    });
  });

  describe('mapTableRowsToRecords', () => {
    it('converts 2D array rows and headers into structured record objects', () => {
      const headers = ['Nama', 'Jabatan', 'Nominal'];
      const rows = [
        ['Budi Santoso', 'Staff Admin', 3500000],
        ['Citra Dewi', 'Lead Engineer', 8000000],
      ];

      const records = mapTableRowsToRecords(headers, rows);
      expect(records.length).toBe(2);
      expect(records[0]).toEqual({
        Nama: 'Budi Santoso',
        Jabatan: 'Staff Admin',
        Nominal: 3500000,
      });
      expect(records[1]).toEqual({
        Nama: 'Citra Dewi',
        Jabatan: 'Lead Engineer',
        Nominal: 8000000,
      });
    });
  });

  describe('composeMailMergeBatch', () => {
    const tableData = {
      headers: ['Nama', 'Instansi', 'NomorSurat', 'Tanggal'],
      rows: [
        ['Dr. Hendra Wijaya', 'Universitas Indonesia', '001/UND/X/2026', '20 Oktober 2026'],
        ['Prof. Ratna Kusuma', 'Institut Teknologi Bandung', '002/UND/X/2026', '20 Oktober 2026'],
      ],
    };

    const templateText = 'SURAT UNDANGAN\nNomor: {{NomorSurat}}\n\nKepada Yth. {{Nama}}\n{{Instansi}}\n\nDengan hormat, kami mengundang Anda pada {{Tanggal}}.';

    it('composes individual personalized page documents and generates summary metrics', () => {
      const { pages, result } = composeMailMergeBatch(
        { templateText },
        tableData
      );

      expect(pages.length).toBe(2);
      expect(pages[0]).toContain('Dr. Hendra Wijaya');
      expect(pages[0]).toContain('001/UND/X/2026');
      expect(pages[1]).toContain('Prof. Ratna Kusuma');
      expect(pages[1]).toContain('002/UND/X/2026');

      expect(result.totalRecordsProcessed).toBe(2);
      expect(result.totalPagesCreated).toBe(2);
      expect(result.recipientNames).toContain('Dr. Hendra Wijaya');
      expect(result.recipientNames).toContain('Prof. Ratna Kusuma');
      expect(result.summaryMessage).toContain('Berhasil');
    });
  });
});
