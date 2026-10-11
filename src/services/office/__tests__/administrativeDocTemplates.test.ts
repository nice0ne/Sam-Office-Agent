import { describe, it, expect } from 'vitest';
import {
  formatIndonesianDate,
  formatOfficialNumber,
  toRomanMonth,
  buildAdministrativeDoc,
  OfficialLetterPayload,
  BastPayload,
  ExecutiveMomPayload,
  InternalMemoPayload,
} from '../administrativeDocTemplates';

describe('Administrative Document Templates & Formatters', () => {
  describe('formatIndonesianDate', () => {
    it('formats ISO date strings correctly to Indonesian format', () => {
      expect(formatIndonesianDate('2026-10-20')).toBe('20 Oktober 2026');
      expect(formatIndonesianDate('2026-01-05')).toBe('5 Januari 2026');
      expect(formatIndonesianDate('2026-08-17')).toBe('17 Agustus 2026');
    });

    it('formats Date instances correctly', () => {
      const d = new Date(2026, 4, 2); // 2 Mei 2026
      expect(formatIndonesianDate(d)).toBe('2 Mei 2026');
    });

    it('handles Indonesian relative or text dates cleanly', () => {
      expect(formatIndonesianDate('11 Oktober 2026')).toBe('11 Oktober 2026');
      // If empty or undefined, returns today's Indonesian date with year
      const todayFormatted = formatIndonesianDate();
      expect(todayFormatted).toMatch(/\d{1,2}\s+[A-Z][a-z]+\s+\d{4}/);
    });
  });

  describe('formatOfficialNumber & toRomanMonth', () => {
    it('converts month numbers to roman numerals', () => {
      expect(toRomanMonth(1)).toBe('I');
      expect(toRomanMonth(4)).toBe('IV');
      expect(toRomanMonth(10)).toBe('X');
      expect(toRomanMonth(12)).toBe('XII');
    });

    it('formats official letter number string hierarchy', () => {
      expect(formatOfficialNumber('042', 'DIR-OPS', 'X', 2026)).toBe('042/DIR-OPS/X/2026');
    });
  });

  describe('buildAdministrativeDoc - OfficialLetter', () => {
    it('constructs a complete official letter structure with kop, metadata, and signatory', () => {
      const payload: OfficialLetterPayload = {
        institutionName: 'PT INOVASI TEKNOLOGI NUSANTARA',
        institutionAddress: 'Gedung Cyber 2 Lt. 15, Jl. HR Rasuna Said, Jakarta Selatan',
        letterNumber: '015/DIR-OPS/X/2026',
        date: '2026-10-20',
        subject: 'Undangan Rapat Koordinasi Teknis Proyek',
        recipient: {
          title: 'Kepala Dinas Komunikasi dan Informatika',
          address: 'di Tempat',
        },
        openingText: 'Dengan hormat,',
        bodyParagraphs: [
          'Sehubungan dengan implementasi sistem otomasi naskah dinas terpadu tahun anggaran 2026, kami mengundang Bapak/Ibu untuk hadir pada rapat koordinasi teknis.',
          'Rapat akan diselenggarakan pada hari Kamis, 22 Oktober 2026 bertempat di Ruang Rapat Utama.',
        ],
        closingText: 'Demikian surat undangan ini kami sampaikan. Atas perhatian dan kerja sama Bapak/Ibu, kami ucapkan terima kasih.',
        signatory: {
          name: 'Bambang Trihatmodjo, S.T., M.Kom.',
          title: 'Direktur Utama',
          nipOrNik: '19780512 200212 1 002',
        },
        copiesTo: ['Dewan Komisaris', 'Arsip Operasional'],
      };

      const doc = buildAdministrativeDoc('OfficialLetter', payload);
      expect(doc.templateType).toBe('OfficialLetter');
      expect(doc.title).toContain('PT INOVASI TEKNOLOGI NUSANTARA');
      expect(doc.sections.length).toBeGreaterThan(3);

      const headerSection = doc.sections.find(s => s.heading.includes('KOP SURAT') || s.heading.includes('LEMBAR'));
      expect(headerSection).toBeDefined();

      const metaSection = doc.sections.find(s => s.heading.includes('NOMOR'));
      expect(metaSection?.content).toContain('015/DIR-OPS/X/2026');
      expect(metaSection?.content).toContain('20 Oktober 2026');

      const sigSection = doc.sections.find(s => s.heading.includes('TANDA TANGAN'));
      expect(sigSection?.content).toContain('Bambang Trihatmodjo');
      expect(sigSection?.content).toContain('Direktur Utama');
    });
  });

  describe('buildAdministrativeDoc - BAST', () => {
    it('constructs a BAST with dual parties, handover items table, and dual signatures', () => {
      const payload: BastPayload = {
        documentNumber: '007/BAST-IT/X/2026',
        date: '2026-10-20',
        firstParty: {
          name: 'Agus Setiawan',
          title: 'Project Lead',
          company: 'PT Vendor Solusi Cipta',
        },
        secondParty: {
          name: 'Siti Rahmawati',
          title: 'Manager IT & Sarpras',
          company: 'PT Mandiri Global Corp',
        },
        handoverItems: {
          headers: ['No', 'Nama Barang / Pekerjaan', 'Jumlah', 'Satuan', 'Keterangan'],
          rows: [
            ['1', 'Server Dell PowerEdge R750', '2', 'Unit', 'Kondisi Baru & Berfungsi Normal'],
            ['2', 'Dokumen Lisensi Microsoft 365 Enterprise', '100', 'Seat', 'Aktif 1 Tahun'],
          ],
        },
        clauses: [
          'PIHAK PERTAMA telah menyerahkan seluruh pekerjaan dan barang di atas kepada PIHAK KEDUA dalam kondisi baik.',
          'PIHAK KEDUA telah memeriksa dan menerima penyerahan tersebut dengan penuh kepuasan.',
        ],
      };

      const doc = buildAdministrativeDoc('BAST', payload);
      expect(doc.templateType).toBe('BAST');
      expect(doc.title).toContain('BERITA ACARA SERAH TERIMA');

      const tableSection = doc.sections.find(s => s.table !== undefined);
      expect(tableSection?.table?.rows.length).toBe(2);
      expect(tableSection?.table?.headers).toContain('Nama Barang / Pekerjaan');

      const sigSection = doc.sections.find(s => s.heading.includes('PENANDATANGAN'));
      expect(sigSection?.table?.headers).toContain('PIHAK PERTAMA (Menyerahkan)');
      expect(sigSection?.table?.headers).toContain('PIHAK KEDUA (Menerima)');
    });
  });

  describe('buildAdministrativeDoc - ExecutiveMoM', () => {
    it('constructs an Executive MoM with structured action items table with PIC and deadline', () => {
      const payload: ExecutiveMomPayload = {
        meetingTitle: 'Rapat Evaluasi Implementasi AI Office Assistant',
        date: '2026-10-20',
        timeAndPlace: '10:00 - 12:00 WIB | Ruang Rapat Eksekutif',
        attendees: ['Budi (PM)', 'Dewi (Lead Engineer)', 'Fajar (HR Lead)'],
        agenda: ['Review Timeline', 'Alokasi Budget Q4', 'Uji Coba Mail Merge'],
        discussionSummary: [
          'Fase 1 template administrasi dinas telah diselesaikan 100%.',
          'Persiapan untuk Fase 2 modul rekonsiliasi data Excel dijadwalkan minggu depan.',
        ],
        actionItems: [
          { task: 'Finalisasi integrasi template Word', pic: 'Dewi', deadline: '22 Okt 2026', status: 'In Progress' },
          { task: 'Penyusunan panduan user manual dinas', pic: 'Budi', deadline: '25 Okt 2026', status: 'Pending' },
        ],
      };

      const doc = buildAdministrativeDoc('ExecutiveMoM', payload);
      expect(doc.templateType).toBe('ExecutiveMoM');
      expect(doc.title).toContain('Rapat Evaluasi');

      const actionItemSection = doc.sections.find(s => s.heading.includes('TINDAK LANJUT') || s.heading.includes('ACTION ITEMS'));
      expect(actionItemSection?.table?.headers).toContain('PIC');
      expect(actionItemSection?.table?.headers).toContain('Deadline');
      expect(actionItemSection?.table?.rows.length).toBe(2);
    });
  });

  describe('buildAdministrativeDoc - NotaDinas', () => {
    it('constructs a formal internal memo with sender, recipient, and subject', () => {
      const payload: InternalMemoPayload = {
        memoNumber: '028/ND/HRD/X/2026',
        from: 'Kepala Divisi Sumber Daya Manusia',
        to: 'Seluruh Kepala Departemen',
        date: '2026-10-20',
        subject: 'Sosialisasi Standarisasi Format Surat dan Dokumen Dinas',
        bodyParagraphs: [
          'Dalam rangka meningkatkan efisiensi dan keseragaman tata naskah dinas, diwajibkan seluruh departemen menggunakan template baku Sam Office Agent.',
          'Implementasi efektif berlaku mulai tanggal 1 November 2026.',
        ],
        signatory: {
          name: 'Hj. Ratna Sari, M.Psi.',
          title: 'Kepala Divisi SDM',
        },
      };

      const doc = buildAdministrativeDoc('NotaDinas', payload);
      expect(doc.templateType).toBe('NotaDinas');
      expect(doc.title).toContain('NOTA DINAS');
      const metaSection = doc.sections.find(s => s.heading.includes('LEMBAR NOTA DINAS'));
      expect(metaSection?.content).toContain('Kepala Divisi Sumber Daya Manusia');
      expect(metaSection?.content).toContain('Seluruh Kepala Departemen');
    });
  });
});
