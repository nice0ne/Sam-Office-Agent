# Desain Arsitektur: Paket Tata Naskah Dinas & Administrasi Bawaan (Fase 1)

**Tanggal:** 2026-10-11  
**Status:** Disetujui (Approved)  
**Tujuan:** Memperkuat Sam Office Agent sebagai asisten administrasi harian dengan standardisasi tata naskah dinas, surat keluar resmi, berita acara serah terima (BAST), notulen rapat eksekutif (MoM), dan nota dinas internal berstandar formal Indonesia dan korporat modern.

---

## 1. Latar Belakang & Ruang Lingkup

Pengguna administrasi perkantoran (sekretaris, staf tata usaha, HR, operational admin, project manager) secara berkala membuat korespondensi dinas dan dokumen legalitas operasional. Sebelum pembaruan ini, pembuatan dokumen Word berfokus pada PRD, SOP umum, dan memo teks sederhana. 

Fase 1 ini menghadirkan generator naskah dinas standar Indonesia dengan tata letak baku:
1. **Surat Keluar / Surat Dinas Resmi (`OfficialLetter`)**: Lengkap dengan Kop Surat, Nomor Surat, Lampiran, Perihal, Tanggal Baku Indonesia, Yth. Penerima, Paragraf Isi, Blok Tanda Tangan, dan Tembusan.
2. **Berita Acara Serah Terima (`BAST`)**: Komparasi Pihak I & Pihak II, Klausul Legalitas Serah Terima, Tabel Rincian Objek/Pekerjaan, dan Blok Tanda Tangan Berdampingan (Pihak Pertama & Kedua).
3. **Notulen Rapat Eksekutif (`ExecutiveMoM`)**: Metadata Rapat, Agenda, Poin Pembahasan, dan Tabel Tindak Lanjut (*Action Items*) lengkap dengan kolom PIC, Tenggat Waktu (*Deadline*), dan Status.
4. **Nota Dinas / Memo Internal (`NotaDinas`)**: Standar komunikasi kedinasan antar-divisi dengan struktur Dari, Kepada, Tanggal, dan Perihal.

---

## 2. Arsitektur Komponen

```
src/
├── services/
│   └── office/
│       ├── administrativeDocTemplates.ts       # [BARU] Template builder & formatters
│       ├── __tests__/
│       │   └── administrativeDocTemplates.test.ts # [BARU] Unit tests builder & formatter
│       ├── wordDriver.ts                       # [UPDATE] Penanganan signature block & borderless layout
│       └── mockDriver.ts                       # [UPDATE] Simulasi driver untuk mock mode
├── agents/
│   └── word/
│       └── wordAgent.ts                        # [UPDATE] Schema tool generate_structured_doc & prompt
└── components/
    └── Chat/
        └── QuickActionPresets.tsx              # [UPDATE] Preset tombol naskah dinas Word
```

---

## 3. Spesifikasi Rinci Antarmuka & Skema Data

### 3.1. Skema Payload Naskah Dinas (`src/services/office/administrativeDocTemplates.ts`)

```typescript
export type AdministrativeDocType =
  | 'OfficialLetter'
  | 'BAST'
  | 'ExecutiveMoM'
  | 'NotaDinas';

export interface OfficialLetterPayload {
  institutionName: string;
  institutionAddress?: string;
  letterNumber: string;
  date: string;
  attachments?: string;
  subject: string;
  recipient: {
    title: string;
    address: string;
  };
  openingText: string;
  bodyParagraphs: string[];
  closingText: string;
  signatory: {
    name: string;
    title: string;
    nipOrNik?: string;
  };
  copiesTo?: string[];
}

export interface BastPayload {
  documentNumber: string;
  date: string;
  firstParty: {
    name: string;
    title: string;
    company: string;
  };
  secondParty: {
    name: string;
    title: string;
    company: string;
  };
  handoverItems: {
    headers: string[];
    rows: string[][];
  };
  clauses: string[];
}

export interface ExecutiveMomPayload {
  meetingTitle: string;
  date: string;
  timeAndPlace: string;
  attendees: string[];
  agenda: string[];
  discussionSummary: string[];
  actionItems: Array<{
    task: string;
    pic: string;
    deadline: string;
    status: string;
  }>;
}

export interface InternalMemoPayload {
  memoNumber: string;
  from: string;
  to: string;
  date: string;
  subject: string;
  bodyParagraphs: string[];
  signatory: {
    name: string;
    title: string;
  };
}
```

### 3.2. Fungsi Pembantu Format Baku Indonesia
- `formatIndonesianDate(inputDate?: string): string`: Mengonversi input tanggal seperti `"2026-10-20"` atau `"20/10/2026"` menjadi `"20 Oktober 2026"`.
- `formatOfficialNumber(num: string, dept: string, monthRoman: string, year: number): string`: Menghasilkan string nomor surat resmi standar seperti `"042/DIR-OPS/X/2026"`.

### 3.3. Penataan Tata Letak (Word Driver)
- **Tanda Tangan Berdampingan (Dual Signature)**: Dibuat menggunakan tabel 2 kolom tanpa batas (*borderless table*) dengan lebar 50%-50% untuk menampung Pihak Pertama di kolom kiri dan Pihak Kedua di kolom kanan, sehingga format dijamin rapi dan sejajar.
- **Tabel Rincian & Action Items**: Tabel bergaris formal dengan baris header berwarna latar abu-abu/navy muda dan teks tebal.

---

## 4. Penanganan Kesalahan (Error Handling)

1. **Input Tidak Lengkap**: Jika pengguna tidak mencantumkan nomor surat, nama pihak, atau tanggal, generator otomatis mengisi dengan penanda terformat (contoh: `[Nomor Surat: .../DIR/X/2026]`, `[Nama Penandatangan]`) tanpa melempar exception atau membatalkan proses.
2. **Konteks Dokumen Word**: Jika fungsi dipanggil di luar lingkungan Word aktif, driver beralih secara aman ke `mockDriver` (pada browser) atau menghasilkan pesan peringatan yang ramah pengguna.
3. **Penyelamatan Kesalahan Runtime**: Setiap pemanggilan API Word dibungkus dalam blok try-catch yang mengembalikan status kegagalan terperinci ke agen ReAct agar agen dapat memberikan respons penjelasan.

---

## 5. Rencana Pengujian (Testing Strategy)

1. **Unit Testing (`administrativeDocTemplates.test.ts`)**:
   - Uji pemformatan nama bulan baku bahasa Indonesia (Januari s.d. Desember).
   - Uji pembangunan struktur `OfficialLetter` (memastikan keberadaan kop, nomor, lampiran, perihal, isi, dan tanda tangan).
   - Uji pembangunan `BAST` (memastikan tabel serah terima dan identitas kedua belah pihak lengkap).
   - Uji pembangunan `ExecutiveMoM` (memastikan tabel action items lengkap dengan header PIC dan Deadline).
   - Uji pembangunan `NotaDinas`.
2. **Integration Testing (`specialistAgents.test.ts`)**:
   - Memastikan `WordAgent` dapat mengenali dan mengeksekusi tool `generate_structured_doc` untuk tipe `OfficialLetter`, `BAST`, `ExecutiveMoM`, dan `NotaDinas`.
3. **Regression Suite**:
   - Menjalankan seluruh pengujian (350+ tes) untuk menjamin stabilitas fungsionalitas yang ada.
