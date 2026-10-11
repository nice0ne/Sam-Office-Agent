# Desain Arsitektur: Smart Mail Merge & Generator Dokumen Massal (Fase 3)

**Tanggal:** 2026-10-11  
**Status:** Disetujui (Approved)  
**Tujuan:** Menghadirkan mesin generator dokumen massal (*Smart Mail Merge Engine*) untuk Microsoft Word yang memanfaatkan data tabel dari Excel (melalui Sam Universal Hub / `crossAppBridge`), melakukan substitusi placeholder cerdas (`{{Nama}}`, `{{Nomor}}`, dll.), dan menyusun lembar dokumen per-penerima dengan pemisah halaman (*Page Break*) otomatis.

---

## 1. Latar Belakang & Ruang Lingkup

Dalam administrasi perkantoran harian, pembuatan surat massal (surat undangan, surat tugas, kuitansi, sertifikat, atau SPK) biasanya memakan waktu berjam-jam jika dilakukan manual salin-tempel.

Fase 3 menghadirkan modul `mailMergeEngine.ts` yang:
1. Mengekstrak variabel placeholder berformat `{{Nama}}`, `{{Alamat}}`, `{{Nominal}}`, dll. dari draf teks atau template dinas Word.
2. Mengambil dataset baris penerima dari Sam Universal Hub (`crossAppBridge`) yang dikirim dari Excel atau dari tabel input langsung.
3. Mencocokkan nama kolom tabel Excel ke tag template secara fleksibel (*case-insensitive* & toleran spasi).
4. Melakukan batch rendering di Word dengan menyisipkan *Page Break* antar-penerima sehingga dokumen siap langsung dicetak atau disimpan ke PDF.

---

## 2. Arsitektur Komponen

```
src/
├── services/
│   └── office/
│       ├── mailMergeEngine.ts            # [BARU] In-memory placeholder extractor & batch composer
│       ├── __tests__/
│       │   └── mailMergeEngine.test.ts   # [BARU] Unit tests mail merge engine
│       ├── wordDriver.ts                 # [UPDATE] Eksekusi executeMailMerge di Word.run
│       ├── mockDriver.ts                 # [UPDATE] Simulasi executeMailMerge di Mock Driver
│       └── types.ts                      # [UPDATE] Tipe MailMergeOptions & MailMergeResult
├── agents/
│   └── word/
│       └── wordAgent.ts                  # [UPDATE] Tool execute_mail_merge & ReAct prompt
└── components/
    └── Chat/
        └── QuickActionPresets.tsx        # [UPDATE] Preset Mail Merge di Word
```

---

## 3. Spesifikasi Rinci Antarmuka & Skema Data

### 3.1. Skema Tipe Data (`src/services/office/types.ts`)

```typescript
export interface MailMergeOptions {
  snapshotId?: string;
  tableData?: {
    headers: string[];
    rows: (string | number)[][];
  };
  templateText?: string;
  templateType?: 'OfficialLetter' | 'CustomText' | 'Kwitansi' | 'SPK';
  institutionName?: string;
  signatory?: {
    name: string;
    title: string;
  };
}

export interface MailMergeResult {
  totalRecordsProcessed: number;
  placeholdersDetected: string[];
  totalPagesCreated: number;
  recipientNames: string[];
  summaryMessage: string;
}
```

### 3.2. Fungsi Parser & Batch Composer (`src/services/office/mailMergeEngine.ts`)
- `extractPlaceholders(templateText: string): string[]`: Mengekstrak daftar tag placeholder unik seperti `['Nama', 'Nominal', 'Tanggal']`.
- `substitutePlaceholders(templateText: string, record: Record<string, any>): string`: Mengganti placeholder dengan nilai baris data.
- `mapTableRowsToRecords(headers: string[], rows: (string | number)[][]): Array<Record<string, any>>`: Mengonversi baris tabel 2D menjadi kumpulan objek kunci-nilai.
- `composeMailMergeBatch(options: MailMergeOptions, tableData: { headers: string[]; rows: (string | number)[][] }): { pages: string[]; result: MailMergeResult }`: Menghasilkan susunan teks halaman siap render di Word.

---

## 4. Penanganan Kesalahan (Error Handling)

1. **Snapshot Excel Tidak Ditemukan**: Jika Hub kosong dan tidak ada `tableData` yang diberikan, fungsi mengembalikan pesan panduan yang jelas kepada pengguna untuk membagikan data dari Excel terlebih dahulu.
2. **Placeholder Tidak Cocok dengan Kolom**: Jika ada tag `{{Jabatan}}` tetapi kolom tersebut tidak ada di Excel, parser secara aman mengisinya dengan strip `"-"` atau membiarkan tag asli tanpa melempar error.
3. **Baris Data Kosong**: Baris kosong otomatis disanitasi.

---

## 5. Rencana Pengujian (Testing Strategy)

1. **Unit Test (`mailMergeEngine.test.ts`)**:
   - Uji ekstraksi placeholder (`{{Nama}}`, `{{ nama }}`, `{{NAMA}}`).
   - Uji pemetaan baris tabel 2D ke objek record.
   - Uji substitusi template tunggal dan batch halaman.
2. **Integration Test (`mockDriver.test.ts` & `specialistAgents.test.ts`)**:
   - Uji eksekusi `executeMailMerge` di Mock Driver.
   - Uji eksekusi tool `execute_mail_merge` di `WordAgent` menggunakan snapshot dari `crossAppBridge`.
3. **UI Test (`chatAndActions.test.tsx`)**:
   - Uji keberadaan preset `✉️ Mail Merge (Dari Excel Hub)`.
4. **Regresi**:
   - 375+ tes yang sudah ada tetap 100% lulus.
