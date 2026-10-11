# Desain Arsitektur: Smart Reconciliation & Data Audit Engine (Fase 2)

**Tanggal:** 2026-10-11  
**Status:** Disetujui (Approved)  
**Tujuan:** Menghadirkan mesin rekonsiliasi data dan audit selisih otomatis (*Smart Reconciliation Diff Matcher*) untuk Microsoft Excel, membantu staf administrasi dan keuangan mencocokkan mutasi bank vs buku kas, invoice vs bukti bayar, atau komparasi inventaris dua arah dengan output lembar kerja laporan terstruktur.

---

## 1. Latar Belakang & Ruang Lingkup

Pada operasional harian kantor, pencocokan data antar dua tabel atau dua sheet (seperti mutasi rekening bank vs buku kas umum) sering dilakukan secara manual menggunakan formula VLOOKUP yang rentan error dan lambat.

Fase 2 menghadirkan modul `dataReconciler.ts` yang:
1. Mendukung perbandingan dua arah (*bidirectional diff*) antara dua range sel pada sheet yang sama atau sheet berbeda.
2. Mendeteksi kolom kunci (No. Ref / Invoice / ID) dan kolom nominal (Uang / Jumlah) secara otomatis atau berdasarkan parameter.
3. Mengelompokkan hasil ke dalam 4 kategori:
   - **Matched**: Ada di kedua sumber dengan nominal sama.
   - **Discrepancy**: Kunci ID cocok tetapi ada selisih nominal (menghitung nilai selisih).
   - **Unmatched in Source A**: Transaksi hanya tercatat di Sumber A.
   - **Unmatched in Source B**: Transaksi hanya tercatat di Sumber B.
4. Membuat lembar kerja baru `"Hasil_Rekonsiliasi"` lengkap dengan kartu ringkasan KPI eksekutif, tabel rincian selisih, format mata uang, dan perataan kolom otomatis.

---

## 2. Arsitektur Komponen

```
src/
├── services/
│   └── office/
│       ├── dataReconciler.ts             # [BARU] In-memory diff calculator & grid builder
│       ├── __tests__/
│       │   └── dataReconciler.test.ts    # [BARU] Unit tests reconciler engine
│       ├── excelDriver.ts                # [UPDATE] Eksekusi reconcileData di Excel.run
│       ├── mockDriver.ts                 # [UPDATE] Simulasi reconcileData di Mock Driver
│       └── types.ts                      # [UPDATE] Tipe ReconcileOptions & ReconcileSummary
├── agents/
│   └── excel/
│       └── excelAgent.ts                 # [UPDATE] Tool reconcile_data & ReAct prompt
└── components/
    └── Chat/
        └── QuickActionPresets.tsx        # [UPDATE] Preset Rekonsiliasi Data di Excel
```

---

## 3. Spesifikasi Rinci Antarmuka & Skema Data

### 3.1. Skema Tipe Data (`src/services/office/types.ts`)

```typescript
export interface ReconcileOptions {
  rangeA?: string;
  rangeB?: string;
  keyColumnA?: string | number;
  keyColumnB?: string | number;
  amountColumnA?: string | number;
  amountColumnB?: string | number;
  tolerance?: number;
  outputSheetName?: string;
}

export interface ReconcileItemDiscrepancy {
  key: string;
  descriptionA?: string;
  descriptionB?: string;
  amountA: number;
  amountB: number;
  diff: number;
}

export interface ReconcileItemUnmatched {
  key: string;
  description?: string;
  amount: number;
  source: 'A' | 'B';
}

export interface ReconcileSummary {
  totalRowsA: number;
  totalRowsB: number;
  matchedCount: number;
  matchedTotalAmount: number;
  discrepancyCount: number;
  netDiscrepancyAmount: number;
  unmatchedCountA: number;
  unmatchedCountB: number;
  discrepancies: ReconcileItemDiscrepancy[];
  unmatchedItems: ReconcileItemUnmatched[];
  markdownSummary: string;
}
```

### 3.2. Fungsi Parser & Perhitungan (`src/services/office/dataReconciler.ts`)
- `parseFlexibleNumber(val: any): number`: Mengonversi nilai numerik, teks mata uang Indonesia (`"Rp 1.500.000,00"`), format strip `"-"`, dan format akuntansi negatif `"(50.000)"` menjadi angka float.
- `detectKeyAndAmountColumns(headers: string[]): { keyIndex: number; amountIndex: number }`: Otomatis mengenali indeks kolom ID/Ref dan kolom nominal dari nama header.
- `reconcileDataSets(tableA: any[][], tableB: any[][], options?: ReconcileOptions): ReconcileSummary`: Menghasilkan analisis diff dua arah.
- `buildReconciliationSheetGrid(summary: ReconcileSummary): { summaryValues: any[][]; detailValues: any[][] }`: Menghasilkan array 2D yang siap ditulis ke Excel.

---

## 4. Penanganan Kesalahan (Error Handling)

1. **Range Tidak Ditemukan / Kosong**: Menghasilkan pesan peringatan ramah tanpa exception crash.
2. **Sheet Hasil Sudah Ada**: Driver otomatis membersihkan sheet yang ada atau membuat sheet bernomor urut berikutnya (`"Hasil_Rekonsiliasi_2"`).
3. **Baris Kosong di Data Sumber**: Baris kosong otomatis diabaikan (*sanitized*).

---

## 5. Rencana Pengujian (Testing Strategy)

1. **Unit Test (`dataReconciler.test.ts`)**:
   - Uji parser angka rupiah heterogen.
   - Uji deteksi kolom otomatis.
   - Uji bidirectional reconciliation (matched, discrepancy, unmatched).
   - Uji pembentukan grid sheet laporan.
2. **Integration Test (`specialistAgents.test.ts` & `mockDriver.test.ts`)**:
   - Uji pemanggilan tool `reconcile_data` pada `ExcelAgent`.
3. **UI Test (`chatAndActions.test.tsx`)**:
   - Memastikan preset rekonsiliasi muncul dan dapat diklik.
4. **Regresi Penuh**:
   - 363+ tes yang sudah ada tetap 100% lulus.
