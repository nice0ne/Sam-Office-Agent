# Spec Desain: Excel Deep Intelligence & Safe Dry-Run Execution

**Tanggal:** 2026-09-22  
**Status:** Approved  
**Sub-Proyek:** 1 dari 4 (Excel Deep Intelligence & Safe Dry-Run)

---

## 1. Latar Belakang & Tujuan

Pengguna membutuhkan AI Assistant yang tidak hanya mengeksekusi instruksi dasar sel, tetapi bertindak sebagai **Expert Data Analyst & Auditor** dalam pekerjaan kantor sehari-hari. 

### Tujuan (Goals):
1. **Formula Auditor & Anomaly Detector (`audit_sheet_data`)**:
   - Mendeteksi error formula native Excel (`#REF!`, `#VALUE!`, `#N/A`, `#DIV/0!`, `#NAME?`, `#NUM!`).
   - Mendeteksi inkonsistensi formula per kolom (misal sel yang seharusnya formula namun diisi nilai manual/hardcoded).
   - Mendeteksi outlier statistik ekstrem dan sel kosong tak wajar pada baris data.
   - Memberikan laporan audit terstruktur dan saran perbaikan.
2. **Executive Data Storyteller (`generate_data_story`)**:
   - Menganalisis statistik rentang data (total, rata-rata, min, max, pertumbuhan, kontribusi persentase terbesar).
   - Menghasilkan ringkasan naratif eksekutif formal dalam bahasa Indonesia bisnis (temuan kunci, tren, peringatan risiko, dan rekomendasi).
3. **Safe Dry-Run & Action Confirmation Gate**:
   - Untuk modifikasi skala besar (>10 sel atau penghapusan data), agen menghasilkan proposal pratinjau perubahan (*action diff preview*).
   - Pengguna memiliki kendali penuh melalui tombol konfirmasi interaktif di UI chat (`Setujui & Terapkan` / `Batalkan`).
4. **Quick Action Presets Toolbar**:
   - Menambahkan preset Excel: `"🔍 Audit Formula & Error"` dan `"📈 Buat Analisis Tren"`.

### Batasan (Non-Goals):
- Sub-proyek ini difokuskan pada Excel dan Safe Execution; kapabilitas Word dan PowerPoint akan dilanjutkan pada Sub-Proyek 2 dan 3.

---

## 2. Model Data & Tipe Antarmuka (`src/services/office/types.ts`)

```typescript
export interface FormulaIssue {
  address: string;
  type: 'formula_error' | 'inconsistent_formula' | 'hardcoded_override' | 'suspicious_blank' | 'statistical_outlier';
  severity: 'critical' | 'warning' | 'info';
  currentValue?: any;
  formula?: string;
  expectedPattern?: string;
  suggestion: string;
}

export interface SheetAuditResult {
  sheetName: string;
  totalCellsAudited: number;
  totalErrorsFound: number;
  criticalIssues: FormulaIssue[];
  warnings: FormulaIssue[];
  summary: string;
}

export interface DataStoryOptions {
  range?: string;
  focusMetric?: string;
  includeRecommendations?: boolean;
}

export interface DataStoryMetric {
  label: string;
  value: string | number;
  changePercent?: number;
  trend?: 'up' | 'down' | 'neutral';
}

export interface DataStoryResult {
  headline: string;
  keyFindings: string[];
  metrics: DataStoryMetric[];
  risksOrAnomalies?: string[];
  recommendations: string[];
}

export interface PendingActionProposal {
  id: string;
  actionType: 'modify_cells' | 'clean_data' | 'format_cells';
  description: string;
  affectedCellsCount: number;
  targetRange?: string;
  payload: any;
}
```

Method baru pada interface `IOfficeDriver`:
```typescript
auditSheetData?(options?: { range?: string }): Promise<SheetAuditResult>;
generateDataStory?(options?: DataStoryOptions): Promise<DataStoryResult>;
```

---

## 3. Implementasi Driver

### 3.1. Excel Native Driver (`src/services/office/excelDriver.ts`)
- Memanfaatkan API `Excel.run` dengan single batch context sync (`context.sync()`).
- Mengambil properti `formulas`, `values`, dan `valueTypes` dari `usedRange` atau range spesifik.
- Memeriksa string value untuk mendeteksi error pattern: `/^#(REF!|VALUE!|DIV\/0!|N\/A|NAME\?|NUM!|NULL!)/i`.
- Memindai konsistensi formula kolom (apabila >80% kolom berisi formula serupa dengan formula R1C1, sel dengan nilai konstan akan ditandai sebagai `hardcoded_override`).
- Mengagregasi ringkasan statistik (mean, median, standard deviation sederhana untuk deteksi outlier).

### 3.2. Mock Driver (`src/services/office/mockDriver.ts`)
- Menyediakan implementasi mock penuh untuk `auditSheetData` dan `generateDataStory` yang mereplikasi kondisi data nyata sehingga 100% dapat diuji di vitest tanpa aplikasi Office nyata.

---

## 4. Registrasi Tool pada Specialist Agent (`src/agents/excel/excelAgent.ts`)

Menambahkan dua tool ke dalam `getTools()`:
1. `audit_sheet_data`:
   - Parameters: `{ range?: { type: 'string', description: 'Rentang opsional yang ingin diaudit, default seluruh sheet aktif' } }`.
2. `generate_data_story`:
   - Parameters: `{ range?: { type: 'string', description: 'Rentang data tabel' }, focusMetric?: { type: 'string', description: 'Metrik fokus analisis' }, includeRecommendations?: { type: 'boolean', description: 'Sertakan rekomendasi tindak lanjut' } }`.

---

## 5. Safe Dry-Run & Confirmation Card Flow

### Alur Kerja:
1. Jika suatu tool modifikasi (misal `write_cell` berskala >10 sel, atau `clean_data` yang menghapus baris) dipicu:
   - Tool dapat menghasilkan respons terstruktur dengan flag `requiresConfirmation: true` beserta proposal `pendingAction`.
2. UI Chat (`MessageBubble.tsx`) mendeteksi proposal tersebut dan me-render komponen `ActionConfirmationCard`:
   - Menampilkan ringkasan dampak (misal: "Akan memperbarui 35 formula di kolom D").
   - Menampilkan 2 tombol: **[Terapkan]** dan **[Batalkan]**.
3. Jika pengguna mengklik **Terapkan**:
   - Agen melanjutkan eksekusi final dan melaporkan hasil sukses.
4. Jika pengguna mengklik **Batalkan**:
   - Agen membatalkan operasi tanpa mengubah dokumen pengguna.

---

## 6. Quick Action Presets Toolbar (`src/components/Chat/QuickActionPresets.tsx`)

Menambahkan preset Excel baru:
- `🔍 Audit Formula & Error`: *"Audit seluruh formula pada sheet ini: cari error #REF!, #VALUE!, #DIV/0!, atau sel dengan rumus tidak konsisten dan berikan rekomendasinya."*
- `📈 Buat Analisis Tren`: *"Analisis data pada tabel aktif ini: buatkan ringkasan eksekutif, tren pertumbuhan, dan insight temuan utama."*

---

## 7. Rencana Pengujian (Testing Strategy)

- **Unit Tests Driver**:
  - `src/services/office/__tests__/mockDriver.test.ts`: verifikasi `auditSheetData` dan `generateDataStory`.
- **Unit Tests Specialist Agent**:
  - `src/agents/__tests__/specialistAgents.test.ts`: verifikasi registrasi tool baru dan penanganan eksekusi.
- **Unit Tests UI**:
  - `src/components/__tests__/chatAndActions.test.tsx`: verifikasi render kartu konfirmasi dan preset toolbar baru.
- **Target Keberhasilan**: Semua pengujian vitest (252+ tests) lulus dengan 0 kegagalan.
