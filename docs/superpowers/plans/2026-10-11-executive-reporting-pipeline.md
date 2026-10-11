# Implementation Plan: 1-Click Executive Reporting Pipeline (Excel ➔ Word ➔ PPT)

**Branch:** `feat/executive-reporting-pipeline`  
**Spec:** `docs/superpowers/specs/2026-10-11-executive-reporting-pipeline-design.md`  

---

### Task Breakdown

- [ ] **Task 1: Core Executive Pipeline Service (`executivePipeline.ts`)**
  - Buat `src/services/office/executivePipeline.ts`:
    - Mengolah data tabel / metrik.
    - Membentuk narasi dokumen laporan eksekutif.
    - Menghasilkan slide deck presentasi (menggunakan `synthesizeDocToDeck`).
    - Menyimpan snapshot cross-app ke Universal Hub (`saveCrossAppSnapshot`).
  - Update `src/services/office/types.ts` & `src/services/office/index.ts`.
  - Tulis unit test `src/services/office/__tests__/executivePipeline.test.ts`.
  - Verifikasi vitest & commit.

- [ ] **Task 2: Agent Tools Integration (`execute_executive_pipeline`)**
  - Tambahkan tool `execute_executive_pipeline` pada `src/agents/coordinator/samCoordinator.ts`.
  - Dukung pemanggilan lintas agen di `wordAgent.ts` & `excelAgent.ts`.
  - Tulis test integrasi di `samCoordinator.test.ts` & `specialistAgents.test.ts`.
  - Verifikasi vitest & commit.

- [ ] **Task 3: UI Quick Action Presets Integration**
  - Tambahkan preset `universal_executive_pipeline` di `QuickActionPresets.tsx`.
  - Tampilkan di Excel, Word, dan PowerPoint.
  - Update test di `chatAndActions.test.tsx` & commit.

- [ ] **Task 4: Full Regression & Merge to Main**
  - Jalankan full regression `npx vitest run`.
  - Notifikasi ke Sam Virtual Office via `notify.js`.
  - Merge `feat/executive-reporting-pipeline` ke `main`.
