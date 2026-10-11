# Smart Reconciliation & Data Audit Engine (Fase 2) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Mengimplementasikan mesin pencocokan dan audit selisih data otomatis (*Smart Reconciliation Diff Matcher*) untuk Excel, memungkinkan staf administrasi/keuangan merekonsiliasi mutasi kas/bank atau dua tabel data dengan pembuatan lembar kerja laporan terstruktur.

**Architecture:** Modul terisolasi `dataReconciler.ts` memproses array 2D dari dua sumber data (Tabel A dan Tabel B), menormalisasi ID kunci dan nominal uang rupiah, melakukan pencocokan dua arah (*matched, discrepancy, unmatched*), dan menyusun layout lembar kerja `"Hasil_Rekonsiliasi"`. Driver `excelDriver.ts` mengeksekusi pembuatan sheet dan formatting, sedangkan `ExcelAgent` memfasilitasi perintah dialog ReAct.

**Tech Stack:** TypeScript 5.6, React 18.3, Office.js (Excel), Vitest 3.2.

**Spec:** `docs/superpowers/specs/2026-10-11-smart-reconciliation-excel-design.md`

## Global Constraints
- 100% Client-Side BYOK: tidak ada dependensi eksternal atau server relay.
- Kompatibilitas Office.js Excel: mendukung rentang sel tunggal atau antar-sheet.
- Menjaga 363 tes yang sudah ada tetap 100% lulus (zero regression).

---

### Task 1: Core Reconciliation Calculation Engine & Grid Builder

**Files:**
- Create: `src/services/office/dataReconciler.ts`
- Create: `src/services/office/__tests__/dataReconciler.test.ts`
- Modify: `src/services/office/types.ts`
- Modify: `src/services/office/index.ts`

**Interfaces:**
- Produces:
  - `parseFlexibleNumber(val: any): number`
  - `detectKeyAndAmountColumns(headers: string[]): { keyIndex: number; amountIndex: number }`
  - `reconcileDataSets(tableA: any[][], tableB: any[][], options?: ReconcileOptions): ReconcileSummary`
  - `buildReconciliationSheetGrid(summary: ReconcileSummary): { summaryRows: any[][]; detailRows: any[][] }`

- [ ] **Step 1: Write failing unit tests for dataReconciler**

Create `src/services/office/__tests__/dataReconciler.test.ts` testing:
1. `parseFlexibleNumber` handles plain numbers, `"Rp 1.500.000,00"`, `"(25.000)"`, `"-"`, null/undefined.
2. `detectKeyAndAmountColumns` detects `"No. Ref"` and `"Nominal"` columns.
3. `reconcileDataSets` detects matched rows, discrepancy rows with net diff, unmatched rows in A, and unmatched rows in B.
4. `buildReconciliationSheetGrid` builds executive KPI grid and discrepancy detail rows.

- [ ] **Step 2: Run test to confirm failure**

Run: `npx vitest run src/services/office/__tests__/dataReconciler.test.ts`

- [ ] **Step 3: Implement `dataReconciler.ts` and update `types.ts` & `index.ts`**

Write `dataReconciler.ts` implementing all calculation, detection, diffing, and grid builder logic.

- [ ] **Step 4: Run test to verify passes**

Run: `npx vitest run src/services/office/__tests__/dataReconciler.test.ts`

- [ ] **Step 5: Commit changes**

Git commit: `feat(office): implement core smart data reconciliation engine and grid builder`

---

### Task 2: Excel & Mock Driver Integration for Reconciliation

**Files:**
- Modify: `src/services/office/excelDriver.ts`
- Modify: `src/services/office/mockDriver.ts`
- Test: `src/services/office/__tests__/mockDriver.test.ts`

**Interfaces:**
- Consumes: `reconcileDataSets` and `buildReconciliationSheetGrid` from Task 1
- Produces: `reconcileData(options: ReconcileOptions): Promise<ReconcileSummary>` on `IDocumentDriver`

- [ ] **Step 1: Write test in `mockDriver.test.ts` for `reconcileData`**

Add tests verifying that `mockDriver.reconcileData` reads the tables, invokes reconciler, generates `"Hasil_Rekonsiliasi"`, and returns the summary metrics.

- [ ] **Step 2: Run test to confirm failure**

Run: `npx vitest run src/services/office/__tests__/mockDriver.test.ts`

- [ ] **Step 3: Implement `reconcileData` in `mockDriver.ts` and `excelDriver.ts`**

Update `mockDriver.ts` and `excelDriver.ts` to implement `reconcileData`. In `excelDriver.ts`, ensure worksheet creation, writing summary grid, writing detail rows, setting cell colors, and applying autofit.

- [ ] **Step 4: Run tests to confirm passes**

Run: `npx vitest run src/services/office/__tests__/mockDriver.test.ts`

- [ ] **Step 5: Commit changes**

Git commit: `feat(office): integrate reconciliation execution in Excel and mock drivers`

---

### Task 3: ExcelAgent Integration & System Prompt Enhancements

**Files:**
- Modify: `src/agents/excel/excelAgent.ts`
- Modify: `src/agents/__tests__/specialistAgents.test.ts`

**Interfaces:**
- Consumes: `reconcileData` from `src/services/office`
- Produces: `reconcile_data` tool in `ExcelAgent`

- [ ] **Step 1: Write failing test in `specialistAgents.test.ts`**

Add test in `src/agents/__tests__/specialistAgents.test.ts` asserting that `ExcelAgent` exposes and executes `reconcile_data`.

- [ ] **Step 2: Run test to confirm failure**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`

- [ ] **Step 3: Update `excelAgent.ts`**

1. Update `getSystemPrompt` to guide users on bank reconciliation and table comparison.
2. Add `reconcile_data` tool definition to `getTools()`.
3. Handle execution of `reconcile_data` in `executeTool()`.

- [ ] **Step 4: Run test to confirm passes**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`

- [ ] **Step 5: Commit changes**

Git commit: `feat(agents): add reconcile_data tool to ExcelAgent`

---

### Task 4: UI Quick Action Preset for Excel Reconciliation

**Files:**
- Modify: `src/components/Chat/QuickActionPresets.tsx`
- Modify: `src/components/__tests__/chatAndActions.test.tsx`

**Interfaces:**
- Produces: `📊 Rekonsiliasi Data & Kas` preset button for Excel host

- [ ] **Step 1: Write test for new Excel preset**

Update `src/components/__tests__/chatAndActions.test.tsx` to verify presence of `📊 Rekonsiliasi Data & Kas`.

- [ ] **Step 2: Run test to confirm failure**

Run: `npx vitest run src/components/__tests__/chatAndActions.test.tsx`

- [ ] **Step 3: Add preset to `QuickActionPresets.tsx`**

Add `excel_reconcile` to `PRESETS_BY_HOST.Excel`.

- [ ] **Step 4: Run test to confirm passes**

Run: `npx vitest run src/components/__tests__/chatAndActions.test.tsx`

- [ ] **Step 5: Commit changes**

Git commit: `feat(ui): add reconciliation quick action preset for Excel`

---

### Task 5: Full Regression Testing & Virtual Office Milestone Update

**Files:**
- Full test suite: `npx vitest run`

- [ ] **Step 1: Run complete test suite**

Ensure all test files pass with 0 regressions.

- [ ] **Step 2: Broadcast completion to Virtual Office 3D**

Execute `node E:/VIBE-CODE-WS/Sam-Virtual-Office/notify.js` to update status on the 3D Whiteboard.
