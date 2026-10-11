# Tata Naskah Dinas & Administrasi Bawaan (Fase 1) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Menghadirkan generator template naskah dinas baku (Surat Keluar Resmi, BAST, Notulen MoM dengan Action Items, dan Nota Dinas) serta preset tombol satu-klik pada Sam Office Agent untuk Microsoft Word.

**Architecture:** Modul terisolasi `administrativeDocTemplates.ts` bertindak sebagai builder yang memvalidasi input, menormalisasi tanggal dan penomoran resmi Indonesia, lalu mengonversi payload naskah dinas menjadi struktur section Word (termasuk dual-signature table borderless) yang dieksekusi oleh `wordDriver.ts`. Agen `WordAgent` diperkaya dengan skema tool baru dan panduan tata naskah dinas formal vs korporat modern.

**Tech Stack:** TypeScript 5.6, React 18.3, Office.js (Word), Vitest 3.2.

**Spec:** `docs/superpowers/specs/2026-10-11-tata-naskah-dinas-administrasi-design.md`

## Global Constraints
- 100% Client-Side BYOK: tidak ada dependensi eksternal atau server relay pihak ketiga.
- Kompatibilitas Office.js: aman dijalankan di desktop Word dan mock browser mode.
- Menjaga 350 tes yang sudah ada tetap 100% lulus (zero regression).

---

### Task 1: Administrative Template Builder & Indonesian Formatters

**Files:**
- Create: `src/services/office/administrativeDocTemplates.ts`
- Create: `src/services/office/__tests__/administrativeDocTemplates.test.ts`
- Modify: `src/services/office/index.ts`

**Interfaces:**
- Produces:
  - `formatIndonesianDate(inputDate?: string): string`
  - `formatOfficialNumber(num: string, dept: string, monthRoman: string, year: number): string`
  - `buildAdministrativeDoc(type: AdministrativeDocType, payload: any): StructuredDocDefinition`

- [ ] **Step 1: Write the failing unit tests for formatters and builders**

Create `src/services/office/__tests__/administrativeDocTemplates.test.ts` testing:
1. `formatIndonesianDate` correctly formats ISO date, slashes, or Indonesian relative dates.
2. `formatOfficialNumber` formats hierarchical department and roman month.
3. `buildAdministrativeDoc('OfficialLetter', ...)` produces kop surat, metadata, paragraphs, dual/single signature, and copiesTo.
4. `buildAdministrativeDoc('BAST', ...)` produces dual parties, handover items table, and dual signatures.
5. `buildAdministrativeDoc('ExecutiveMoM', ...)` produces meeting metadata and action items table with PIC & Deadline.
6. `buildAdministrativeDoc('NotaDinas', ...)` produces memo headers and body.

- [ ] **Step 2: Run test to confirm failure**

Run: `npx vitest run src/services/office/__tests__/administrativeDocTemplates.test.ts`

- [ ] **Step 3: Implement `administrativeDocTemplates.ts`**

Write `src/services/office/administrativeDocTemplates.ts` with type definitions, date helper (Januari - Desember), roman month mapper, and structured section builders for `OfficialLetter`, `BAST`, `ExecutiveMoM`, and `NotaDinas`. Export the builder and formatters in `src/services/office/index.ts`.

- [ ] **Step 4: Run test to verify passes**

Run: `npx vitest run src/services/office/__tests__/administrativeDocTemplates.test.ts`

- [ ] **Step 5: Commit changes**

Git commit: `feat(office): implement administrative document templates and Indonesian formatters`

---

### Task 2: Word & Mock Driver Dual-Signature Borderless Layout

**Files:**
- Modify: `src/services/office/wordDriver.ts`
- Modify: `src/services/office/mockDriver.ts`
- Modify: `src/services/office/types.ts`
- Test: `src/services/office/__tests__/mockDriver.test.ts`

**Interfaces:**
- Consumes: `StructuredDocDefinition` from Task 1
- Produces: `insertStructuredDocument(doc: StructuredDocDefinition): Promise<void>` supporting signature blocks and borderless dual tables.

- [ ] **Step 1: Write test for dual-signature and structured doc rendering in mock driver**

Update `src/services/office/__tests__/mockDriver.test.ts` to assert that tables flagged as borderless or signature blocks are tracked and recorded accurately in mock state.

- [ ] **Step 2: Run mock driver tests to confirm failure or coverage gap**

Run: `npx vitest run src/services/office/__tests__/mockDriver.test.ts`

- [ ] **Step 3: Update `types.ts`, `wordDriver.ts`, and `mockDriver.ts`**

Enhance `StructuredDocSection` and table options to support `borderless?: boolean` or signature alignment, ensuring that in Word.run, dual signatures are rendered cleanly side-by-side without visible borders.

- [ ] **Step 4: Run tests to confirm passes**

Run: `npx vitest run src/services/office/__tests__/mockDriver.test.ts`

- [ ] **Step 5: Commit changes**

Git commit: `feat(office): support borderless signature blocks in Word and mock drivers`

---

### Task 3: WordAgent Integration & System Prompt Enhancements

**Files:**
- Modify: `src/agents/word/wordAgent.ts`
- Modify: `src/agents/__tests__/specialistAgents.test.ts`

**Interfaces:**
- Consumes: `buildAdministrativeDoc` from `src/services/office`
- Produces: Updated `generate_structured_doc` tool accepting new template types (`OfficialLetter`, `BAST`, `ExecutiveMoM`, `NotaDinas`) and administrative parameters.

- [ ] **Step 1: Write failing integration test in `specialistAgents.test.ts`**

Add tests to `src/agents/__tests__/specialistAgents.test.ts`:
- WordAgent executes `generate_structured_doc` with `templateType: 'OfficialLetter'`.
- WordAgent executes `generate_structured_doc` with `templateType: 'BAST'`.
- WordAgent generates executive MoM with action items table.

- [ ] **Step 2: Run test to confirm failure**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`

- [ ] **Step 3: Update `wordAgent.ts`**

1. Update `getSystemPrompt` to include instructions for Tata Naskah Dinas Formal (Pemerintah/BUMN) vs Korporat Modern, format tanggal baku Indonesia, dan tabel tindak lanjut MoM.
2. Update tool `generate_structured_doc` parameters schema to include new types and administrative properties (`letterNumber`, `signatory`, `firstParty`, `secondParty`, `actionItems`, `copiesTo`).
3. Connect tool execution handler to `buildAdministrativeDoc`.

- [ ] **Step 4: Run tests to confirm passes**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`

- [ ] **Step 5: Commit changes**

Git commit: `feat(agents): support administrative document generation in WordAgent`

---

### Task 4: UI Quick Action Presets for Administration

**Files:**
- Modify: `src/components/Chat/QuickActionPresets.tsx`
- Modify: `src/components/__tests__/chatAndActions.test.tsx`

**Interfaces:**
- Produces: New preset buttons for Word host in UI:
  - `🏛️ Surat Dinas Resmi`
  - `🤝 Berita Acara (BAST)`
  - `📋 Nota Dinas / Memo`

- [ ] **Step 1: Write test for new administrative presets**

Update `src/components/__tests__/chatAndActions.test.tsx` to assert that the new presets appear when host is Word and their prompts match expectations.

- [ ] **Step 2: Run test to confirm failure**

Run: `npx vitest run src/components/__tests__/chatAndActions.test.tsx`

- [ ] **Step 3: Implement presets in `QuickActionPresets.tsx`**

Add `word_official_letter`, `word_bast`, and `word_nota_dinas` to `PRESETS_BY_HOST.Word`.

- [ ] **Step 4: Run test to verify passes**

Run: `npx vitest run src/components/__tests__/chatAndActions.test.tsx`

- [ ] **Step 5: Commit changes**

Git commit: `feat(ui): add administrative document quick action presets for Word`

---

### Task 5: Full Regression Testing & Virtual Office Milestone Update

**Files:**
- Full test suite: `npx vitest run`

- [ ] **Step 1: Run complete test suite**

Ensure all 32+ test files pass without any regressions.

- [ ] **Step 2: Broadcast completion to Virtual Office 3D**

Execute `node E:/VIBE-CODE-WS/Sam-Virtual-Office/notify.js` to update status on the 3D Whiteboard.
