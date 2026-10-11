# Smart Mail Merge & Generator Dokumen Massal (Fase 3) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Mengimplementasikan mesin pembuatan surat massal (*Smart Mail Merge Engine*) untuk Microsoft Word yang memanfaatkan data tabel Excel dari Universal Hub (`crossAppBridge`), mengganti variabel placeholder `{{...}}`, dan menyusun dokumen multi-halaman dengan pemisah halaman (*Page Break*) otomatis.

**Architecture:** Modul terisolasi `mailMergeEngine.ts` mengekstrak tag placeholder, memetakan baris data tabel Excel ke variabel template secara case-insensitive, dan menyusun halaman dokumen individual. Lapisan driver Word mengeksekusi penulisan batch dokumen beserta *Page Break* dalam satu panggilan aman. `WordAgent` mengekspos tool `execute_mail_merge` ke ReAct loop.

**Tech Stack:** TypeScript 5.6, React 18.3, Office.js (Word), Vitest 3.2.

**Spec:** `docs/superpowers/specs/2026-10-11-smart-mail-merge-word-design.md`

## Global Constraints
- 100% Client-Side BYOK: tidak ada dependensi eksternal atau server relay.
- Kompatibilitas Office.js Word: multi-page batch write dengan page break.
- Menjaga 375 tes yang sudah ada tetap 100% lulus (zero regression).

---

### Task 1: Core Mail Merge Engine & Batch Composer

**Files:**
- Create: `src/services/office/mailMergeEngine.ts`
- Create: `src/services/office/__tests__/mailMergeEngine.test.ts`
- Modify: `src/services/office/types.ts`
- Modify: `src/services/office/index.ts`

**Interfaces:**
- Produces:
  - `extractPlaceholders(templateText: string): string[]`
  - `substitutePlaceholders(templateText: string, record: Record<string, any>): string`
  - `mapTableRowsToRecords(headers: string[], rows: (string | number)[][]): Array<Record<string, any>>`
  - `composeMailMergeBatch(options: MailMergeOptions, tableData: { headers: string[]; rows: (string | number)[][] }): { pages: string[]; result: MailMergeResult }`

- [ ] **Step 1: Write failing unit tests for mailMergeEngine**

Create `src/services/office/__tests__/mailMergeEngine.test.ts` testing:
1. `extractPlaceholders` extracts unique placeholder names from `{{Nama}}`, `{{ Alamat }}`, etc.
2. `substitutePlaceholders` replaces variables case-insensitively and handles missing keys safely.
3. `mapTableRowsToRecords` transforms 2D array rows into key-value records based on headers.
4. `composeMailMergeBatch` generates individual page texts and summary result.

- [ ] **Step 2: Run test to confirm failure**

Run: `npx vitest run src/services/office/__tests__/mailMergeEngine.test.ts`

- [ ] **Step 3: Implement `mailMergeEngine.ts` and update `types.ts` & `index.ts`**

Write `mailMergeEngine.ts` implementing the extraction, mapping, substitution, and batch composition logic.

- [ ] **Step 4: Run test to verify passes**

Run: `npx vitest run src/services/office/__tests__/mailMergeEngine.test.ts`

- [ ] **Step 5: Commit changes**

Git commit: `feat(office): implement core mail merge engine and placeholder substitution`

---

### Task 2: Word & Mock Driver Integration for Mail Merge

**Files:**
- Modify: `src/services/office/wordDriver.ts`
- Modify: `src/services/office/mockDriver.ts`
- Test: `src/services/office/__tests__/mockDriver.test.ts`

**Interfaces:**
- Consumes: `composeMailMergeBatch` from Task 1 and `getLatestCrossAppSnapshot` from storage
- Produces: `executeMailMerge(options: MailMergeOptions): Promise<MailMergeResult>` on `IDocumentDriver`

- [ ] **Step 1: Write test in `mockDriver.test.ts` for `executeMailMerge`**

Add tests verifying that `mockDriver.executeMailMerge` creates pages with page breaks and returns `MailMergeResult`.

- [ ] **Step 2: Run test to confirm failure**

Run: `npx vitest run src/services/office/__tests__/mockDriver.test.ts`

- [ ] **Step 3: Implement `executeMailMerge` in `mockDriver.ts` and `wordDriver.ts`**

Implement `executeMailMerge` using Word.run with `insertParagraph` and `insertBreak(Page)`.

- [ ] **Step 4: Run tests to confirm passes**

Run: `npx vitest run src/services/office/__tests__/mockDriver.test.ts`

- [ ] **Step 5: Commit changes**

Git commit: `feat(office): implement executeMailMerge in Word and mock drivers`

---

### Task 3: WordAgent Integration & System Prompt Enhancements

**Files:**
- Modify: `src/agents/word/wordAgent.ts`
- Modify: `src/agents/__tests__/specialistAgents.test.ts`

**Interfaces:**
- Consumes: `executeMailMerge` from `src/services/office`
- Produces: `execute_mail_merge` tool in `WordAgent`

- [ ] **Step 1: Write failing test in `specialistAgents.test.ts`**

Add test in `src/agents/__tests__/specialistAgents.test.ts` asserting that `WordAgent` exposes and executes `execute_mail_merge`.

- [ ] **Step 2: Run test to confirm failure**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`

- [ ] **Step 3: Update `wordAgent.ts`**

1. Update `getSystemPrompt` to instruct the agent on Mail Merge from Excel snapshots.
2. Add `execute_mail_merge` tool definition to `getTools()`.
3. Handle execution of `execute_mail_merge` in `executeTool()`.

- [ ] **Step 4: Run test to confirm passes**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`

- [ ] **Step 5: Commit changes**

Git commit: `feat(agents): add execute_mail_merge tool to WordAgent`

---

### Task 4: UI Quick Action Preset for Word Mail Merge

**Files:**
- Modify: `src/components/Chat/QuickActionPresets.tsx`
- Modify: `src/components/__tests__/chatAndActions.test.tsx`

**Interfaces:**
- Produces: `✉️ Mail Merge (Dari Excel Hub)` preset button for Word host

- [ ] **Step 1: Write test for new Word preset**

Update `src/components/__tests__/chatAndActions.test.tsx` to verify presence of `✉️ Mail Merge (Dari Excel Hub)`.

- [ ] **Step 2: Run test to confirm failure**

Run: `npx vitest run src/components/__tests__/chatAndActions.test.tsx`

- [ ] **Step 3: Add preset to `QuickActionPresets.tsx`**

Add `word_mail_merge` to `PRESETS_BY_HOST.Word`.

- [ ] **Step 4: Run test to confirm passes**

Run: `npx vitest run src/components/__tests__/chatAndActions.test.tsx`

- [ ] **Step 5: Commit changes**

Git commit: `feat(ui): add mail merge quick action preset for Word`

---

### Task 5: Full Regression Testing & Virtual Office Milestone Update

**Files:**
- Full test suite: `npx vitest run`

- [ ] **Step 1: Run complete test suite**

Ensure all test files pass with 0 regressions.

- [ ] **Step 2: Broadcast completion to Virtual Office 3D**

Execute `node E:/VIBE-CODE-WS/Sam-Virtual-Office/notify.js` to update status on the 3D Whiteboard.
