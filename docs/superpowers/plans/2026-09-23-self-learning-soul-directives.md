# Self-Learning Corporate Brand Voice & SOUL Directives Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement a persistent, self-learning corporate brand voice and directives system (`SOUL.md`) that allows Sam to autonomously learn user preferences and formatting rules from conversation, inject them into all agent prompts, and provide a full management UI in Settings.

**Architecture:** A dedicated storage module (`src/services/storage/soulStorage.ts`) persists `SoulConfig`, presets, raw Markdown, and `learnedDirectives`. An autonomous ReAct tool (`learn_corporate_directive`) is registered across all specialist agents (`ExcelAgent`, `WordAgent`, `PPTAgent`) to capture and persist user feedback. `SamCoordinator.buildSystemPrompt()` dynamically retrieves and injects active directives into the LLM system prompt. `SettingsModal.tsx` provides preset switching, Markdown editing, file import/export, and pruning of learned rules.

**Tech Stack:** TypeScript, React, LocalStorage, Vitest.

**Spec:** `docs/superpowers/specs/2026-09-23-self-learning-soul-directives-design.md`

## Global Constraints
- Zero external dependencies.
- 100% client-side persistence via `localStorage` with safe fallback for Node/browser environments.
- Strict Markdown formatting for `SOUL.md` export/import.
- Presets: `formal_executive`, `modern_professional`, `financial_compliance`, and `custom`.
- All tests must pass with 100% green and 0 regressions on `npm run build` and `npx vitest run`.

---

### Task 1: SOUL Storage Engine & Data Types

**Files:**
- Modify: `src/types/index.ts`
- Create: `src/services/storage/soulStorage.ts`
- Create: `tests/soulStorage.test.ts`

**Interfaces:**
- Produces:
  - `DirectiveCategory`: `'tone' | 'terminology' | 'formatting' | 'constraint' | 'general'`
  - `LearnedDirective`: `{ id: string; category: DirectiveCategory; rule: string; learnedAt: number; source?: string; explanation?: string; }`
  - `BrandVoicePreset`: `'formal_executive' | 'modern_professional' | 'financial_compliance' | 'custom'`
  - `SoulConfig`: `{ enabled: boolean; corporateName: string; brandVoicePreset: BrandVoicePreset; rawSoulMarkdown: string; learnedDirectives: LearnedDirective[]; lastUpdated: number; }`
  - `getSoulConfig(): SoulConfig`
  - `saveSoulConfig(config: Partial<SoulConfig>): SoulConfig`
  - `addLearnedDirective(rule: string, category?: DirectiveCategory, explanation?: string, source?: string): LearnedDirective`
  - `removeLearnedDirective(id: string): void`
  - `exportSoulToMarkdown(): string`
  - `importSoulFromMarkdown(markdownText: string): SoulConfig`
  - `clearSoulConfig(): void`

- [ ] **Step 1: Write the failing test**

Create `tests/soulStorage.test.ts`:

```typescript
import { describe, it, expect, beforeEach } from 'vitest';
import {
  getSoulConfig,
  saveSoulConfig,
  addLearnedDirective,
  removeLearnedDirective,
  exportSoulToMarkdown,
  importSoulFromMarkdown,
  clearSoulConfig,
} from '../src/services/storage/soulStorage';

describe('soulStorage', () => {
  beforeEach(() => {
    clearSoulConfig();
  });

  it('provides default corporate SoulConfig', () => {
    const config = getSoulConfig();
    expect(config.enabled).toBe(true);
    expect(config.brandVoicePreset).toBe('formal_executive');
    expect(config.corporateName).toBeDefined();
    expect(config.rawSoulMarkdown).toContain('SOUL & CORPORATE BRAND DIRECTIVES');
    expect(config.learnedDirectives).toEqual([]);
  });

  it('updates SoulConfig partially and persists changes', () => {
    const updated = saveSoulConfig({
      corporateName: 'PT Samudra Mega Corpora',
      brandVoicePreset: 'modern_professional',
    });

    expect(updated.corporateName).toBe('PT Samudra Mega Corpora');
    expect(updated.brandVoicePreset).toBe('modern_professional');

    const fresh = getSoulConfig();
    expect(fresh.corporateName).toBe('PT Samudra Mega Corpora');
  });

  it('adds and removes learned directives correctly', () => {
    const directive = addLearnedDirective(
      'Gunakan istilah inisiatif alih-alih skema',
      'terminology',
      'Koreksi dari pimpinan rapat'
    );

    expect(directive.id).toBeDefined();
    expect(directive.rule).toContain('inisiatif');
    expect(directive.category).toBe('terminology');

    let config = getSoulConfig();
    expect(config.learnedDirectives.length).toBe(1);

    removeLearnedDirective(directive.id);
    config = getSoulConfig();
    expect(config.learnedDirectives.length).toBe(0);
  });

  it('exports to markdown and imports back correctly', () => {
    addLearnedDirective('Format mata uang wajib Rp', 'formatting');
    const md = exportSoulToMarkdown();

    expect(md).toContain('SOUL & CORPORATE BRAND DIRECTIVES');
    expect(md).toContain('Format mata uang wajib Rp');

    clearSoulConfig();
    expect(getSoulConfig().learnedDirectives.length).toBe(0);

    const imported = importSoulFromMarkdown(md);
    expect(imported.rawSoulMarkdown).toContain('SOUL & CORPORATE BRAND DIRECTIVES');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/soulStorage.test.ts`
Expected: FAIL with "Cannot find module '../src/services/storage/soulStorage'".

- [ ] **Step 3: Write minimal implementation**

1. In `src/types/index.ts`, export the types:
   `DirectiveCategory`, `LearnedDirective`, `BrandVoicePreset`, `SoulConfig`.
2. In `src/services/storage/soulStorage.ts`:
   - Define presets:
     - `formal_executive`: Formal, data-driven, executive summaries, Indonesian corporate conventions.
     - `modern_professional`: Active voice, concise, agile, tech-oriented.
     - `financial_compliance`: High precision, audit trails, strict currency/numbering.
     - `custom`: User-defined.
   - Implement storage methods with safe fallback for non-browser/SSR environments:
     - `getSoulConfig()`
     - `saveSoulConfig()`
     - `addLearnedDirective()`
     - `removeLearnedDirective()`
     - `exportSoulToMarkdown()`
     - `importSoulFromMarkdown()`
     - `clearSoulConfig()`

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/soulStorage.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/types/index.ts src/services/storage/soulStorage.ts tests/soulStorage.test.ts
git commit -m "feat(storage): implement SOUL storage engine and corporate brand directives"
```

---

### Task 2: Autonomous ReAct Tool `learn_corporate_directive`

**Files:**
- Modify: `src/agents/excel/excelAgent.ts`
- Modify: `src/agents/word/wordAgent.ts`
- Modify: `src/agents/powerpoint/pptAgent.ts`
- Test: `tests/selfLearningSoul.test.ts`

**Interfaces:**
- Consumes: `addLearnedDirective` from `../../services/storage/soulStorage`
- Produces: `learn_corporate_directive` tool across `ExcelAgent`, `WordAgent`, `PPTAgent`

- [ ] **Step 1: Write the failing test**

Create `tests/selfLearningSoul.test.ts`:

```typescript
import { describe, it, expect, beforeEach } from 'vitest';
import { ExcelAgent } from '../src/agents/excel/excelAgent';
import { WordAgent } from '../src/agents/word/wordAgent';
import { PPTAgent } from '../src/agents/powerpoint/pptAgent';
import { MockOfficeDriver } from '../src/services/office/mockDriver';
import { clearSoulConfig, getSoulConfig } from '../src/services/storage/soulStorage';

describe('Self-Learning SOUL Directives Tooling', () => {
  beforeEach(() => {
    clearSoulConfig();
  });

  it('provides learn_corporate_directive tool in all specialist agents', () => {
    const mockDriver = new MockOfficeDriver();
    const excelAgent = new ExcelAgent(mockDriver);
    const wordAgent = new WordAgent(mockDriver);
    const pptAgent = new PPTAgent(mockDriver);

    expect(excelAgent.getTools().some(t => t.name === 'learn_corporate_directive')).toBe(true);
    expect(wordAgent.getTools().some(t => t.name === 'learn_corporate_directive')).toBe(true);
    expect(pptAgent.getTools().some(t => t.name === 'learn_corporate_directive')).toBe(true);
  });

  it('executes learn_corporate_directive and persists learned rule', async () => {
    const mockDriver = new MockOfficeDriver();
    const wordAgent = new WordAgent(mockDriver);

    const res = await wordAgent.executeTool(
      {
        id: 'learn-1',
        name: 'learn_corporate_directive',
        arguments: {
          rule: 'Gunakan istilah mitra strategis untuk klien prioritas',
          category: 'terminology',
          explanation: 'Koreksi penulisan dokumen kontrak',
        },
        status: 'pending',
      },
      { host: 'Word' }
    );

    expect(res.success).toBe(true);
    expect(res.result).toContain('Berhasil mempelajari direktif korporat');
    expect(res.result).toContain('mitra strategis');

    const config = getSoulConfig();
    expect(config.learnedDirectives.length).toBe(1);
    expect(config.learnedDirectives[0].rule).toContain('mitra strategis');
    expect(config.learnedDirectives[0].category).toBe('terminology');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/selfLearningSoul.test.ts`
Expected: FAIL with "learn_corporate_directive tool not found".

- [ ] **Step 3: Write minimal implementation**

1. Define common tool declaration and handler in `src/agents/excel/excelAgent.ts`, `wordAgent.ts`, `pptAgent.ts`:
   - Tool definition:
     ```typescript
     {
       name: 'learn_corporate_directive',
       description: 'Mencatat dan mempelajari aturan gaya baru, preferensi format, koreksi istilah, atau direktif penulisan dari pengguna ke dalam memori persisten korporat (SOUL.md) agar selalu dipatuhi di masa mendatang.',
       parameters: {
         type: 'object',
         properties: {
           rule: { type: 'string', description: 'Aturan atau preferensi spesifik yang harus dipelajari' },
           category: {
             type: 'string',
             enum: ['tone', 'terminology', 'formatting', 'constraint', 'general'],
             description: 'Kategori direktif yang dipelajari.',
           },
           explanation: { type: 'string', description: 'Alasan atau konteks aturan ini dipelajari.' },
         },
         required: ['rule'],
       },
     }
     ```
   - In `executeTool()`:
     - Call `addLearnedDirective(rule, category, explanation, 'react_tool')`.
     - Return formatted confirmation:
       `"🧠 Berhasil mempelajari direktif korporat: '${rule}' (Kategori: ${category || 'general'}). Aturan ini telah dicatat ke dalam SOUL.md dan akan otomatis diterapkan pada seluruh dokumen mendatang."`
   - In `getSystemPrompt()`: Add guidance to autonomously call `learn_corporate_directive` when the user provides styling rules or corrections.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/selfLearningSoul.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/agents/excel/excelAgent.ts src/agents/word/wordAgent.ts src/agents/powerpoint/pptAgent.ts tests/selfLearningSoul.test.ts
git commit -m "feat(agents): register and execute learn_corporate_directive tool across specialist agents"
```

---

### Task 3: Dynamic System Prompt Injection in SamCoordinator

**Files:**
- Modify: `src/agents/coordinator/samCoordinator.ts`
- Test: `src/agents/coordinator/__tests__/samCoordinator.test.ts`

**Interfaces:**
- Consumes: `getSoulConfig` from `../../services/storage/soulStorage`
- Produces: Dynamic inclusion of corporate brand voice & learned directives in `SamCoordinator.buildSystemPrompt()`

- [ ] **Step 1: Write the failing test**

In `src/agents/coordinator/__tests__/samCoordinator.test.ts`, add a test verifying dynamic `SOUL.md` injection:

```typescript
it('dynamically injects active SOUL.md directives and learned rules into system prompt', () => {
  const coordinator = new SamCoordinator();
  saveSoulConfig({
    enabled: true,
    corporateName: 'PT Telko Digital Nusantara',
    brandVoicePreset: 'formal_executive',
  });
  addLearnedDirective('Format seluruh angka dengan pemisah titik', 'formatting');

  const prompt = coordinator.buildSystemPrompt({ host: 'Word' });
  expect(prompt).toContain('CORPORATE BRAND VOICE & SOUL DIRECTIVES');
  expect(prompt).toContain('PT Telko Digital Nusantara');
  expect(prompt).toContain('Format seluruh angka dengan pemisah titik');

  // When disabled, SOUL directives should be excluded
  saveSoulConfig({ enabled: false });
  const disabledPrompt = coordinator.buildSystemPrompt({ host: 'Word' });
  expect(disabledPrompt).not.toContain('CORPORATE BRAND VOICE & SOUL DIRECTIVES');
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/agents/coordinator/__tests__/samCoordinator.test.ts`
Expected: FAIL with "expected prompt to contain 'CORPORATE BRAND VOICE & SOUL DIRECTIVES'".

- [ ] **Step 3: Write minimal implementation**

In `src/agents/coordinator/samCoordinator.ts`:
- Import `getSoulConfig` from `../../services/storage/soulStorage`.
- In `buildSystemPrompt()`:
  - Retrieve `const soul = getSoulConfig()`.
  - If `soul.enabled`:
    - Append formatted section:
      ```markdown
      === CORPORATE BRAND VOICE & SOUL DIRECTIVES (SOUL.md) ===
      Nama Entitas: ${soul.corporateName}
      Preset Karakter: ${soul.brandVoicePreset}

      ${soul.rawSoulMarkdown}

      ${soul.learnedDirectives.length > 0 ? `Aturan yang Dipelajari Secara Mandiri (Self-Learned Directives):\n${soul.learnedDirectives.map(d => `- [${d.category.toUpperCase()}]: ${d.rule}`).join('\n')}` : ''}

      PANDUAN KEPATUHAN KORPORAT: Anda WAJIB mematuhi seluruh direktif SOUL di atas dalam setiap penyusunan naskah, tabel angka, dan slide presentasi.
      === END CORPORATE DIRECTIVES ===
      ```
  - Add Rule 8 to the coordinator rules list:
    `"8. CORPORATE DIRECTIVES & SELF-LEARNING (SOUL.md): Jika pengguna memberikan arahan gaya baru, koreksi istilah, atau aturan penulisan spesifik, gunakan tool learn_corporate_directive untuk menyimpannya ke memori korporat."`

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/agents/coordinator/__tests__/samCoordinator.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/agents/coordinator/samCoordinator.ts src/agents/coordinator/__tests__/samCoordinator.test.ts
git commit -m "feat(coordinator): dynamically inject active SOUL directives and learned rules into system prompt"
```

---

### Task 4: UI Settings Management & Presets

**Files:**
- Modify: `src/components/Settings/SettingsModal.tsx`
- Modify: `src/components/Chat/QuickActionPresets.tsx`
- Test: `src/components/__tests__/headerAndSettings.test.tsx`
- Test: `src/components/__tests__/chatAndActions.test.tsx`

**Interfaces:**
- Consumes: `getSoulConfig`, `saveSoulConfig`, `addLearnedDirective`, `removeLearnedDirective`, `exportSoulToMarkdown`, `importSoulFromMarkdown`
- Produces:
  - "Brand Voice & SOUL.md" panel in SettingsModal
  - `universal_learn_soul` quick action preset

- [ ] **Step 1: Write the failing tests**

1. In `src/components/__tests__/headerAndSettings.test.tsx`:
   - Add test verifying that "Brand Voice & SOUL.md" section and preset buttons render in `SettingsModal`.
2. In `src/components/__tests__/chatAndActions.test.tsx`:
   - Add test verifying that `🧠 Pelajari Preferensi Ini` preset renders.

- [ ] **Step 2: Run tests to verify they fail**

Run: `npx vitest run src/components/__tests__/headerAndSettings.test.tsx src/components/__tests__/chatAndActions.test.tsx`
Expected: FAIL.

- [ ] **Step 3: Write minimal implementation**

1. In `src/components/Settings/SettingsModal.tsx`:
   - Add "Brand Voice & SOUL.md" configuration card:
     - Toggle: Enable / Disable.
     - Corporate Name input.
     - Brand Voice Preset Selector (`formal_executive`, `modern_professional`, `financial_compliance`, `custom`). Changing preset updates `rawSoulMarkdown`.
     - Textarea for editing `rawSoulMarkdown`.
     - Badges list of `learnedDirectives` with a delete button (`✕`) calling `removeLearnedDirective(id)`.
     - Export button: downloads `SOUL.md`.
     - Import button: file input that reads `.md` and calls `importSoulFromMarkdown()`.
2. In `src/components/Chat/QuickActionPresets.tsx`:
   - Add universal preset across hosts:
     ```typescript
     {
       id: 'universal_learn_soul',
       label: '🧠 Pelajari Preferensi Ini',
       prompt: 'Catat gaya bahasa, aturan format angka, dan preferensi istilah dari dokumen ini ke dalam direktif korporat (SOUL.md) agar selalu diterapkan pada pekerjaan berikutnya.',
     }
     ```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npx vitest run src/components/__tests__/headerAndSettings.test.tsx src/components/__tests__/chatAndActions.test.tsx`
Expected: PASS.

- [ ] **Step 5: Full verification & regression suite**

Run:
```bash
npm run build
npx vitest run
```
Expected: All build checks pass with exit code 0, and all test files pass 100% green.

- [ ] **Step 6: Commit**

```bash
git add src/components/Settings/SettingsModal.tsx src/components/Chat/QuickActionPresets.tsx src/components/__tests__/
git commit -m "feat(ui): add Brand Voice & SOUL directives management panel and quick action presets"
```
