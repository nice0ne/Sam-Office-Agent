# PowerPoint Doc-to-Deck Transformer Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement the PowerPoint Doc-to-Deck Transformer for Sam Office Agent, enabling automated structuring of prose text and reports into an executive 5-slide storyline arc with full presenter speech notes in chat.

**Architecture:** Build a dedicated transformation engine (`docToDeckTransformer.ts`) that extracts titles, partitions narratives into the Executive Storyline Arc (Cover, Context, Strategy, Metrics, Roadmap), and synthesizes verbal speaker scripts. Integrate `transformDocToDeck` into `mockDriver.ts` and `pptDriver.ts`, expose it via `transform_doc_to_deck` in `PPTAgent.ts`, and provide a one-click Quick Action Preset in `QuickActionPresets.tsx`.

**Tech Stack:** TypeScript, React 18, Office.js (PowerPoint), Vitest, Tailwind CSS.

**Spec:** [docs/superpowers/specs/2026-09-23-powerpoint-doc-to-deck-transformer-design.md](file:///E:/VIBE-CODE-WS/Sam-Office-Agent/docs/superpowers/specs/2026-09-23-powerpoint-doc-to-deck-transformer-design.md)

## Global Constraints

- Zero external backend relays: 100% Client-Side BYOK in WebView2 sandbox.
- Single-sync batching for all Office.js PowerPoint shape operations via `await context.sync()`.
- Backward compatibility: All existing 276 tests must remain passing.
- TDD workflow: Write tests before implementation for every task.
- Presenter notes: Presented richly in chat taskpane with hook, key points, and transition.

---

### Task 1: Type Definitions & Driver Interfaces

**Files:**
- Modify: `src/services/office/types.ts`
- Modify: `tests/types.test.ts`

**Interfaces:**
- Produces:
  - `DocToDeckSlide`
  - `DocToDeckOptions`
  - `DocToDeckResult`
  - Signature in `IOfficeDriver` & `IDocumentDriver`:
    `transformDocToDeck?(options?: DocToDeckOptions): Promise<DocToDeckResult>;`

- [ ] **Step 1: Write the failing test**

In `tests/types.test.ts`:
```typescript
import { DocToDeckSlide, DocToDeckOptions, DocToDeckResult } from '../src/services/office/types';

describe('DocToDeck Types', () => {
  it('validates DocToDeckSlide and DocToDeckResult contracts', () => {
    const slide: DocToDeckSlide = {
      title: 'Tinjauan Strategis',
      category: 'context',
      bullets: ['Tantangan integrasi sistem', 'Peningkatan kebutuhan efisiensi'],
      speakerScript: {
        hook: 'Bapak/Ibu sekalian, mari kita mulai dengan memahami latar belakang urgensi proyek ini.',
        keyTalkingPoints: ['Sistem saat ini menghadapi tantangan integrasi.', 'Efisiensi operasional perlu ditingkatkan segera.'],
        transition: 'Selanjutnya, mari kita telusuri pilar solusi yang diusulkan.',
      },
    };
    const result: DocToDeckResult = {
      deckTitle: 'Presentasi Strategis',
      appliedTheme: 'corporate_blue',
      totalSlidesCreated: 1,
      slides: [slide],
      summaryMessage: 'Berhasil membuat 1 slide.',
    };
    expect(result.deckTitle).toBe('Presentasi Strategis');
    expect(result.slides[0].category).toBe('context');
    expect(result.slides[0].speakerScript.hook).toBeDefined();
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/types.test.ts`
Expected: FAIL with TS compilation / missing export error.

- [ ] **Step 3: Add type definitions in `src/services/office/types.ts`**

Define and export `DocToDeckSlide`, `DocToDeckOptions`, `DocToDeckResult`. Update `IOfficeDriver` and `IDocumentDriver`.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/types.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/services/office/types.ts tests/types.test.ts
git commit -m "feat(types): add DocToDeck contracts and driver signatures"
```

---

### Task 2: Core Transformation Engine (`docToDeckTransformer.ts`)

**Files:**
- Create: `src/services/office/docToDeckTransformer.ts`
- Create: `tests/docToDeckTransformer.test.ts`

**Interfaces:**
- Consumes: `DocToDeckSlide`, `DocToDeckOptions`, `DocToDeckResult` from `src/services/office/types.ts`.
- Produces: `synthesizeDocToDeck(rawText?: string, options?: DocToDeckOptions): DocToDeckResult`

- [ ] **Step 1: Write the failing test**

In `tests/docToDeckTransformer.test.ts`:
```typescript
import { describe, it, expect } from 'vitest';
import { synthesizeDocToDeck } from '../src/services/office/docToDeckTransformer';

describe('synthesizeDocToDeck', () => {
  it('partitions raw report text into 5-slide Executive Storyline Arc', () => {
    const rawText = `Laporan Inisiatif Transformasi Digital 2026
Latar Belakang: Proses pelaporan manual membutuhkan waktu 4 hari per minggu dan rawan kesalahan manusia.
Solusi: Penerapan Sam Office Agent untuk otomatisasi dokumen Excel, Word, dan PowerPoint.
Capaian & Metrik: Peningkatan efisiensi waktu hingga 75%, akurasi data mencapai 99.8%, kepuasan pengguna 95%.
Rencana Aksi: Sosialisasi seluruh divisi pada Oktober 2026 dan evaluasi triwulanan.`;

    const result = synthesizeDocToDeck(rawText, { theme: 'corporate_blue' });

    expect(result.totalSlidesCreated).toBe(5);
    expect(result.deckTitle).toContain('Transformasi Digital');
    expect(result.appliedTheme).toBe('corporate_blue');

    const categories = result.slides.map(s => s.category);
    expect(categories).toEqual(['cover', 'context', 'strategy', 'metrics', 'roadmap']);

    for (const slide of result.slides) {
      expect(slide.title.length).toBeGreaterThan(0);
      expect(slide.speakerScript.hook.length).toBeGreaterThan(0);
      expect(slide.speakerScript.keyTalkingPoints.length).toBeGreaterThan(0);
      expect(slide.speakerScript.transition.length).toBeGreaterThan(0);
    }
  });

  it('handles empty or brief text gracefully with default briefing structure', () => {
    const result = synthesizeDocToDeck('', { presentationTitle: 'Briefing Eksekutif' });
    expect(result.totalSlidesCreated).toBeGreaterThanOrEqual(4);
    expect(result.deckTitle).toBe('Briefing Eksekutif');
    expect(result.slides[0].category).toBe('cover');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/docToDeckTransformer.test.ts`
Expected: FAIL with module not found.

- [ ] **Step 3: Implement `synthesizeDocToDeck` in `src/services/office/docToDeckTransformer.ts`**

Implement text parsing, keyword extraction, executive arc generator, and Indonesian presenter speech script synthesizer.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/docToDeckTransformer.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/services/office/docToDeckTransformer.ts tests/docToDeckTransformer.test.ts
git commit -m "feat(ppt): implement core docToDeckTransformer synthesis engine"
```

---

### Task 3: Driver Integration (`mockDriver.ts` & `pptDriver.ts`)

**Files:**
- Modify: `src/services/office/mockDriver.ts`
- Modify: `src/services/office/pptDriver.ts`
- Modify: `src/services/office/__tests__/mockDriver.test.ts`

**Interfaces:**
- Consumes: `synthesizeDocToDeck` from `src/services/office/docToDeckTransformer.ts`.
- Produces: `MockOfficeDriver.transformDocToDeck` and `PPTDriver.transformDocToDeck`.

- [ ] **Step 1: Write the failing test**

In `src/services/office/__tests__/mockDriver.test.ts`:
```typescript
it('transforms document text into executive slide deck in mock driver', async () => {
  const driver = new MockOfficeDriver();
  const sampleDoc = `Rencana Kerja Q4
Tantangan: Keterlambatan konsolidasi anggaran antar divisi.
Solusi: Sistem approval satu pintu berbasis Office Add-in.
Target: SLA berkurang dari 5 hari menjadi 1 hari.`;

  const result = await driver.transformDocToDeck!({
    documentText: sampleDoc,
    theme: 'corporate_blue',
  });

  expect(result.totalSlidesCreated).toBeGreaterThanOrEqual(4);
  expect(driver.slides.length).toBe(result.totalSlidesCreated);
  expect(result.slides[0].title).toBeDefined();
  expect(result.slides[0].speakerScript.hook).toBeDefined();
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/services/office/__tests__/mockDriver.test.ts`
Expected: FAIL with `driver.transformDocToDeck is not a function`.

- [ ] **Step 3: Implement `transformDocToDeck` in `mockDriver.ts` and `pptDriver.ts`**

In `mockDriver.ts`:
- Call `synthesizeDocToDeck`.
- Push each slide into `this.slides` with title, bullets, notes.
- Return `DocToDeckResult`.

In `pptDriver.ts`:
- Call `synthesizeDocToDeck`.
- In `PowerPoint.run`: create slides, add title and bullet textboxes with theme styling, and sync.
- Return `DocToDeckResult`.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/services/office/__tests__/mockDriver.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/services/office/mockDriver.ts src/services/office/pptDriver.ts src/services/office/__tests__/mockDriver.test.ts
git commit -m "feat(driver): implement transformDocToDeck in mockDriver and pptDriver"
```

---

### Task 4: PowerPoint Specialist Agent Tool (`PPTAgent.ts`)

**Files:**
- Modify: `src/agents/powerpoint/pptAgent.ts`
- Modify: `src/agents/__tests__/specialistAgents.test.ts`

**Interfaces:**
- Consumes: `driver.transformDocToDeck`.
- Produces: Native tool `transform_doc_to_deck` in `PPTAgent.getTools()` and handler in `PPTAgent.executeTool()`.

- [ ] **Step 1: Write the failing test**

In `src/agents/__tests__/specialistAgents.test.ts`:
```typescript
it('provides transform_doc_to_deck tool and formats executive presentation dashboard', async () => {
  const mockDriver = new MockOfficeDriver();
  const testAgent = new PPTAgent(mockDriver);
  const tools = testAgent.getTools();
  expect(tools.some(t => t.name === 'transform_doc_to_deck')).toBe(true);

  const res = await testAgent.executeTool(
    {
      id: 'ppt-doc-1',
      name: 'transform_doc_to_deck',
      arguments: {
        documentText: 'Proposal Modernisasi TI. Masalah: Sistem warisan. Solusi: Migrasi cloud modern. Metrik: Uptime 99.9%.',
        theme: 'corporate_blue',
      },
      status: 'pending',
    },
    { host: 'PowerPoint' }
  );

  expect(res.success).toBe(true);
  expect(res.result).toContain('Transformasi Dokumen ke Slide Presentasi');
  expect(res.result).toContain('Naskah Presenter');
  expect(res.result).toContain('Pembuka');
  expect(res.result).toContain('Transisi');
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`
Expected: FAIL with missing tool or tool error.

- [ ] **Step 3: Implement `transform_doc_to_deck` in `src/agents/powerpoint/pptAgent.ts`**

- Add `constructor(private driverOverride?: any) {}` to `PPTAgent`.
- Register `transform_doc_to_deck` in `getTools()`.
- Implement tool execution in `executeTool()`.
- Update `getSystemPrompt()` to guide on Doc-to-Deck usage.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/agents/powerpoint/pptAgent.ts src/agents/__tests__/specialistAgents.test.ts
git commit -m "feat(agent): register and handle transform_doc_to_deck in PPTAgent"
```

---

### Task 5: UI Quick Action Preset & Full Suite Verification

**Files:**
- Modify: `src/components/Chat/QuickActionPresets.tsx`
- Modify: `src/components/__tests__/chatAndActions.test.tsx`

**Interfaces:**
- Consumes: Preset in `PRESETS_BY_HOST.PowerPoint`.
- Produces: Visual preset button in PowerPoint host mode.

- [ ] **Step 1: Write the failing test**

In `src/components/__tests__/chatAndActions.test.tsx`:
```typescript
it('renders ppt_doc_to_deck preset in PowerPoint host mode', () => {
  render(<QuickActionPresets host="PowerPoint" onSelectPreset={vi.fn()} />);
  expect(screen.getByText(/Dokumen ke Slide/i)).toBeInTheDocument();
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/components/__tests__/chatAndActions.test.tsx`
Expected: FAIL (text not found).

- [ ] **Step 3: Add preset to `src/components/Chat/QuickActionPresets.tsx`**

Add preset item to `PRESETS_BY_HOST.PowerPoint`:
```typescript
{
  id: 'ppt_doc_to_deck',
  label: '📊 Dokumen ke Slide (Doc-to-Deck)',
  prompt: 'Ubah teks/dokumen laporan ini menjadi 5 slide presentasi eksekutif terstruktur lengkap dengan naskah pembicara (speaker notes).',
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/components/__tests__/chatAndActions.test.tsx`
Expected: PASS.

- [ ] **Step 5: Run full test suite to ensure 0 regressions**

Run: `npx vitest run`
Expected: All tests PASS with 0 failures.

- [ ] **Step 6: Commit**

```bash
git add src/components/Chat/QuickActionPresets.tsx src/components/__tests__/chatAndActions.test.tsx
git commit -m "feat(ui): add Doc-to-Deck quick action preset for PowerPoint"
```
