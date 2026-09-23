# Word Corporate Styler & Contract/Compliance Reviewer Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement Legal/Contract Compliance Reviewer and Corporate Brand Styler in Microsoft Word for Sam Office Agent.

**Architecture:** Extend `IOfficeDriver` / `IDocumentDriver` with `reviewComplianceClauses` and `applyCorporateStyle`. Implement both in `mockDriver.ts` and `wordDriver.ts`. Expose as native tools in `wordAgent.ts` and add presets in `QuickActionPresets.tsx`.

**Tech Stack:** TypeScript, React 18, Office.js (Word), Vitest, Tailwind CSS, Lucide icons.

**Spec:** [docs/superpowers/specs/2026-09-23-word-corporate-styler-compliance-design.md](file:///E:/VIBE-CODE-WS/Sam-Office-Agent/docs/superpowers/specs/2026-09-23-word-corporate-styler-compliance-design.md)

## Global Constraints

- Zero external backend relays: all execution runs in-process or via local server (`localhost:5173`).
- Single-sync batching for all Office.js Word operations via `await context.sync()`.
- Backward compatibility: All existing 269 tests must remain passing.
- TDD workflow: Write tests before implementation for every task.

---

### Task 1: Type Definitions & Driver Interfaces

**Files:**
- Modify: `src/services/office/types.ts`
- Modify: `tests/types.test.ts`

**Interfaces:**
- Produces:
  - `ComplianceClause`
  - `ComplianceReviewOptions`
  - `ComplianceReviewResult`
  - `CorporateStyleOptions`
  - `CorporateStyleResult`
  - Signatures in `IOfficeDriver` / `IDocumentDriver`:
    - `reviewComplianceClauses?(options?: ComplianceReviewOptions): Promise<ComplianceReviewResult>;`
    - `applyCorporateStyle?(options?: CorporateStyleOptions): Promise<CorporateStyleResult>;`

- [ ] **Step 1: Write the failing test**

In `tests/types.test.ts`:
```typescript
import { ComplianceClause, ComplianceReviewResult, CorporateStyleResult } from '../src/services/office/types';

describe('Word Compliance & Corporate Style Types', () => {
  it('validates ComplianceClause and ComplianceReviewResult structure', () => {
    const clause: ComplianceClause = {
      category: 'liability_indemnity',
      excerpt: 'Pihak kedua menanggung seluruh ganti rugi tanpa batas',
      status: 'high_risk',
      analysis: 'Klausul tanggung jawab tidak terbatas (unlimited liability)',
      recommendation: 'Batasi maksimal nilai ganti rugi senilai nilai kontrak',
    };
    const review: ComplianceReviewResult = {
      contractType: 'vendor_service',
      overallRiskLevel: 'high',
      clausesReviewedCount: 1,
      identifiedClauses: [clause],
      missingCriticalClauses: ['force_majeure'],
      executiveSummary: 'Ditemukan klausul risiko tinggi ganti rugi tanpa batas.',
      actionableRecommendations: ['Tambahkan pembatasan liability cap.'],
    };
    expect(review.overallRiskLevel).toBe('high');
    expect(review.identifiedClauses[0].status).toBe('high_risk');
  });

  it('validates CorporateStyleResult structure', () => {
    const res: CorporateStyleResult = {
      appliedTheme: 'corporate_navy',
      fontFamily: 'Calibri',
      styledParagraphsCount: 15,
      headingsCount: 3,
      message: 'Berhasil menerapkan tema Corporate Navy ke 15 paragraf',
    };
    expect(res.appliedTheme).toBe('corporate_navy');
    expect(res.headingsCount).toBe(3);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/types.test.ts`
Expected: FAIL with TS compilation / missing export error.

- [ ] **Step 3: Add type definitions in `src/services/office/types.ts`**

Define and export `ComplianceClause`, `ComplianceReviewOptions`, `ComplianceReviewResult`, `CorporateStyleOptions`, `CorporateStyleResult`. Update `IOfficeDriver` and `IDocumentDriver`.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/types.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/services/office/types.ts tests/types.test.ts
git commit -m "feat(types): add ComplianceReviewResult and CorporateStyleResult interfaces"
```

---

### Task 2: MockDriver & WordDriver Implementation

**Files:**
- Modify: `src/services/office/mockDriver.ts`
- Modify: `src/services/office/wordDriver.ts`
- Modify: `src/services/office/__tests__/mockDriver.test.ts`

**Interfaces:**
- Consumes: Types from `src/services/office/types.ts`.
- Produces: `reviewComplianceClauses` and `applyCorporateStyle` implemented in `MockOfficeDriver` and `WordDriver`.

- [ ] **Step 1: Write the failing tests**

In `src/services/office/__tests__/mockDriver.test.ts`:
```typescript
it('reviews compliance clauses and identifies high risk clauses', async () => {
  const driver = new MockOfficeDriver();
  driver.mockWordBody = 'Perjanjian Kerjasama. Klausul 1: Pembayaran 30 hari. Klausul 2: Pihak kedua menanggung seluruh ganti rugi tanpa batas untuk segala tuntutan.';
  
  const review = await driver.reviewComplianceClauses!({ contractType: 'vendor_service' });
  expect(review.clausesReviewedCount).toBeGreaterThan(0);
  expect(review.overallRiskLevel).toBe('high');
  const liability = review.identifiedClauses.find(c => c.category === 'liability_indemnity');
  expect(liability).toBeDefined();
  expect(liability?.status).toBe('high_risk');
  expect(review.missingCriticalClauses).toContain('force_majeure');
});

it('applies corporate style to mock document', async () => {
  const driver = new MockOfficeDriver();
  driver.mockWordBody = 'Judul Dokumen\nBab 1 Pendahuluan\nIni adalah isi paragraf dokumen kerja.';
  
  const styleRes = await driver.applyCorporateStyle!({ theme: 'corporate_navy', fontFamily: 'Calibri' });
  expect(styleRes.appliedTheme).toBe('corporate_navy');
  expect(styleRes.styledParagraphsCount).toBeGreaterThan(0);
  expect(styleRes.headingsCount).toBeGreaterThan(0);
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/services/office/__tests__/mockDriver.test.ts`
Expected: FAIL with `driver.reviewComplianceClauses is not a function`.

- [ ] **Step 3: Implement methods in `mockDriver.ts` and `wordDriver.ts`**

In `mockDriver.ts`:
- Implement `reviewComplianceClauses` scanning `this.mockWordBody` for keywords (liability, indemnity, termin pembayaran, denda, terminasi, kerahasiaan) and detecting missing force majeure clauses.
- Implement `applyCorporateStyle` simulating paragraph styling and headings count.

In `wordDriver.ts`:
- Implement `reviewComplianceClauses`: fetch text from selection or body, run compliance analysis, return structured result.
- Implement `applyCorporateStyle`: iterate over `body.paragraphs`, apply font formatting, heading sizes, and colors according to theme (`corporate_navy` with #1E3A8A accents, `executive_emerald` with #065F46, etc.).

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/services/office/__tests__/mockDriver.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/services/office/mockDriver.ts src/services/office/wordDriver.ts src/services/office/__tests__/mockDriver.test.ts
git commit -m "feat(driver): implement reviewComplianceClauses and applyCorporateStyle in WordDriver and MockOfficeDriver"
```

---

### Task 3: Word Specialist Agent Tools (`review_compliance_clauses` & `apply_corporate_style`)

**Files:**
- Modify: `src/agents/word/wordAgent.ts`
- Modify: `src/agents/__tests__/specialistAgents.test.ts`

**Interfaces:**
- Consumes: Driver methods from `IDocumentDriver`.
- Produces: Native tools `review_compliance_clauses` and `apply_corporate_style` available in `wordAgent.getTools()` and callable via `wordAgent.executeTool()`.

- [ ] **Step 1: Write the failing tests**

In `src/agents/__tests__/specialistAgents.test.ts`:
```typescript
it('provides review_compliance_clauses tool and executes it successfully', async () => {
  const driver = new MockOfficeDriver();
  driver.mockWordBody = 'Perjanjian Sewa. Pihak kedua wajib membayar denda keterlambatan 5% per hari tanpa batas maksimal.';
  const agent = new WordAgent(driver);
  const tools = agent.getTools();
  expect(tools.some(t => t.name === 'review_compliance_clauses')).toBe(true);

  const res = await agent.executeTool(
    { id: 'w1', name: 'review_compliance_clauses', arguments: { contractType: 'vendor_service' }, status: 'pending' },
    { host: 'Word' }
  );
  expect(res.success).toBe(true);
  expect(res.result).toContain('Review Kepatuhan Kontrak');
  expect(res.result).toContain('Tingkat Risiko');
});

it('provides apply_corporate_style tool and executes it successfully', async () => {
  const driver = new MockOfficeDriver();
  driver.mockWordBody = 'Laporan Tahunan\nBab 1 Latar Belakang\nIsi paragraf laporan resmi.';
  const agent = new WordAgent(driver);
  const tools = agent.getTools();
  expect(tools.some(t => t.name === 'apply_corporate_style')).toBe(true);

  const res = await agent.executeTool(
    { id: 'w2', name: 'apply_corporate_style', arguments: { theme: 'corporate_navy' }, status: 'pending' },
    { host: 'Word' }
  );
  expect(res.success).toBe(true);
  expect(res.result).toContain('Corporate Navy');
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`
Expected: FAIL with `tool review_compliance_clauses not found`.

- [ ] **Step 3: Register and handle tools in `src/agents/word/wordAgent.ts`**

Support `constructor(private driverOverride?: any)`. Register `review_compliance_clauses` and `apply_corporate_style` in `getTools()`. In `executeTool()`, call the driver methods and format results into clear, structured Markdown.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/agents/word/wordAgent.ts src/agents/__tests__/specialistAgents.test.ts
git commit -m "feat(agent): register review_compliance_clauses and apply_corporate_style in WordAgent"
```

---

### Task 4: Quick Action Presets Toolbar Expansion for Word

**Files:**
- Modify: `src/components/Chat/QuickActionPresets.tsx`
- Modify: `src/components/__tests__/chatAndActions.test.tsx`

**Interfaces:**
- Produces: Updated Word presets in `PRESETS_BY_HOST.Word` with `"⚖️ Review Kontrak & Risiko"` and `"🎨 Format Brand Korporat"`.

- [ ] **Step 1: Write the failing test**

In `src/components/__tests__/chatAndActions.test.tsx`:
```typescript
it('renders new Word presets for Review Kontrak and Format Brand Korporat', () => {
  const onSelect = vi.fn();
  render(<QuickActionPresets host="Word" onSelectPreset={onSelect} />);

  expect(screen.getByText(/Review Kontrak/i)).toBeInTheDocument();
  expect(screen.getByText(/Format Brand Korporat/i)).toBeInTheDocument();

  fireEvent.click(screen.getByText(/Review Kontrak/i));
  expect(onSelect).toHaveBeenCalledWith(expect.stringContaining('Audit dan review kepatuhan klausul'));
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/components/__tests__/chatAndActions.test.tsx`
Expected: FAIL with preset button not found.

- [ ] **Step 3: Implement presets in `QuickActionPresets.tsx`**

Add presets:
- `{ id: 'word_review_contract', label: '⚖️ Review Kontrak & Risiko', prompt: 'Audit dan review kepatuhan klausul pada dokumen/kontrak ini: periksa SLA, denda, termin pembayaran, klausul risiko tinggi, dan berikan rekomendasi perbaikan.' }`
- `{ id: 'word_corporate_style', label: '🎨 Format Brand Korporat', prompt: 'Terapkan standarisasi gaya dan format korporat profesional (Corporate Navy) pada seluruh dokumen ini: tata hierarki heading, font, dan spasi yang rapi.' }`

- [ ] **Step 4: Run all tests to verify 100% pass**

Run: `npm test -- --run`
Expected: All 270+ tests PASS.

- [ ] **Step 5: Run production build**

Run: `npm run build`
Expected: Zero TypeScript or Vite compilation errors.

- [ ] **Step 6: Commit**

```bash
git add src/components/Chat/QuickActionPresets.tsx src/components/__tests__/chatAndActions.test.tsx
git commit -m "feat(ui): add Review Kontrak and Format Brand Korporat presets for Word"
```
