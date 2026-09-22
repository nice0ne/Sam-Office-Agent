# Excel Deep Intelligence & Safe Dry-Run Execution Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement Excel Deep Intelligence (Formula Auditor & Anomaly Detector, Executive Data Storyteller) and Safe Dry-Run Confirmation Card in Sam Office Agent.

**Architecture:** Extend `IOfficeDriver` and drivers (`excelDriver.ts`, `mockDriver.ts`) with `auditSheetData` and `generateDataStory`. Expose them as native tools in `excelAgent.ts`. Introduce `ActionConfirmationCard` in the chat UI for mass mutations (>10 cells) and add quick action presets in `QuickActionPresets.tsx`.

**Tech Stack:** TypeScript, React 18, Office.js (Excel), Vitest, Tailwind CSS, Lucide icons.

**Spec:** [docs/superpowers/specs/2026-09-22-excel-deep-intelligence-safe-execution-design.md](file:///E:/VIBE-CODE-WS/Sam-Office-Agent/docs/superpowers/specs/2026-09-22-excel-deep-intelligence-safe-execution-design.md)

## Global Constraints

- Zero external backend relays: all execution runs in-process or via local server (`localhost:5173`).
- Single-sync batching for all Office.js Excel operations via `await context.sync()`.
- Backward compatibility: All existing 252 tests must remain passing.
- TDD workflow: Write tests before implementation for every task.

---

### Task 1: Type Definitions & Driver Interfaces for Audit & Storytelling

**Files:**
- Modify: `src/services/office/types.ts`
- Modify: `tests/types.test.ts`

**Interfaces:**
- Produces:
  - `FormulaIssue`: `{ address: string; type: 'formula_error' | 'inconsistent_formula' | 'hardcoded_override' | 'suspicious_blank' | 'statistical_outlier'; severity: 'critical' | 'warning' | 'info'; currentValue?: any; formula?: string; expectedPattern?: string; suggestion: string; }`
  - `SheetAuditResult`: `{ sheetName: string; totalCellsAudited: number; totalErrorsFound: number; criticalIssues: FormulaIssue[]; warnings: FormulaIssue[]; summary: string; }`
  - `DataStoryOptions`: `{ range?: string; focusMetric?: string; includeRecommendations?: boolean; }`
  - `DataStoryMetric`: `{ label: string; value: string | number; changePercent?: number; trend?: 'up' | 'down' | 'neutral'; }`
  - `DataStoryResult`: `{ headline: string; keyFindings: string[]; metrics: DataStoryMetric[]; risksOrAnomalies?: string[]; recommendations: string[]; }`
  - `PendingActionProposal`: `{ id: string; actionType: 'modify_cells' | 'clean_data' | 'format_cells'; description: string; affectedCellsCount: number; targetRange?: string; payload: any; }`
  - Signatures in `IOfficeDriver`:
    - `auditSheetData?(options?: { range?: string }): Promise<SheetAuditResult>;`
    - `generateDataStory?(options?: DataStoryOptions): Promise<DataStoryResult>;`

- [ ] **Step 1: Write the failing test**

In `tests/types.test.ts`:
```typescript
import { FormulaIssue, SheetAuditResult, DataStoryResult, PendingActionProposal } from '../src/services/office/types';

describe('Excel Deep Intelligence Types', () => {
  it('validates FormulaIssue and SheetAuditResult structure', () => {
    const issue: FormulaIssue = {
      address: 'C5',
      type: 'formula_error',
      severity: 'critical',
      formula: '=A5/B5',
      currentValue: '#DIV/0!',
      suggestion: 'B5 bernilai 0, gunakan IFERROR(A5/B5, 0)',
    };
    const audit: SheetAuditResult = {
      sheetName: 'Sheet1',
      totalCellsAudited: 50,
      totalErrorsFound: 1,
      criticalIssues: [issue],
      warnings: [],
      summary: 'Ditemukan 1 error kritis pada sheet Sheet1.',
    };
    expect(audit.totalErrorsFound).toBe(1);
    expect(audit.criticalIssues[0].type).toBe('formula_error');
  });

  it('validates DataStoryResult and PendingActionProposal structure', () => {
    const story: DataStoryResult = {
      headline: 'Pendapatan Q3 Meningkat 24%',
      keyFindings: ['Penjualan produk X naik tajam'],
      metrics: [{ label: 'Total Revenue', value: 'Rp 1.250.000.000', changePercent: 24, trend: 'up' }],
      recommendations: ['Tingkatkan alokasi stok untuk produk X'],
    };
    const action: PendingActionProposal = {
      id: 'act-1',
      actionType: 'modify_cells',
      description: 'Perbarui 20 formula di kolom Total',
      affectedCellsCount: 20,
      targetRange: 'D2:D21',
      payload: { formula: '=B2*C2' },
    };
    expect(story.headline).toContain('Pendapatan');
    expect(action.affectedCellsCount).toBe(20);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/types.test.ts`
Expected: FAIL with TS compilation / missing export error.

- [ ] **Step 3: Add type definitions in `src/services/office/types.ts`**

Add `FormulaIssue`, `SheetAuditResult`, `DataStoryOptions`, `DataStoryMetric`, `DataStoryResult`, and `PendingActionProposal`. Update `IOfficeDriver` with the new optional methods.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/types.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/services/office/types.ts tests/types.test.ts
git commit -m "feat(types): add FormulaIssue, SheetAuditResult, DataStoryResult, and PendingActionProposal interfaces"
```

---

### Task 2: MockDriver & ExcelDriver Implementation for Audit & Data Story

**Files:**
- Modify: `src/services/office/mockDriver.ts`
- Modify: `src/services/office/excelDriver.ts`
- Modify: `src/services/office/__tests__/mockDriver.test.ts`

**Interfaces:**
- Consumes: Types from `src/services/office/types.ts`.
- Produces: `auditSheetData` and `generateDataStory` methods implemented in `MockOfficeDriver` and `ExcelDriver`.

- [ ] **Step 1: Write the failing tests**

In `src/services/office/__tests__/mockDriver.test.ts`:
```typescript
it('audits sheet data and detects formula errors and inconsistent overrides', async () => {
  const driver = new MockOfficeDriver();
  // seed mock data with error and inconsistent value
  driver.mockData = [
    ['Header1', 'Header2', 'Total'],
    [10, 20, 200],
    [15, 0, '#DIV/0!'],
    [20, 5, 100],
    [25, 4, 'manual_text_instead_of_number'],
  ];

  const audit = await driver.auditSheetData!();
  expect(audit.totalCellsAudited).toBeGreaterThan(0);
  expect(audit.totalErrorsFound).toBeGreaterThanOrEqual(1);
  const divZero = audit.criticalIssues.find(i => i.currentValue === '#DIV/0!');
  expect(divZero).toBeDefined();
  expect(divZero?.address).toBe('C3');
});

it('generates executive data story with metrics and recommendations', async () => {
  const driver = new MockOfficeDriver();
  driver.mockData = [
    ['Wilayah', 'Penjualan Q1', 'Penjualan Q2'],
    ['Barat', 1000, 1200],
    ['Timur', 800, 1100],
    ['Pusat', 1500, 1400],
  ];

  const story = await driver.generateDataStory!({ focusMetric: 'Penjualan Q2' });
  expect(story.headline).toBeDefined();
  expect(story.keyFindings.length).toBeGreaterThan(0);
  expect(story.metrics.length).toBeGreaterThan(0);
  expect(story.recommendations.length).toBeGreaterThan(0);
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/services/office/__tests__/mockDriver.test.ts`
Expected: FAIL with `driver.auditSheetData is not a function`.

- [ ] **Step 3: Implement methods in `mockDriver.ts` and `excelDriver.ts`**

In `mockDriver.ts`:
- Scan `this.mockData` grid for strings matching `/^#(REF!|VALUE!|DIV\/0!|N\/A|NAME\?|NUM!|NULL!)/i`.
- Check column data types: if a column is predominantly numeric, flag non-numeric values as warnings (`hardcoded_override` or `suspicious_blank`).
- Synthesize `DataStoryResult` with statistics (sum, average, max contributor, trend percentage) and Indonesian business narratives.

In `excelDriver.ts`:
- Use `context.workbook.worksheets.getActiveWorksheet()`.
- Load `usedRange.formulas`, `usedRange.values`, `usedRange.valueTypes`.
- Perform single `await context.sync()`.
- Run identical analytical evaluation algorithm and return structured results.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/services/office/__tests__/mockDriver.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/services/office/mockDriver.ts src/services/office/excelDriver.ts src/services/office/__tests__/mockDriver.test.ts
git commit -m "feat(driver): implement auditSheetData and generateDataStory in ExcelDriver and MockOfficeDriver"
```

---

### Task 3: Specialist Agent Tools for Excel (`audit_sheet_data` & `generate_data_story`)

**Files:**
- Modify: `src/agents/excel/excelAgent.ts`
- Modify: `src/agents/__tests__/specialistAgents.test.ts`

**Interfaces:**
- Consumes: Driver methods from `IOfficeDriver`.
- Produces: Native tools `audit_sheet_data` and `generate_data_story` available in `excelAgent.getTools()` and callable via `excelAgent.executeTool()`.

- [ ] **Step 1: Write the failing tests**

In `src/agents/__tests__/specialistAgents.test.ts`:
```typescript
it('provides audit_sheet_data tool and executes it successfully', async () => {
  const driver = new MockOfficeDriver();
  driver.mockData = [
    ['Product', 'Qty', 'Price', 'Total'],
    ['A', 10, 5, 50],
    ['B', 2, 0, '#DIV/0!'],
  ];
  const agent = new ExcelAgent(driver);
  const tools = agent.getTools();
  expect(tools.some(t => t.name === 'audit_sheet_data')).toBe(true);

  const res = await agent.executeTool(
    { id: 'c1', name: 'audit_sheet_data', arguments: {} },
    mockContext
  );
  expect(res.success).toBe(true);
  expect(res.result).toContain('Total Error Formula');
});

it('provides generate_data_story tool and executes it successfully', async () => {
  const driver = new MockOfficeDriver();
  driver.mockData = [
    ['Category', 'Revenue'],
    ['Laptops', 50000000],
    ['Phones', 75000000],
  ];
  const agent = new ExcelAgent(driver);
  const tools = agent.getTools();
  expect(tools.some(t => t.name === 'generate_data_story')).toBe(true);

  const res = await agent.executeTool(
    { id: 'c2', name: 'generate_data_story', arguments: { focusMetric: 'Revenue' } },
    mockContext
  );
  expect(res.success).toBe(true);
  expect(res.result).toContain('Headline');
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`
Expected: FAIL with `tool audit_sheet_data not found`.

- [ ] **Step 3: Register and handle tools in `src/agents/excel/excelAgent.ts`**

Register `audit_sheet_data` and `generate_data_story` in `getTools()` with parameter definitions. Add handlers in `executeTool()` that call `driver.auditSheetData` and `driver.generateDataStory` and format markdown output.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/agents/excel/excelAgent.ts src/agents/__tests__/specialistAgents.test.ts
git commit -m "feat(agent): register audit_sheet_data and generate_data_story tools in ExcelAgent"
```

---

### Task 4: Safe Dry-Run Proposal & ActionConfirmationCard Component

**Files:**
- Create: `src/components/Chat/ActionConfirmationCard.tsx`
- Modify: `src/components/Chat/MessageBubble.tsx`
- Modify: `src/types/index.ts`
- Modify: `src/components/__tests__/chatAndActions.test.tsx`

**Interfaces:**
- Produces: `ActionConfirmationCard` component taking `{ proposal: PendingActionProposal; onConfirm: () => void; onCancel: () => void; disabled?: boolean }`.
- Integrates into `MessageBubble.tsx` when a message contains `pendingAction`.

- [ ] **Step 1: Write the failing test**

In `src/components/__tests__/chatAndActions.test.tsx`:
```typescript
it('renders ActionConfirmationCard with affected cells and buttons', () => {
  const proposal: PendingActionProposal = {
    id: 'prop-1',
    actionType: 'modify_cells',
    description: 'Akan memperbarui formula di 25 sel pada kolom Total',
    affectedCellsCount: 25,
    targetRange: 'D2:D26',
    payload: {},
  };
  const onConfirm = vi.fn();
  const onCancel = vi.fn();

  render(
    <ActionConfirmationCard
      proposal={proposal}
      onConfirm={onConfirm}
      onCancel={onCancel}
    />
  );

  expect(screen.getByText(/Konfirmasi Perubahan/i)).toBeInTheDocument();
  expect(screen.getByText(/25 sel/i)).toBeInTheDocument();
  fireEvent.click(screen.getByRole('button', { name: /Terapkan/i }));
  expect(onConfirm).toHaveBeenCalled();
  fireEvent.click(screen.getByRole('button', { name: /Batalkan/i }));
  expect(onCancel).toHaveBeenCalled();
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/components/__tests__/chatAndActions.test.tsx`
Expected: FAIL with component not found.

- [ ] **Step 3: Create `ActionConfirmationCard.tsx` and integrate into `MessageBubble.tsx`**

Implement `ActionConfirmationCard.tsx` with clean alert styling, affected cell badges, and primary/secondary confirmation buttons. Attach to `MessageBubble.tsx` if `message.pendingAction` is present.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/components/__tests__/chatAndActions.test.tsx`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/components/Chat/ActionConfirmationCard.tsx src/components/Chat/MessageBubble.tsx src/types/index.ts src/components/__tests__/chatAndActions.test.tsx
git commit -m "feat(ui): implement ActionConfirmationCard for safe dry-run preview and confirmation"
```

---

### Task 5: QuickActionPresets Toolbar Expansion & ReAct Loop Integration

**Files:**
- Modify: `src/components/Chat/QuickActionPresets.tsx`
- Modify: `src/App.tsx`
- Modify: `src/components/__tests__/chatAndActions.test.tsx`
- Modify: `tests/app.test.tsx`

**Interfaces:**
- Produces: Updated Excel presets with `"🔍 Audit Formula & Error"` and `"📈 Buat Analisis Tren"`.
- Wires confirmation handling in `App.tsx` (`handleConfirmAction` and `handleCancelAction`).

- [ ] **Step 1: Write the failing test**

In `src/components/__tests__/chatAndActions.test.tsx`:
```typescript
it('renders new Excel presets for Audit Formula and Analisis Tren', () => {
  const onSelect = vi.fn();
  render(<QuickActionPresets host="Excel" onSelectPreset={onSelect} />);

  expect(screen.getByText(/Audit Formula/i)).toBeInTheDocument();
  expect(screen.getByText(/Analisis Tren/i)).toBeInTheDocument();

  fireEvent.click(screen.getByText(/Audit Formula/i));
  expect(onSelect).toHaveBeenCalledWith(expect.stringContaining('Audit seluruh formula'));
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/components/__tests__/chatAndActions.test.tsx`
Expected: FAIL with preset button not found.

- [ ] **Step 3: Implement QuickActionPresets additions and App.tsx confirmation callbacks**

In `QuickActionPresets.tsx`:
Add presets:
- `"🔍 Audit Formula & Error"` -> `"Audit seluruh formula pada sheet ini: cari error #REF!, #VALUE!, #DIV/0!, atau sel dengan rumus tidak konsisten dan berikan rekomendasinya."`
- `"📈 Buat Analisis Tren"` -> `"Analisis data pada tabel aktif ini: buatkan ringkasan eksekutif, tren pertumbuhan, dan insight temuan utama."`

In `App.tsx`:
Add `handleConfirmAction(actionId: string)` and `handleCancelAction(actionId: string)` state and handler to dispatch pending tool operations.

- [ ] **Step 4: Run all tests to verify 100% pass**

Run: `npm test -- --run`
Expected: All 255+ tests PASS.

- [ ] **Step 5: Run production build**

Run: `npm run build`
Expected: Zero TypeScript or Vite compilation errors.

- [ ] **Step 6: Commit**

```bash
git add src/components/Chat/QuickActionPresets.tsx src/App.tsx src/components/__tests__/chatAndActions.test.tsx tests/app.test.tsx
git commit -m "feat(ui): add Audit and Data Story presets to toolbar and wire confirmation handlers in App"
```
