# Cross-App Universal Hub & Data Bridge Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement the Sam Universal Hub and Cross-App Data Bridge, enabling seamless sharing and automated translation of tables, executive summaries, and metrics across Excel, Word, and PowerPoint with a proactive Smart Context Card UI.

**Architecture:** Build a client-side persistent storage engine (`crossAppBridge.ts`) that manages `CrossAppSnapshot` records in `localStorage`. Equip `ExcelAgent`, `WordAgent`, and `PPTAgent` with `share_to_cross_app_hub` and `import_from_cross_app_hub` tools. Provide a proactive `CrossAppSnapshotCard.tsx` component in the chat UI along with 1-click Quick Action Presets.

**Tech Stack:** TypeScript, React 18, Vitest, Tailwind CSS, Lucide icons.

**Spec:** [docs/superpowers/specs/2026-09-23-cross-app-universal-hub-design.md](file:///E:/VIBE-CODE-WS/Sam-Office-Agent/docs/superpowers/specs/2026-09-23-cross-app-universal-hub-design.md)

## Global Constraints

- 100% Client-Side BYOK: All cross-app state is stored in origin-scoped local storage (`localStorage`).
- Zero external backend relays or telemetries.
- TDD workflow: Write tests before implementation for every task.
- Backward compatibility: All 283 existing tests must remain passing.

---

### Task 1: Cross-App Data Types & Bridge Storage Engine (`crossAppBridge.ts`)

**Files:**
- Modify: `src/types/index.ts`
- Create: `src/services/storage/crossAppBridge.ts`
- Create: `tests/crossAppBridge.test.ts`

**Interfaces:**
- Produces:
  - Types: `CrossAppArtifactType`, `CrossAppTableData`, `CrossAppMetric`, `CrossAppSnapshot`.
  - Functions: `saveCrossAppSnapshot`, `getLatestCrossAppSnapshot`, `listCrossAppSnapshots`, `dismissCrossAppSnapshot`, `isCrossAppSnapshotDismissed`, `clearCrossAppSnapshots`.

- [x] **Step 1: Write the failing test**

In `tests/crossAppBridge.test.ts`:
```typescript
import { describe, it, expect, beforeEach } from 'vitest';
import {
  saveCrossAppSnapshot,
  getLatestCrossAppSnapshot,
  listCrossAppSnapshots,
  dismissCrossAppSnapshot,
  isCrossAppSnapshotDismissed,
  clearCrossAppSnapshots,
} from '../src/services/storage/crossAppBridge';

describe('CrossAppBridge Storage Engine', () => {
  beforeEach(() => {
    clearCrossAppSnapshots();
    localStorage.clear();
  });

  it('saves and retrieves latest cross-app snapshot excluding active host', () => {
    saveCrossAppSnapshot({
      sourceHost: 'Excel',
      title: 'Rekap Penjualan Q3',
      artifactType: 'table_data',
      tableData: {
        headers: ['Wilayah', 'Target', 'Realisasi'],
        rows: [['Barat', 1000, 1200]],
        totalRows: 1,
      },
      summaryText: 'Realisasi wilayah barat melampaui target 120%.',
    });

    // In Excel: excludeHost='Excel' should return null
    expect(getLatestCrossAppSnapshot('Excel')).toBeNull();

    // In Word: excludeHost='Word' should find Excel's snapshot
    const wordView = getLatestCrossAppSnapshot('Word');
    expect(wordView).not.toBeNull();
    expect(wordView?.title).toBe('Rekap Penjualan Q3');
    expect(wordView?.sourceHost).toBe('Excel');
  });

  it('supports dismissing a snapshot', () => {
    const snap = saveCrossAppSnapshot({
      sourceHost: 'Excel',
      title: 'Data Uji',
      artifactType: 'table_data',
    });

    expect(isCrossAppSnapshotDismissed(snap.id)).toBe(false);
    dismissCrossAppSnapshot(snap.id);
    expect(isCrossAppSnapshotDismissed(snap.id)).toBe(true);

    // Dismissed snapshot should no longer be returned as latest
    expect(getLatestCrossAppSnapshot('Word')).toBeNull();
  });
});
```

- [x] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/crossAppBridge.test.ts`
Expected: FAIL with module not found.

- [x] **Step 3: Implement `src/types/index.ts` and `src/services/storage/crossAppBridge.ts`**

Export interfaces and implement localStorage serialization with TTL filter (24 hours) and dismissal tracking.

- [x] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/crossAppBridge.test.ts`
Expected: PASS.

- [x] **Step 5: Commit**

```bash
git add src/types/index.ts src/services/storage/crossAppBridge.ts tests/crossAppBridge.test.ts
git commit -m "feat(storage): implement CrossAppBridge engine and contracts"
```

---

### Task 2: Excel Specialist Tool: `share_to_cross_app_hub`

**Files:**
- Modify: `src/agents/excel/excelAgent.ts`
- Modify: `src/agents/__tests__/specialistAgents.test.ts`

**Interfaces:**
- Consumes: `saveCrossAppSnapshot` from `crossAppBridge.ts`.
- Produces: Tool `share_to_cross_app_hub` in `ExcelAgent.getTools()` and handler in `executeTool()`.

- [x] **Step 1: Write the failing test**

In `src/agents/__tests__/specialistAgents.test.ts`:
```typescript
it('provides share_to_cross_app_hub tool in ExcelAgent and saves table snapshot', async () => {
  const mockDriver = new MockOfficeDriver();
  mockDriver.mockData = [
    ['Kategori', 'Q1', 'Q2'],
    ['Elektronik', 500, 700],
    ['Pakaian', 300, 450],
  ];
  const testAgent = new ExcelAgent(mockDriver);
  const tools = testAgent.getTools();
  expect(tools.some(t => t.name === 'share_to_cross_app_hub')).toBe(true);

  const res = await testAgent.executeTool(
    {
      id: 'ex-hub-1',
      name: 'share_to_cross_app_hub',
      arguments: { title: 'Penjualan Semester 1' },
      status: 'pending',
    },
    { host: 'Excel' }
  );

  expect(res.success).toBe(true);
  expect(res.result).toContain('Universal Hub');
  expect(res.result).toContain('Penjualan Semester 1');
});
```

- [x] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`
Expected: FAIL with tool not found.

- [x] **Step 3: Implement `share_to_cross_app_hub` in `src/agents/excel/excelAgent.ts`**

Register the tool and implement execution: read sheet data, format table payload, call `saveCrossAppSnapshot`, and return confirmation.

- [x] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`
Expected: PASS.

- [x] **Step 5: Commit**

```bash
git add src/agents/excel/excelAgent.ts src/agents/__tests__/specialistAgents.test.ts
git commit -m "feat(agent): implement share_to_cross_app_hub in ExcelAgent"
```

---

### Task 3: Word Specialist Tools: `import_from_cross_app_hub` & `share_to_cross_app_hub`

**Files:**
- Modify: `src/agents/word/wordAgent.ts`
- Modify: `src/agents/__tests__/specialistAgents.test.ts`

**Interfaces:**
- Consumes: `getLatestCrossAppSnapshot`, `saveCrossAppSnapshot`.
- Produces: Tools `import_from_cross_app_hub` and `share_to_cross_app_hub` in `WordAgent`.

- [x] **Step 1: Write the failing test**

In `src/agents/__tests__/specialistAgents.test.ts`:
```typescript
it('provides import_from_cross_app_hub in WordAgent and inserts table and narrative', async () => {
  saveCrossAppSnapshot({
    sourceHost: 'Excel',
    title: 'Data Operasional',
    artifactType: 'table_data',
    tableData: {
      headers: ['Unit', 'Skor'],
      rows: [['A', 90]],
      totalRows: 1,
    },
    summaryText: 'Unit A meraih performa optimal.',
  });

  const mockDriver = new MockOfficeDriver();
  const testAgent = new WordAgent(mockDriver);
  const tools = testAgent.getTools();
  expect(tools.some(t => t.name === 'import_from_cross_app_hub')).toBe(true);

  const res = await testAgent.executeTool(
    {
      id: 'w-hub-1',
      name: 'import_from_cross_app_hub',
      arguments: {},
      status: 'pending',
    },
    { host: 'Word' }
  );

  expect(res.success).toBe(true);
  expect(res.result).toContain('Data Operasional');
  expect(mockDriver.wordContent.join('\n')).toContain('Data Operasional');
});
```

- [x] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`
Expected: FAIL.

- [x] **Step 3: Implement `import_from_cross_app_hub` and `share_to_cross_app_hub` in `src/agents/word/wordAgent.ts`**

Register tools and implement execution calling `driver.insertTable` and `driver.insertContent`.

- [x] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`
Expected: PASS.

- [x] **Step 5: Commit**

```bash
git add src/agents/word/wordAgent.ts src/agents/__tests__/specialistAgents.test.ts
git commit -m "feat(agent): implement cross-app hub import and share in WordAgent"
```

---

### Task 4: PowerPoint Specialist Tool: `import_from_cross_app_hub`

**Files:**
- Modify: `src/agents/powerpoint/pptAgent.ts`
- Modify: `src/agents/__tests__/specialistAgents.test.ts`

**Interfaces:**
- Consumes: `getLatestCrossAppSnapshot`, `driver.transformDocToDeck`.
- Produces: Tool `import_from_cross_app_hub` in `PPTAgent`.

- [x] **Step 1: Write the failing test**

In `src/agents/__tests__/specialistAgents.test.ts`:
```typescript
it('provides import_from_cross_app_hub in PPTAgent and transforms snapshot into slide deck', async () => {
  saveCrossAppSnapshot({
    sourceHost: 'Excel',
    title: 'Ringkasan Kinerja Tahunan',
    artifactType: 'table_data',
    tableData: {
      headers: ['Divisi', 'Capaian'],
      rows: [['Pemasaran', '110%']],
      totalRows: 1,
    },
    summaryText: 'Tantangan: Keterbatasan anggaran. Solusi: Optimalisasi saluran digital. Metrik: Capaian 110%. Rencana: Ekspansi pasar Q1.',
  });

  const mockDriver = new MockOfficeDriver();
  const testAgent = new PPTAgent(mockDriver);
  const tools = testAgent.getTools();
  expect(tools.some(t => t.name === 'import_from_cross_app_hub')).toBe(true);

  const res = await testAgent.executeTool(
    {
      id: 'ppt-hub-1',
      name: 'import_from_cross_app_hub',
      arguments: { theme: 'corporate_blue' },
      status: 'pending',
    },
    { host: 'PowerPoint' }
  );

  expect(res.success).toBe(true);
  expect(res.result).toContain('Ringkasan Kinerja Tahunan');
  expect(mockDriver.slides.length).toBeGreaterThanOrEqual(4);
});
```

- [x] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`
Expected: FAIL.

- [x] **Step 3: Implement `import_from_cross_app_hub` in `src/agents/powerpoint/pptAgent.ts`**

Register the tool and implement execution by pulling the latest snapshot, assembling formatted narrative, and executing `driver.transformDocToDeck`.

- [x] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`
Expected: PASS.

- [x] **Step 5: Commit**

```bash
git add src/agents/powerpoint/pptAgent.ts src/agents/__tests__/specialistAgents.test.ts
git commit -m "feat(agent): implement import_from_cross_app_hub in PPTAgent"
```

---

### Task 5: Proactive UI Component (`CrossAppSnapshotCard.tsx`), Presets & Full Verification

**Files:**
- Create: `src/components/Chat/CrossAppSnapshotCard.tsx`
- Modify: `src/components/Chat/ChatContainer.tsx`
- Modify: `src/components/Chat/QuickActionPresets.tsx`
- Modify: `src/components/__tests__/chatAndActions.test.tsx`

**Interfaces:**
- Produces: `CrossAppSnapshotCard` rendered when a fresh cross-app snapshot is present, with context-appropriate action buttons and dismissal.

- [ ] **Step 1: Write the failing test**

In `src/components/__tests__/chatAndActions.test.tsx`:
```typescript
it('renders CrossAppSnapshotCard when fresh snapshot from another host exists', () => {
  saveCrossAppSnapshot({
    sourceHost: 'Excel',
    title: 'Rekap Penjualan Q3',
    artifactType: 'table_data',
    summaryText: 'Tabel 15 baris',
  });

  render(
    <ChatContainer
      host="Word"
      messages={[]}
      onSendMessage={vi.fn()}
      onApplyToolCall={vi.fn()}
    />
  );

  expect(screen.getByText(/Rekap Penjualan Q3/i)).toBeInTheDocument();
  expect(screen.getByText(/Ditemukan data dari Excel/i)).toBeInTheDocument();
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/components/__tests__/chatAndActions.test.tsx`
Expected: FAIL.

- [ ] **Step 3: Implement `CrossAppSnapshotCard.tsx` and integrate into `ChatContainer.tsx` and `QuickActionPresets.tsx`**

- Create `CrossAppSnapshotCard.tsx` with host badges (Excel/Word/PPT), summary details, action buttons, and dismiss handler.
- Mount `CrossAppSnapshotCard` in `ChatContainer.tsx` above the message stream / input bar.
- Add cross-app preset buttons to `PRESETS_BY_HOST` for Excel, Word, and PowerPoint.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/components/__tests__/chatAndActions.test.tsx`
Expected: PASS.

- [ ] **Step 5: Run full test suite to ensure 0 regressions**

Run: `npx vitest run`
Expected: All test files pass with 0 failures.

- [ ] **Step 6: Commit**

```bash
git add src/components/Chat/CrossAppSnapshotCard.tsx src/components/Chat/ChatContainer.tsx src/components/Chat/QuickActionPresets.tsx src/components/__tests__/chatAndActions.test.tsx
git commit -m "feat(ui): add proactive CrossAppSnapshotCard and Hub quick action presets"
```
