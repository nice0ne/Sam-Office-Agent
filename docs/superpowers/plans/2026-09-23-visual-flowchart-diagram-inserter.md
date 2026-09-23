# Visual Flowchart & Process Diagram Inserter Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Provide an autonomous, zero-dependency 2D Canvas flowchart and process diagram engine that converts natural language SOPs or process descriptions into crisp Base64 PNG diagrams and inserts them directly into Word documents and PowerPoint slides.

**Architecture:** Pure client-side HTML5 Canvas 2D layout calculation and image synthesis (`src/utils/diagramRenderer.ts`) without external NPM dependencies. Extended Office drivers (`WordDriver`, `PPTDriver`, `MockOfficeDriver`) insert the resulting Base64 PNG via native Office.js APIs. Registered ReAct tool (`insert_process_flowchart`) in `WordAgent` and `PPTAgent` coordinates natural language text extraction and document insertion.

**Tech Stack:** TypeScript, HTML5 Canvas 2D Context, Office.js (Word & PowerPoint APIs), Vitest, React.

**Spec:** `docs/superpowers/specs/2026-09-23-visual-flowchart-diagram-inserter-design.md`

## Global Constraints
- Zero external charting or diagramming npm dependencies (no mermaid.js, cytoscape, or d3).
- 100% client-side rendering; no cloud rendering APIs or remote server calls.
- High-DPI crisp export (2x scale) with stripped Base64 string for direct Office.js compatibility.
- Seamless fallback when canvas is unavailable in headless environments.
- 100% test coverage and 0 test regressions on `npm run build` and `npx vitest run`.

---

### Task 1: Flowchart & Diagram Type Contracts

**Files:**
- Modify: `src/services/office/types.ts`
- Test: `tests/types.test.ts`

**Interfaces:**
- Produces:
  - `FlowchartNodeType` ('start' | 'end' | 'process' | 'decision' | 'document' | 'subroutine')
  - `FlowchartNode` ({ id: string; label: string; type: FlowchartNodeType; subText?: string; color?: string; })
  - `FlowchartEdge` ({ from: string; to: string; label?: string; style?: 'solid' | 'dashed'; })
  - `FlowchartDefinition` ({ title?: string; direction?: 'TD' | 'LR'; theme?: string; nodes: FlowchartNode[]; edges: FlowchartEdge[]; })
  - `FlowchartOptions` ({ title?: string; theme?: string; direction?: 'TD' | 'LR'; scale?: number; })
  - `InsertFlowchartOptions` ({ textOrSteps?: string; definition?: FlowchartDefinition; caption?: string; title?: string; theme?: string; direction?: 'TD' | 'LR'; })
  - `DiagramResult` ({ title: string; nodeCount: number; edgeCount: number; appliedTheme: string; base64Png: string; inserted: boolean; })
  - `IOfficeDriver.insertProcessFlowchart?(options: InsertFlowchartOptions): Promise<DiagramResult>`
  - `IDocumentDriver.insertProcessFlowchart?(options: InsertFlowchartOptions): Promise<DiagramResult>`

- [ ] **Step 1: Write the failing test**

In `tests/types.test.ts`, add a test verifying the new diagram types and contracts:

```typescript
import {
  FlowchartNodeType,
  FlowchartNode,
  FlowchartEdge,
  FlowchartDefinition,
  InsertFlowchartOptions,
  DiagramResult,
} from '../src/services/office/types';

describe('Diagram & Flowchart Type Contracts', () => {
  it('validates FlowchartDefinition and DiagramResult structure', () => {
    const node: FlowchartNode = {
      id: 'step-1',
      label: 'Pengajuan Cuti',
      type: 'start',
      subText: 'Karyawan mengisi form',
    };

    const edge: FlowchartEdge = {
      from: 'step-1',
      to: 'step-2',
      label: 'Kirim',
      style: 'solid',
    };

    const def: FlowchartDefinition = {
      title: 'SOP Cuti Karyawan',
      direction: 'TD',
      theme: 'corporate_navy',
      nodes: [node],
      edges: [edge],
    };

    const result: DiagramResult = {
      title: def.title!,
      nodeCount: def.nodes.length,
      edgeCount: def.edges.length,
      appliedTheme: def.theme!,
      base64Png: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
      inserted: true,
    };

    expect(result.nodeCount).toBe(1);
    expect(result.edgeCount).toBe(1);
    expect(result.appliedTheme).toBe('corporate_navy');
    expect(result.inserted).toBe(true);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/types.test.ts`
Expected: FAIL with "Module '.../types' has no exported member 'FlowchartNodeType'".

- [ ] **Step 3: Write minimal implementation**

In `src/services/office/types.ts`, export the types and extend `IOfficeDriver` & `IDocumentDriver`:

```typescript
export type FlowchartNodeType = 'start' | 'end' | 'process' | 'decision' | 'document' | 'subroutine';

export interface FlowchartNode {
  id: string;
  label: string;
  type: FlowchartNodeType;
  subText?: string;
  color?: string;
}

export interface FlowchartEdge {
  from: string;
  to: string;
  label?: string;
  style?: 'solid' | 'dashed';
}

export interface FlowchartDefinition {
  title?: string;
  direction?: 'TD' | 'LR';
  theme?: 'corporate_navy' | 'emerald_executive' | 'modern_dark' | 'amber_warm';
  nodes: FlowchartNode[];
  edges: FlowchartEdge[];
}

export interface FlowchartOptions {
  title?: string;
  theme?: 'corporate_navy' | 'emerald_executive' | 'modern_dark' | 'amber_warm';
  direction?: 'TD' | 'LR';
  scale?: number;
}

export interface InsertFlowchartOptions extends FlowchartOptions {
  textOrSteps?: string;
  definition?: FlowchartDefinition;
  caption?: string;
}

export interface DiagramResult {
  title: string;
  nodeCount: number;
  edgeCount: number;
  appliedTheme: string;
  base64Png: string;
  inserted: boolean;
}
```
And inside `IOfficeDriver` and `IDocumentDriver`, add:
```typescript
insertProcessFlowchart?(options: InsertFlowchartOptions): Promise<DiagramResult>;
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/types.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/services/office/types.ts tests/types.test.ts
git commit -m "feat(types): add flowchart and diagram contracts to office types"
```

---

### Task 2: Core Diagram Layout & Canvas Renderer (`src/utils/diagramRenderer.ts`)

**Files:**
- Create: `src/utils/diagramRenderer.ts`
- Create: `src/utils/__tests__/diagramRenderer.test.ts`

**Interfaces:**
- Consumes: `FlowchartNode`, `FlowchartEdge`, `FlowchartDefinition`, `FlowchartOptions` from `../services/office/types`
- Produces:
  - `synthesizeFlowchartFromText(rawText?: string, options?: FlowchartOptions): FlowchartDefinition`
  - `renderFlowchartToPngBase64(def: FlowchartDefinition, options?: FlowchartOptions): string`

- [ ] **Step 1: Write the failing test**

Create `src/utils/__tests__/diagramRenderer.test.ts`:

```typescript
import { describe, it, expect } from 'vitest';
import { synthesizeFlowchartFromText, renderFlowchartToPngBase64 } from '../diagramRenderer';
import { FlowchartDefinition } from '../../services/office/types';

describe('diagramRenderer', () => {
  it('synthesizes flowchart definition from structured SOP text', () => {
    const rawSOP = `
    1. Mulai: Karyawan mengajukan permohonan cuti tahunan melalui portal.
    2. Apakah sisa kuota cuti mencukupi? Jika tidak, tolak pengajuan.
    3. Manajer memvalidasi dan menyetujui jadwal cuti.
    4. HRD mencatat dokumen dan memotong kuota cuti.
    5. Selesai: Karyawan menerima notifikasi persetujuan cuti.
    `;

    const def = synthesizeFlowchartFromText(rawSOP, {
      title: 'Alur Pengajuan Cuti',
      theme: 'corporate_navy',
    });

    expect(def.title).toBe('Alur Pengajuan Cuti');
    expect(def.nodes.length).toBeGreaterThanOrEqual(4);
    expect(def.edges.length).toBeGreaterThanOrEqual(3);

    // Node 0 should be start
    expect(def.nodes[0].type).toBe('start');
    // Decision node should be detected
    const decisionNode = def.nodes.find(n => n.type === 'decision');
    expect(decisionNode).toBeDefined();
    // End node should be detected
    const endNode = def.nodes.find(n => n.type === 'end');
    expect(endNode).toBeDefined();
  });

  it('synthesizes default flowchart when text is empty or minimal', () => {
    const def = synthesizeFlowchartFromText('');
    expect(def.nodes.length).toBeGreaterThan(0);
    expect(def.edges.length).toBeGreaterThan(0);
  });

  it('renders flowchart definition to clean Base64 PNG string', () => {
    const sampleDef: FlowchartDefinition = {
      title: 'Proses Approval Anggaran',
      direction: 'TD',
      theme: 'emerald_executive',
      nodes: [
        { id: '1', label: 'Input RAB', type: 'start' },
        { id: '2', label: 'Verifikasi Finance', type: 'decision' },
        { id: '3', label: 'Disetujui Direksi', type: 'end' },
      ],
      edges: [
        { from: '1', to: '2' },
        { from: '2', to: '3', label: 'Ya' },
      ],
    };

    const base64 = renderFlowchartToPngBase64(sampleDef);
    expect(typeof base64).toBe('string');
    expect(base64.length).toBeGreaterThan(20);
    expect(base64).not.toContain('data:image/png;base64,');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/utils/__tests__/diagramRenderer.test.ts`
Expected: FAIL with "Cannot find module '../diagramRenderer'".

- [ ] **Step 3: Write minimal implementation**

Create `src/utils/diagramRenderer.ts`:
- Define themes: `corporate_navy`, `emerald_executive`, `modern_dark`, `amber_warm` with primary, secondary, border, background, and text colors.
- Implement `synthesizeFlowchartFromText(rawText?: string, options?: FlowchartOptions): FlowchartDefinition`:
  - Splits input into lines/sentences.
  - Cleans numbering like `1.`, `Step 1:`, `-`.
  - Infers `FlowchartNodeType` using regex:
    - `/mulai|start|awal|input|ajukan/i` -> `'start'`
    - `/apakah|cek|validasi|review|verifikasi|setuju|approval|\?/i` -> `'decision'`
    - `/dokumen|berkas|formulir|laporan|surat/i` -> `'document'`
    - `/selesai|end|tuntas|kirim|arsip|terima/i` -> `'end'`
    - Else `'process'`.
  - Builds sequential edges connecting node `i` to `i + 1`. If node is `decision`, adds label `Ya` or conditional branch.
- Implement `renderFlowchartToPngBase64(def: FlowchartDefinition, options?: FlowchartOptions): string`:
  - Canvas dimensions (e.g., width 800-1000px, dynamic height based on levels).
  - Background fill and title text rendering.
  - Computes level layout (top-down or left-right).
  - Renders connecting arrows with filled arrowheads and edge labels.
  - Renders shapes:
    - `start` / `end`: Stadium/pill (`ctx.arc` / `roundRect`).
    - `decision`: Diamond polygon (`ctx.moveTo(cx, top); ctx.lineTo(right, cy)...`).
    - `document`: Curved bottom wave box.
    - `process`: Rounded rectangle.
  - Labels drawn centered with contrasting text.
  - Strips `data:image/png;base64,` and returns clean string.
  - Safe fallback to `FALLBACK_1X1_PNG` if canvas or context is unavailable.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/utils/__tests__/diagramRenderer.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/utils/diagramRenderer.ts src/utils/__tests__/diagramRenderer.test.ts
git commit -m "feat(diagram): implement zero-dependency canvas flowchart renderer and text synthesizer"
```

---

### Task 3: Office Drivers Integration (`mockDriver.ts`, `wordDriver.ts`, `pptDriver.ts`)

**Files:**
- Modify: `src/services/office/mockDriver.ts`
- Modify: `src/services/office/wordDriver.ts`
- Modify: `src/services/office/pptDriver.ts`
- Test: `src/services/office/__tests__/mockDriver.test.ts`

**Interfaces:**
- Consumes: `synthesizeFlowchartFromText`, `renderFlowchartToPngBase64` from `../../utils/diagramRenderer`
- Produces: `insertProcessFlowchart(options: InsertFlowchartOptions): Promise<DiagramResult>` across drivers

- [ ] **Step 1: Write the failing test**

In `src/services/office/__tests__/mockDriver.test.ts`, add a test for `insertProcessFlowchart`:

```typescript
it('inserts process flowchart and stores diagram result in mock driver', async () => {
  const driver = new MockOfficeDriver();
  const result = await driver.insertProcessFlowchart!({
    textOrSteps: '1. Pengajuan cuti -> 2. Approval atasan -> 3. Selesai',
    title: 'SOP Cuti Karyawan',
    theme: 'corporate_navy',
    caption: 'Gambar 1: Alur Pengajuan Cuti',
  });

  expect(result.inserted).toBe(true);
  expect(result.nodeCount).toBeGreaterThanOrEqual(3);
  expect(result.appliedTheme).toBe('corporate_navy');
  expect(result.base64Png).toBeDefined();
  expect(driver.diagrams.length).toBe(1);
  expect(driver.diagrams[0].title).toBe('SOP Cuti Karyawan');
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/services/office/__tests__/mockDriver.test.ts`
Expected: FAIL with "insertProcessFlowchart is not a function".

- [ ] **Step 3: Write minimal implementation**

1. In `src/services/office/mockDriver.ts`:
   - Add property `diagrams: DiagramResult[] = [];`
   - Implement `async insertProcessFlowchart(options: InsertFlowchartOptions): Promise<DiagramResult>`:
     - Call `synthesizeFlowchartFromText` or use `options.definition`.
     - Call `renderFlowchartToPngBase64`.
     - Construct `DiagramResult` with `inserted: true`.
     - Push to `this.diagrams`.
     - Return the result.
2. In `src/services/office/wordDriver.ts`:
   - Implement `async insertProcessFlowchart(options: InsertFlowchartOptions): Promise<DiagramResult>`:
     - Generate definition & base64 PNG using `diagramRenderer`.
     - Inside `Word.run`: call `context.document.body.insertInlinePictureFromBase64(base64, Word.InsertLocation.end)`.
     - If caption is present, insert caption paragraph below.
     - Call `await context.sync()`.
     - Return `DiagramResult`.
3. In `src/services/office/pptDriver.ts`:
   - Implement `async insertProcessFlowchart(options: InsertFlowchartOptions): Promise<DiagramResult>`:
     - Generate definition & base64 PNG using `diagramRenderer`.
     - Inside `PowerPoint.run`: add a dedicated slide with title and insert picture using Office PowerPoint shape/image API, or call `context.presentation.slides.add()`.
     - Return `DiagramResult`.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/services/office/__tests__/mockDriver.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/services/office/mockDriver.ts src/services/office/wordDriver.ts src/services/office/pptDriver.ts src/services/office/__tests__/mockDriver.test.ts
git commit -m "feat(driver): implement insertProcessFlowchart across mock, word, and ppt drivers"
```

---

### Task 4: Specialist Agent Tools (`wordAgent.ts` & `pptAgent.ts`)

**Files:**
- Modify: `src/agents/word/wordAgent.ts`
- Modify: `src/agents/powerpoint/pptAgent.ts`
- Modify: `src/agents/coordinator/samCoordinator.ts`
- Test: `tests/specialistAgents.test.ts`

**Interfaces:**
- Consumes: `driver.insertProcessFlowchart`
- Produces: `insert_process_flowchart` tool in `WordAgent` and `PPTAgent`

- [ ] **Step 1: Write the failing test**

In `tests/specialistAgents.test.ts`, add tests for `insert_process_flowchart`:

```typescript
it('provides insert_process_flowchart tool in WordAgent and executes successfully', async () => {
  const mockDriver = new MockOfficeDriver();
  const wordAgent = new WordAgent(mockDriver);

  const tools = wordAgent.getTools();
  expect(tools.some(t => t.name === 'insert_process_flowchart')).toBe(true);

  const res = await wordAgent.executeTool(
    {
      id: 'tool-flow-1',
      name: 'insert_process_flowchart',
      arguments: {
        textOrSteps: '1. Pengajuan berkas -> 2. Cek kelengkapan -> 3. Terbitkan surat',
        title: 'Alur Penerbitan Surat',
        theme: 'corporate_navy',
      },
      status: 'pending',
    },
    { host: 'Word' }
  );

  expect(res.success).toBe(true);
  expect(res.result).toContain('Diagram Alur Proses Berhasil Disisipkan');
  expect(res.result).toContain('Alur Penerbitan Surat');
  expect(mockDriver.diagrams.length).toBe(1);
});

it('provides insert_process_flowchart tool in PPTAgent and executes successfully', async () => {
  const mockDriver = new MockOfficeDriver();
  const pptAgent = new PPTAgent(mockDriver);

  const tools = pptAgent.getTools();
  expect(tools.some(t => t.name === 'insert_process_flowchart')).toBe(true);

  const res = await pptAgent.executeTool(
    {
      id: 'tool-flow-ppt-1',
      name: 'insert_process_flowchart',
      arguments: {
        textOrSteps: '1. Inisiasi Proyek -> 2. Analisis Kebutuhan -> 3. Eksekusi -> 4. Evaluasi Akhir',
        title: 'Siklus Proyek',
        theme: 'emerald_executive',
      },
      status: 'pending',
    },
    { host: 'PowerPoint' }
  );

  expect(res.success).toBe(true);
  expect(res.result).toContain('Diagram Alur Proses Berhasil Disisipkan');
  expect(mockDriver.diagrams.length).toBe(1);
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/specialistAgents.test.ts`
Expected: FAIL with "tools.some(...) is false" or "Unknown tool".

- [ ] **Step 3: Write minimal implementation**

1. In `src/agents/word/wordAgent.ts`:
   - Register `insert_process_flowchart` in `getTools()` with parameters (`textOrSteps`, `title`, `theme`, `direction`).
   - Handle in `executeTool()`:
     - Call `driver.insertProcessFlowchart(...)`.
     - Return formatted markdown card with node breakdown, edge flow, and confirmation of insertion into Word document.
   - Mention `insert_process_flowchart` in `getSystemPrompt()`.
2. In `src/agents/powerpoint/pptAgent.ts`:
   - Register `insert_process_flowchart` in `getTools()`.
   - Handle in `executeTool()`:
     - Call `driver.insertProcessFlowchart(...)`.
     - Return formatted markdown with slide creation confirmation and presenter speaking notes.
   - Mention in `getSystemPrompt()`.
3. In `src/agents/coordinator/samCoordinator.ts`:
   - In `buildSystemPrompt()`, add guideline for flowchart creation:
     `"7. DIAGRAM ALUR PROSES (FLOWCHART): Ketika pengguna meminta dibuatkan diagram alur, flowchart, SOP, atau visualisasi alur kerja di Word atau PowerPoint, gunakan tool insert_process_flowchart untuk menghasilkan diagram visual profesional dan langsung menyisipkannya ke dokumen/slide."`

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/specialistAgents.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/agents/word/wordAgent.ts src/agents/powerpoint/pptAgent.ts src/agents/coordinator/samCoordinator.ts tests/specialistAgents.test.ts
git commit -m "feat(agents): register and execute insert_process_flowchart tool in Word and PPT agents"
```

---

### Task 5: UI Quick Action Presets & Full Verification

**Files:**
- Modify: `src/components/Chat/QuickActionPresets.tsx`
- Test: `src/components/__tests__/chatAndActions.test.tsx`

**Interfaces:**
- Consumes: QuickActionPresets for Word and PowerPoint hosts
- Produces: User-facing quick action buttons for Flowcharts

- [ ] **Step 1: Write the failing test**

In `src/components/__tests__/chatAndActions.test.tsx`, add a test verifying the new flowchart presets:

```typescript
it('renders flowchart preset in Word and PowerPoint hosts', () => {
  const { rerender } = render(
    <QuickActionPresets host="Word" onSelectPreset={vi.fn()} />
  );
  expect(screen.getByText(/Buat Diagram Alur/i)).toBeInTheDocument();

  rerender(
    <QuickActionPresets host="PowerPoint" onSelectPreset={vi.fn()} />
  );
  expect(screen.getByText(/Slide Flowchart Alur/i)).toBeInTheDocument();
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/components/__tests__/chatAndActions.test.tsx`
Expected: FAIL with "Unable to find an element with text: /Buat Diagram Alur/i".

- [ ] **Step 3: Write minimal implementation**

In `src/components/Chat/QuickActionPresets.tsx`:
- Under `PRESETS_BY_HOST.Word`:
  ```typescript
  {
    id: 'word_insert_flowchart',
    label: '📐 Buat Diagram Alur (Flowchart)',
    prompt: 'Buat diagram alur proses visual (flowchart) profesional berdasarkan SOP atau proses kerja dalam dokumen ini, lalu sisipkan langsung ke naskah.',
  },
  ```
- Under `PRESETS_BY_HOST.PowerPoint`:
  ```typescript
  {
    id: 'ppt_insert_flowchart',
    label: '📐 Sisipkan Slide Flowchart Alur',
    prompt: 'Buatkan 1 slide visual khusus diagram alur proses (flowchart) bisnis dengan tema profesional dan naskah penjelasan pemateri.',
  },
  ```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/components/__tests__/chatAndActions.test.tsx`
Expected: PASS.

- [ ] **Step 5: Run full test suite and build verification**

Run:
```bash
npm run build
npx vitest run
```
Expected: All build checks pass and all 23+ test files (310+ tests) pass 100% green.

- [ ] **Step 6: Commit**

```bash
git add src/components/Chat/QuickActionPresets.tsx src/components/__tests__/chatAndActions.test.tsx
git commit -m "feat(ui): add flowchart quick action presets for Word and PowerPoint"
```
