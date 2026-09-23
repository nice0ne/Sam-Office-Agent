# Visual Flowchart & Process Diagram Inserter Design Spec

## 1. Overview & Business Value

In corporate daily office work, Standard Operating Procedures (SOPs), business process workflows, approval hierarchies, and system lifecycle diagrams are ubiquitous. Documenting and presenting them effectively in Microsoft Word and PowerPoint usually requires clunky external tools (Visio, draw.io, Lucidchart) or cumbersome manual drawing of shapes, connectors, and text alignment.

### Key Pain Points
1. **Manual Drawing Burden**: Inserting shapes, aligning boxes, and drawing connectors manually in Office is tedious, time-consuming, and frequently breaks when text wraps or layouts shift.
2. **Context Switching Friction**: Users leave Office to draw diagrams in external web tools, export them to local disk, and manually re-import them as pictures.
3. **Heavy External Dependencies**: Adding large diagramming libraries like `mermaid.js` (~4MB) directly to an Office Add-in risks memory bloat, sluggish load times, and rendering incompatibilities within embedded Office WebView2 runtimes.
4. **Data Privacy**: Sending confidential internal processes to cloud-based diagram rendering APIs violates enterprise privacy standards.

### The Solution: Zero-Dependency Visual Flowchart & Process Diagram Engine
Equip **Sam Office Agent** with an autonomous, 100% client-side, zero-dependency 2D Canvas rendering engine and natural language process extractor:
- **Zero-Dependency Native Rendering (`src/utils/diagramRenderer.ts`)**: Generates crisp, Hi-DPI Base64 PNG diagrams directly using browser HTML5 Canvas 2D API without heavy external libraries.
- **Autonomous Process Extractor (`synthesizeFlowchartFromText`)**: Intelligently extracts steps, decisions, and connections from user prompts or active document text.
- **Direct Office Insertion Pipeline**:
  - In Word: Uses `WordDriver.insertInlinePictureFromBase64()` to cleanly insert process flowcharts with captions.
  - In PowerPoint: Uses `PPTDriver` to insert dedicated diagram slides with clean layouts or add picture boxes to active slides.
- **Corporate Styling**: Modern rounded nodes, pill start/end nodes, decision diamonds, precise orthogonal/curved connector arrows, and professional color themes (`corporate_navy`, `emerald_executive`, `modern_dark`, `amber_warm`).
- **100% Confidential & Client-Side**: No data leaves the user's machine; all layout calculation and pixel generation occurs in-process.

---

## 2. Architecture & Data Contracts

### 2.1 Types (`src/utils/diagramRenderer.ts` / `src/services/office/types.ts`)

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
  label?: string; // e.g., "Ya" / "Tidak", "Lolos" / "Gagal"
  style?: 'solid' | 'dashed';
}

export interface FlowchartDefinition {
  title?: string;
  direction?: 'TD' | 'LR'; // Top-Down (default) or Left-Right
  theme?: 'corporate_navy' | 'emerald_executive' | 'modern_dark' | 'amber_warm';
  nodes: FlowchartNode[];
  edges: FlowchartEdge[];
}

export interface FlowchartOptions {
  title?: string;
  theme?: 'corporate_navy' | 'emerald_executive' | 'modern_dark' | 'amber_warm';
  direction?: 'TD' | 'LR';
  scale?: number; // e.g. 2 for Hi-DPI
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

### 2.2 Driver Interface Extensions (`src/services/office/types.ts`)

```typescript
export interface IOfficeDriver {
  // Existing methods...
  insertProcessFlowchart?(options: InsertFlowchartOptions): Promise<DiagramResult>;
}

export interface IDocumentDriver extends IOfficeDriver {
  insertInlinePictureFromBase64?(base64: string, location?: string): Promise<void>;
  insertProcessFlowchart?(options: InsertFlowchartOptions): Promise<DiagramResult>;
}
```

---

## 3. Component Design & Implementation

### 3.1 Flowchart Layout & Rendering Engine (`src/utils/diagramRenderer.ts`)

- **Text Synthesizer (`synthesizeFlowchartFromText`)**:
  - Parses input text (numbered steps, SOP bullet points, or raw narrative).
  - Categorizes nodes:
    - Keywords like `Mulai`, `Start`, `Awal` -> `type: 'start'`
    - Keywords like `Apakah`, `Cek`, `Validasi`, `Review`, `Persetujuan` -> `type: 'decision'`
    - Keywords like `Dokumen`, `Laporan`, `Formulir`, `Berkas` -> `type: 'document'`
    - Keywords like `Selesai`, `Tuntas`, `Kirim`, `End` -> `type: 'end'`
    - Standard action verbs -> `type: 'process'`
  - Connects sequential nodes with directed edges. Decision nodes branch into "Ya" (next step) and "Tidak" (revision or termination).
  
- **Layout Calculator**:
  - Assigns hierarchical levels (`level = 0, 1, 2...`) using dependency depth.
  - Determines node dimensions (auto-padding based on label text length).
  - Aligns nodes symmetrically along the primary axis (`TD` or `LR`).
  - Computes orthogonal connector paths between nodes with arrowheads.

- **Canvas 2D Renderer (`renderFlowchartToPngBase64`)**:
  - Renders to an offscreen HTML5 `<canvas>` (scaled at 2x Hi-DPI for crisp text and edges).
  - Applies selected corporate color palette:
    - `corporate_navy`: Primary `#1E3A8A`, Accent `#3B82F6`, Background `#F8FAFC`, Border `#94A3B8`.
    - `emerald_executive`: Primary `#065F46`, Accent `#10B981`, Background `#F0FDF4`, Border `#A7F3D0`.
    - `modern_dark`: Primary `#0F172A`, Accent `#64748B`, Background `#1E293B`, Border `#475569`.
    - `amber_warm`: Primary `#92400E`, Accent `#F59E0B`, Background `#FFFBEB`, Border `#FDE68A`.
  - Node drawing:
    - `start` / `end`: Stadium / Pill shape.
    - `process`: Rounded rectangle (radius 8px).
    - `decision`: Diamond polygon with branch labels.
    - `document`: Curved bottom wave box.
  - Returns raw Base64 PNG string without data URI scheme header.

### 3.2 Office Drivers Implementation

- **`mockDriver.ts`**:
  - Generates `DiagramResult` using `synthesizeFlowchartFromText` and `renderFlowchartToPngBase64`.
  - Records the inserted diagram state in a `diagrams: DiagramResult[]` array for unit testing.
- **`wordDriver.ts`**:
  - In `Word.run`, calls `context.document.body.insertInlinePictureFromBase64(base64Png, Word.InsertLocation.end)`.
  - Appends caption paragraph below the picture: `Gambar: [Judul Flowchart]`.
- **`pptDriver.ts`**:
  - In `PowerPoint.run`, adds a dedicated slide with clean title and inserts the flowchart picture centered on the slide canvas.

### 3.3 Specialist Agent ReAct Tools

Registered in `WordAgent` and `PPTAgent`:
- **Tool Schema**:
  - `name`: `insert_process_flowchart`
  - `description`: `"Membuat diagram alur proses bisnis/SOP (flowchart visual profesional) dari teks atau deskripsi alur kerja, lalu langsung menyisipkannya sebagai gambar tajam ke dokumen Word atau slide PowerPoint."`
  - `parameters`:
    - `textOrSteps` (string, required)
    - `title` (string, optional)
    - `theme` (enum, optional)
    - `direction` (enum `TD` | `LR`, optional)
- **Execution**:
  - Obtains active Office driver.
  - Calls `driver.insertProcessFlowchart({ textOrSteps, title, theme, direction })`.
  - Returns a formatted summary of generated nodes and edges with visual confirmation.

### 3.4 UI Quick Action Presets (`QuickActionPresets.tsx`)

- **Word Host**:
  - Label: `📐 Buat Diagram Alur (Flowchart)`
  - Prompt: `Buat diagram alur proses visual (flowchart) profesional berdasarkan SOP atau proses kerja dalam dokumen ini, lalu sisipkan langsung ke naskah.`
- **PowerPoint Host**:
  - Label: `📐 Sisipkan Slide Flowchart Alur`
  - Prompt: `Buatkan 1 slide visual khusus diagram alur proses (flowchart) bisnis dengan tema profesional dan naskah penjelasan pemateri.`

---

## 4. Verification & Testing Strategy

1. **Unit Tests for Diagram Renderer (`tests/diagramRenderer.test.ts`)**:
   - Verify `synthesizeFlowchartFromText` correctly extracts start, process, decision, and end nodes.
   - Verify node classification and edge connections for multi-step SOP text.
   - Verify `renderFlowchartToPngBase64` outputs valid Base64 PNG image data.
2. **Office Drivers Tests (`src/services/office/__tests__/mockDriver.test.ts`)**:
   - Verify `insertProcessFlowchart` executes and records diagram result in `MockOfficeDriver`.
3. **Specialist Agents Tests (`src/agents/__tests__/specialistAgents.test.ts`)**:
   - Verify `insert_process_flowchart` is declared in `WordAgent` and `PPTAgent`.
   - Verify tool execution formats rich response with node/edge summary and confirmation.
4. **UI Presets Tests (`src/components/__tests__/chatAndActions.test.tsx`)**:
   - Verify flowchart quick action preset buttons render in Word and PowerPoint hosts.
5. **Full Regression Test Suite**:
   - Verify all existing 302+ tests continue to pass with 0 regressions.
