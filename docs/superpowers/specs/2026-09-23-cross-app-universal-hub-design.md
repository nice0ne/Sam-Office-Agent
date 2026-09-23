# Cross-App Universal Hub & Data Bridge Design Spec

## 1. Overview & Business Value

In standard corporate office workflows, work is rarely confined to a single application:
1. **Excel to Word**: Financial analysts and project managers aggregate numerical tables and metrics in Excel, and then manually copy-paste them into Word documents to write formal executive reports, proposals, or memos.
2. **Excel to PowerPoint**: Commercial teams analyze quarterly sales or KPIs in Excel, and then re-type or screenshot data into PowerPoint slides for board and client presentations.
3. **Word to PowerPoint**: Planners draft comprehensive Product Requirement Documents (PRDs) or Standard Operating Procedures (SOPs) in Word, and then manually condense them into slide decks.

### The Problem
- **Context Fragmentation**: Word, Excel, and PowerPoint operate as separate desktop processes. In native WebView2 Office Add-ins, there is no shared in-memory object between host windows.
- **Copy-Paste Overhead & Formatter Friction**: Copying tables between Excel and Word often causes broken cell formatting, inconsistent column widths, and missing headers.
- **Manual Translation**: Raw tabular data lacks automated narrative extraction when pasted into Word, requiring repetitive prose writing.

### The Solution: Sam Universal Hub
The **Sam Universal Hub** is a 100% client-side, zero-backend persistent cross-application bridge. It allows any Office host (Excel, Word, PowerPoint) to publish and consume rich multi-modal artifacts (`table_data`, `executive_summary`, `chart_metrics`, `presentation_outline`) securely via origin-isolated browser storage (`localStorage`), complete with a proactive Smart Context Card UI that prompts one-click actions when switching applications.

---

## 2. Architecture & Data Contracts

### 2.1 Types & Artifacts (`src/types/index.ts` & `src/services/storage/crossAppBridge.ts`)

```typescript
export type CrossAppArtifactType =
  | 'table_data'
  | 'executive_summary'
  | 'chart_metrics'
  | 'presentation_outline';

export interface CrossAppTableData {
  headers: string[];
  rows: (string | number)[][];
  totalRows: number;
}

export interface CrossAppMetric {
  label: string;
  value: string;
  trend?: string;
}

export interface CrossAppSnapshot {
  id: string;
  sourceHost: HostType;
  title: string;
  artifactType: CrossAppArtifactType;
  tableData?: CrossAppTableData;
  summaryText?: string;
  metrics?: CrossAppMetric[];
  timestamp: number;
  metadata?: Record<string, any>;
}
```

### 2.2 Bridge Engine API (`src/services/storage/crossAppBridge.ts`)

```typescript
export function saveCrossAppSnapshot(
  snapshot: Omit<CrossAppSnapshot, 'id' | 'timestamp'>
): CrossAppSnapshot;

export function getLatestCrossAppSnapshot(
  excludeHost?: HostType,
  maxAgeMs?: number
): CrossAppSnapshot | null;

export function listCrossAppSnapshots(maxAgeMs?: number): CrossAppSnapshot[];

export function dismissCrossAppSnapshot(id: string): void;

export function isCrossAppSnapshotDismissed(id: string): boolean;

export function clearCrossAppSnapshots(): void;
```

---

## 3. Specialist Agent Capabilities

### 3.1 Excel Specialist (`ExcelAgent.ts`)
- **Tool: `share_to_cross_app_hub`**
  - Extracts active table range or read sheet data.
  - Automatically identifies header row and data rows.
  - Computes or attaches summary text/metrics.
  - Invokes `saveCrossAppSnapshot` with `sourceHost: 'Excel'` and `artifactType: 'table_data'`.
  - Returns friendly confirmation in chat.

### 3.2 Word Specialist (`WordAgent.ts`)
- **Tool: `import_from_cross_app_hub`**
  - Fetches latest snapshot via `getLatestCrossAppSnapshot('Word')` or by specified ID.
  - If `tableData` is present: calls `driver.insertTable(rows, cols, data)`.
  - If `summaryText` is present: inserts heading and executive summary paragraph via `driver.insertContent`.
  - Returns structured report in chat.
- **Tool: `share_to_cross_app_hub`**
  - Shares active document outline or PRD sections to the Hub for PowerPoint consumption.

### 3.3 PowerPoint Specialist (`PPTAgent.ts`)
- **Tool: `import_from_cross_app_hub`**
  - Fetches latest snapshot via `getLatestCrossAppSnapshot('PowerPoint')` or by specified ID.
  - Formats snapshot content into structured text.
  - Forwards to `driver.transformDocToDeck()` to generate the 5-slide Executive Storyline Arc with Presenter Scripts.

---

## 4. UI Design: Smart Context Card & Presets

### 4.1 Proactive UI Component (`src/components/Chat/CrossAppSnapshotCard.tsx`)
- Displayed prominently directly above the Chat InputBar when:
  - `getLatestCrossAppSnapshot(host)` returns a non-null, un-dismissed snapshot.
- **Visual Presentation**:
  - Color-coded host badge: 🟢 Excel (green), 🔵 Word (blue), 🟠 PowerPoint (orange).
  - Title and details: e.g. `📌 Ditemukan data dari Excel: "Rekap Penjualan Q3" (Tabel 15 baris, 4 kolom • 2 mnt lalu)`.
  - Context-aware action buttons:
    - In Word: `[📄 Buat Laporan & Sisipkan Tabel]`
    - In PowerPoint: `[📊 Buat 5 Slide Presentasi]`
    - In Excel: `[📋 Tulis ke Lembar Kerja]`
  - Dismiss button `[✕]`: Hides the card and records dismissal in localStorage.

### 4.2 Quick Action Presets (`src/components/Chat/QuickActionPresets.tsx`)
- In Excel: `📤 Bagikan ke Hub (Word/PPT)`
- In Word: `📥 Impor Data dari Excel Hub`
- In PowerPoint: `📥 Buat Slide dari Hub (Excel/Word)`

---

## 5. Verification & Testing Strategy

1. **Storage Unit Tests (`tests/crossAppBridge.test.ts`)**:
   - Verify storing, retrieving by host, excluding active host, dismissal, and 24-hour expiration filter.
2. **Agent Integration Tests (`src/agents/__tests__/specialistAgents.test.ts`)**:
   - Verify `share_to_cross_app_hub` in `ExcelAgent` and `WordAgent`.
   - Verify `import_from_cross_app_hub` in `WordAgent` and `PPTAgent`.
3. **UI Integration Tests (`src/components/__tests__/chatAndActions.test.tsx`)**:
   - Verify `CrossAppSnapshotCard` rendering, dismissal, and action button triggering.
4. **Full Test Suite Run**:
   - Verify that all tests continue to pass with 0 regressions.
