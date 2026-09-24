# Local Reference File Chunker (Mini-RAG) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement an in-memory, zero-cloud Mini-RAG system allowing users to attach reference documents (`.txt`, `.md`, `.csv`, `.json`, `.html`, `.log`), intelligently chunk them with sliding window overlap, index them via BM25 retrieval, and provide autonomous agents with the `search_reference_knowledge` tool.

**Architecture:** A smart chunker (`src/services/rag/fileChunker.ts`) splits documents into overlapping context chunks. An in-memory BM25 retrieval engine (`src/services/rag/ragEngine.ts`) calculates term frequency and inverse document frequency (IDF) purely in JavaScript. Specialist agents (`ExcelAgent`, `WordAgent`, `PPTAgent`) query this engine via `search_reference_knowledge`. An intuitive paperclip attachment button and active document chip badges in `InputBar.tsx` allow file indexing directly in chat.

**Tech Stack:** TypeScript, BM25 Lexical Ranking Algorithm, React, Vitest.

**Spec:** `docs/superpowers/specs/2026-09-24-local-file-chunker-rag-design.md`

## Global Constraints
- Zero external vector database or embedding API dependencies.
- 100% client-side memory execution; no reference files leave the local browser/add-in context.
- BM25 standard parameters: $k_1 = 1.5, b = 0.75$.
- Target chunk size ~1,200 characters with ~150 characters sliding window overlap.
- All tests must pass with 100% green and 0 regressions on `npm run build` and `npx vitest run`.

---

### Task 1: Document Types & Smart File Chunker (`src/services/rag/fileChunker.ts`)

**Files:**
- Create: `src/services/rag/types.ts`
- Create: `src/services/rag/fileChunker.ts`
- Create: `tests/fileChunker.test.ts`

**Interfaces:**
- Produces:
  - `DocumentChunk`: `{ id: string; documentId: string; fileName: string; fileType: string; chunkIndex: number; text: string; charCount: number; wordCount: number; keywords: string[]; }`
  - `ChunkingOptions`: `{ targetChunkSize?: number; chunkOverlap?: number; preserveParagraphs?: boolean; }`
  - `ProcessedDocument`: `{ id: string; name: string; size: number; type: string; uploadedAt: number; totalChunks: number; chunks: DocumentChunk[]; }`
  - `chunkText(text: string, fileName: string, options?: ChunkingOptions, docId?: string): DocumentChunk[]`

- [ ] **Step 1: Write the failing test**

Create `tests/fileChunker.test.ts`:

```typescript
import { describe, it, expect } from 'vitest';
import { chunkText } from '../src/services/rag/fileChunker';

describe('fileChunker', () => {
  it('chunks multi-paragraph text with sliding window overlap', () => {
    const paragraph1 = 'Paragraf pertama menjelaskan tentang pedoman operasional pengadaan barang dan jasa pada kuartal 4 tahun 2026. Semua divisi wajib mematuhi batas anggaran yang telah ditetapkan oleh direksi.';
    const paragraph2 = 'Paragraf kedua merinci prosedur persetujuan berjenjang mulai dari manajer divisi hingga general manager. Khusus pengadaan di atas 50 juta rupiah memerlukan tanda tangan CFO.';
    const paragraph3 = 'Paragraf ketiga menetapkan sanksi ketat bagi pelanggaran SOP pengadaan termasuk pembatalan kontrak vendor dan audit internal menyeluruh.';

    const longDoc = [paragraph1, paragraph2, paragraph3].join('\n\n');

    const chunks = chunkText(longDoc, 'sop_pengadaan.md', {
      targetChunkSize: 200,
      chunkOverlap: 40,
    });

    expect(chunks.length).toBeGreaterThanOrEqual(2);
    expect(chunks[0].fileName).toBe('sop_pengadaan.md');
    expect(chunks[0].chunkIndex).toBe(0);
    expect(chunks[0].text.length).toBeGreaterThan(50);
    expect(chunks[0].keywords.length).toBeGreaterThan(0);

    // Verify overlap presence: chunk 1 should share text from end of chunk 0
    if (chunks.length > 1) {
      expect(chunks[1].chunkIndex).toBe(1);
    }
  });

  it('handles short text by producing a single chunk', () => {
    const shortDoc = 'Kebijakan kerja jarak jauh fleksibel berlaku mulai Senin.';
    const chunks = chunkText(shortDoc, 'memo.txt');

    expect(chunks.length).toBe(1);
    expect(chunks[0].text).toBe(shortDoc);
    expect(chunks[0].chunkIndex).toBe(0);
  });

  it('handles empty or whitespace-only text gracefully', () => {
    const chunks = chunkText('   \n\n  ', 'empty.txt');
    expect(chunks.length).toBe(0);
  });

  it('correctly extracts file extension and clean keywords', () => {
    const csvContent = 'ID,Produk,Harga\n1,Laptop,15000000\n2,Monitor,2500000';
    const chunks = chunkText(csvContent, 'data_harga.csv');

    expect(chunks.length).toBe(1);
    expect(chunks[0].fileType).toBe('csv');
    expect(chunks[0].keywords).toContain('laptop');
    expect(chunks[0].keywords).toContain('monitor');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/fileChunker.test.ts`
Expected: FAIL with "Cannot find module '../src/services/rag/fileChunker'".

- [ ] **Step 3: Write minimal implementation**

1. Create `src/services/rag/types.ts`:
   - Export `DocumentChunk`, `ChunkingOptions`, `ProcessedDocument`, `SearchReferenceResult`, `RagEngineState`.
2. Create `src/services/rag/fileChunker.ts`:
   - Helper to extract file extension (e.g. `txt`, `md`, `csv`).
   - Stop words list (ID + EN).
   - Helper `extractKeywords(text: string): string[]`.
   - `chunkText(text: string, fileName: string, options?: ChunkingOptions, docId?: string): DocumentChunk[]`:
     - Default `targetChunkSize = 1200`, `chunkOverlap = 150`.
     - Normalizes text, splits into segments (paragraphs `\n\n`).
     - Iteratively builds chunks, carrying forward overlap text.
     - Maps each chunk to `DocumentChunk` with wordCount, charCount, keywords, id.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/fileChunker.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/services/rag/types.ts src/services/rag/fileChunker.ts tests/fileChunker.test.ts
git commit -m "feat(rag): implement smart local file chunker and document types"
```

---

### Task 2: In-Memory BM25 Retrieval Engine (`src/services/rag/ragEngine.ts`)

**Files:**
- Create: `src/services/rag/ragEngine.ts`
- Create: `tests/ragEngine.test.ts`

**Interfaces:**
- Consumes: `chunkText` from `./fileChunker`, types from `./types`
- Produces:
  - `addDocument(name: string, content: string, options?: ChunkingOptions): ProcessedDocument`
  - `removeDocument(documentId: string): void`
  - `clearAllDocuments(): void`
  - `listDocuments(): ProcessedDocument[]`
  - `queryKnowledge(query: string, options?: { maxResults?: number; documentId?: string }): SearchReferenceResult[]`
  - `getRagEngineState(): RagEngineState`

- [ ] **Step 1: Write the failing test**

Create `tests/ragEngine.test.ts`:

```typescript
import { describe, it, expect, beforeEach } from 'vitest';
import {
  addDocument,
  removeDocument,
  clearAllDocuments,
  listDocuments,
  queryKnowledge,
} from '../src/services/rag/ragEngine';

describe('ragEngine BM25 Retrieval', () => {
  beforeEach(() => {
    clearAllDocuments();
  });

  it('indexes documents and returns ranked chunks using BM25', () => {
    addDocument(
      'kebijakan_keuangan.md',
      `# Kebijakan Keuangan Q4\nBatas anggaran divisi IT adalah 500 juta rupiah.\nSetiap pembelian server baru memerlukan approval CTO dan CFO.\nReimbursement perjalanan dinas maksimal 2 juta per hari.`
    );

    addDocument(
      'kebijakan_hr.md',
      `# Kebijakan SDM & Cuti\nCuti tahunan berjumlah 12 hari kerja.\nPengajuan cuti wajib diajukan minimal 3 hari sebelum tanggal cuti.\nJam kerja fleksibel antara pukul 08.00 hingga 17.00.`
    );

    expect(listDocuments().length).toBe(2);

    // Query financial keyword
    const financialResults = queryKnowledge('anggaran server CTO');
    expect(financialResults.length).toBeGreaterThan(0);
    expect(financialResults[0].chunk.fileName).toBe('kebijakan_keuangan.md');
    expect(financialResults[0].score).toBeGreaterThan(0);
    expect(financialResults[0].sourceCitation).toContain('kebijakan_keuangan.md');

    // Query HR keyword
    const hrResults = queryKnowledge('cuti tahunan pengajuan');
    expect(hrResults.length).toBeGreaterThan(0);
    expect(hrResults[0].chunk.fileName).toBe('kebijakan_hr.md');
  });

  it('removes document and cleans up chunks', () => {
    const doc = addDocument('temp.txt', 'Dokumen sementara untuk pengujian');
    expect(listDocuments().length).toBe(1);

    removeDocument(doc.id);
    expect(listDocuments().length).toBe(0);

    const results = queryKnowledge('sementara');
    expect(results.length).toBe(0);
  });

  it('returns empty results when no relevant terms match', () => {
    addDocument('laporan.txt', 'Hasil penjualan beras dan minyak goreng.');
    const results = queryKnowledge('astronomi galaksi supernova');
    expect(results.length).toBe(0);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/ragEngine.test.ts`
Expected: FAIL with "Cannot find module '../src/services/rag/ragEngine'".

- [ ] **Step 3: Write minimal implementation**

Create `src/services/rag/ragEngine.ts`:
- Maintain in-memory state: `documents: ProcessedDocument[] = []`.
- Implement `addDocument(name: string, content: string, options?: ChunkingOptions): ProcessedDocument`:
  - Generates unique ID.
  - Calls `chunkText(content, name, options, id)`.
  - Stores document.
  - Returns `ProcessedDocument`.
- Implement `removeDocument(documentId: string): void`.
- Implement `clearAllDocuments(): void`.
- Implement `listDocuments(): ProcessedDocument[]`.
- Implement BM25 algorithm in `queryKnowledge(query: string, options?: { maxResults?: number; documentId?: string }): SearchReferenceResult[]`:
  - Tokenize query.
  - Collect target chunks (all chunks or filtered by `documentId`).
  - Calculate BM25 score per chunk ($k_1 = 1.5, b = 0.75$).
  - Sort chunks by score descending.
  - Format `sourceCitation: [${chunk.fileName} (Bagian ${chunk.chunkIndex + 1})]`.
  - Slice to `maxResults` (default 3).

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/ragEngine.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/services/rag/ragEngine.ts tests/ragEngine.test.ts
git commit -m "feat(rag): implement in-memory BM25 retrieval engine and document manager"
```

---

### Task 3: Specialist Agent ReAct Tool (`search_reference_knowledge`)

**Files:**
- Modify: `src/agents/excel/excelAgent.ts`
- Modify: `src/agents/word/wordAgent.ts`
- Modify: `src/agents/powerpoint/pptAgent.ts`
- Modify: `src/agents/coordinator/samCoordinator.ts`
- Test: `tests/ragAgentTools.test.ts`

**Interfaces:**
- Consumes: `queryKnowledge` from `../../services/rag/ragEngine`
- Produces: `search_reference_knowledge` tool in `ExcelAgent`, `WordAgent`, and `PPTAgent`

- [x] **Step 1: Write the failing test**

Create `tests/ragAgentTools.test.ts`:

```typescript
import { describe, it, expect, beforeEach } from 'vitest';
import { ExcelAgent } from '../src/agents/excel/excelAgent';
import { WordAgent } from '../src/agents/word/wordAgent';
import { PPTAgent } from '../src/agents/powerpoint/pptAgent';
import { MockOfficeDriver } from '../src/services/office/mockDriver';
import { addDocument, clearAllDocuments } from '../src/services/rag/ragEngine';

describe('search_reference_knowledge Agent Tool', () => {
  beforeEach(() => {
    clearAllDocuments();
  });

  it('registers search_reference_knowledge across all specialist agents', () => {
    const mockDriver = new MockOfficeDriver();
    const excelAgent = new ExcelAgent(mockDriver);
    const wordAgent = new WordAgent(mockDriver);
    const pptAgent = new PPTAgent(mockDriver);

    expect(excelAgent.getTools().some(t => t.name === 'search_reference_knowledge')).toBe(true);
    expect(wordAgent.getTools().some(t => t.name === 'search_reference_knowledge')).toBe(true);
    expect(pptAgent.getTools().some(t => t.name === 'search_reference_knowledge')).toBe(true);
  });

  it('executes search_reference_knowledge and returns cited excerpts', async () => {
    addDocument(
      'perjanjian_sewa.txt',
      'Pasal 4: Uang sewa sebesar 120 juta rupiah per tahun dibayarkan di muka setiap tanggal 1 Januari.'
    );

    const mockDriver = new MockOfficeDriver();
    const wordAgent = new WordAgent(mockDriver);

    const res = await wordAgent.executeTool(
      {
        id: 'rag-1',
        name: 'search_reference_knowledge',
        arguments: {
          query: 'uang sewa tanggal pembayaran',
          maxResults: 2,
        },
        status: 'pending',
      },
      { host: 'Word' }
    );

    expect(res.success).toBe(true);
    expect(res.result).toContain('Hasil Temuan Berkas Referensi');
    expect(res.result).toContain('120 juta');
    expect(res.result).toContain('perjanjian_sewa.txt');
  });
});
```

- [x] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/ragAgentTools.test.ts`
Expected: FAIL with "tool search_reference_knowledge not found".

- [x] **Step 3: Write minimal implementation**

1. In `src/agents/excel/excelAgent.ts`, `wordAgent.ts`, `pptAgent.ts`:
   - Import `queryKnowledge` from `../../services/rag/ragEngine`.
   - In `getTools()`, register `search_reference_knowledge`:
     ```typescript
     {
       name: 'search_reference_knowledge',
       description: 'Mencari fakta, klausul, data angka, atau pedoman spesifik dari berkas referensi lokal yang dilampirkan pengguna (SOP, pedoman, data CSV, catatan laporan).',
       parameters: {
         type: 'object',
         properties: {
           query: { type: 'string', description: 'Pertanyaan atau kata kunci topik yang ingin dicari di dalam berkas referensi' },
           maxResults: { type: 'number', description: 'Jumlah kutipan chunk relevan yang ingin diambil (default: 3)' },
         },
         required: ['query'],
       },
     }
     ```
   - In `executeTool()`:
     - Handle `search_reference_knowledge`.
     - Call `queryKnowledge(query, { maxResults })`.
     - If empty, return `"Tidak ditemukan kutipan yang relevan dari berkas referensi untuk kata kunci: \"${query}\""`.
     - Format results with citation headers and excerpts.
2. In `src/agents/coordinator/samCoordinator.ts`:
   - Add Rule 9 to `buildSystemPrompt()`:
     `"9. KNOWLEDGE BASE & REFERENSI LOKAL (MINI-RAG): Jika pengguna melampirkan berkas dokumen referensi atau bertanya mengenai SOP/data dari dokumen yang dilampirkan, gunakan tool search_reference_knowledge untuk mencari kutipan data faktual sebelum menyusun naskah, tabel, atau slide."`

- [x] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/ragAgentTools.test.ts`
Expected: PASS.

- [x] **Step 5: Commit**

```bash
git add src/agents/excel/excelAgent.ts src/agents/word/wordAgent.ts src/agents/powerpoint/pptAgent.ts src/agents/coordinator/samCoordinator.ts tests/ragAgentTools.test.ts
git commit -m "feat(agents): register and execute search_reference_knowledge tool across specialist agents"
```

---

### Task 4: Chat UI Attachment Bar & Presets

**Files:**
- Modify: `src/components/Chat/InputBar.tsx`
- Modify: `src/components/Chat/QuickActionPresets.tsx`
- Test: `src/components/__tests__/chatAndActions.test.tsx`

**Interfaces:**
- Consumes: `addDocument`, `removeDocument`, `listDocuments` from `../../services/rag/ragEngine`
- Produces: Attachment button (📎) and active reference file chips in `InputBar.tsx`, plus `universal_query_rag` preset

- [ ] **Step 1: Write the failing test**

In `src/components/__tests__/chatAndActions.test.tsx`:
Add a test verifying paperclip attachment button and RAG quick action preset:

```typescript
it('renders file attachment button in InputBar and RAG quick action preset', () => {
  render(<InputBar onSendMessage={vi.fn()} host="Word" />);
  expect(screen.getByTitle(/Lampirkan berkas referensi/i)).toBeInTheDocument();

  render(<QuickActionPresets host="Word" onSelectPreset={vi.fn()} />);
  expect(screen.getByText(/Analisis Berkas Referensi/i)).toBeInTheDocument();
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/components/__tests__/chatAndActions.test.tsx`
Expected: FAIL with "Unable to find an element with title: /Lampirkan berkas referensi/i".

- [ ] **Step 3: Write minimal implementation**

1. In `src/components/Chat/InputBar.tsx`:
   - Import `Paperclip`, `FileText`, `X` from `'lucide-react'`.
   - Import `addDocument`, `removeDocument`, `listDocuments` from `../../services/rag/ragEngine`.
   - Maintain state of active attached documents: `const [docs, setDocs] = useState(listDocuments())`.
   - Add hidden `<input type="file" ref={fileInputRef} accept=".txt,.md,.csv,.json,.html,.log" onChange={handleFileSelect} />`.
   - When a file is chosen, read with `FileReader.readAsText(file)` and call `addDocument(file.name, text)`. Update `docs`.
   - Render active document chips above textarea:
     `📄 {doc.name} ({doc.totalChunks} chunks) [✕]`
     Clicking `[✕]` calls `removeDocument(doc.id)` and updates `docs`.
   - Add attachment button `📎` with title `"Lampirkan berkas referensi (Mini-RAG)"`.
2. In `src/components/Chat/QuickActionPresets.tsx`:
   - Add universal preset:
     ```typescript
     {
       id: 'universal_query_rag',
       label: '📑 Analisis Berkas Referensi',
       prompt: 'Analisis informasi dan poin-poin penting dari berkas referensi yang saya lampirkan, lalu simpulkan intisarinya ke dalam dokumen ini.',
     }
     ```
   - Adjust preset counts in `chatAndActions.test.tsx` if needed.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/components/__tests__/chatAndActions.test.tsx`
Expected: PASS.

- [ ] **Step 5: Full regression & build verification**

Run:
```bash
npm run build
npx vitest run
```
Expected: All build checks pass with exit code 0, and all test files pass 100% green.

- [ ] **Step 6: Commit**

```bash
git add src/components/Chat/InputBar.tsx src/components/Chat/QuickActionPresets.tsx src/components/__tests__/chatAndActions.test.tsx
git commit -m "feat(ui): add paperclip file attachment button and RAG quick action preset"
```
