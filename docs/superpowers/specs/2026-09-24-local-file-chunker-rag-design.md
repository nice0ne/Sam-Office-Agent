# Local Reference File Chunker (Mini-RAG) Design Spec

## 1. Overview & Business Value

Knowledge workers drafting contracts in Word, consolidating financial statements in Excel, or preparing executive decks in PowerPoint routinely consult external reference materials: internal SOPs, vendor guidelines, prior quarterly reports, product specifications, or raw CSV/JSON dumps.

### Key Pain Points
1. **Context Window Token Exhaustion**: Pasting multi-page manuals or large datasets directly into chat prompts quickly overflows LLM context windows and degrades model attention.
2. **Hallucination on Specific Facts**: Without grounded retrieval from local documents, LLMs guess or hallucinate specific clauses, product codes, or compliance requirements.
3. **Data Confidentiality & Privacy**: Cloud-based vector databases or remote embedding APIs require transmitting proprietary corporate documents outside the local workstation, violating enterprise security policies.
4. **Setup Complexity**: Traditional RAG systems require setting up vector databases, Python backends, or embedding model subscriptions.

### The Solution: Zero-Dependency In-Memory Mini-RAG Engine
Equip **Sam Office Agent** with an autonomous, 100% client-side, zero-cloud Mini-RAG system:
- **Zero-Cloud Local Ingestion (`src/services/rag/fileChunker.ts`)**: Ingests `.txt`, `.md`, `.csv`, `.json`, `.html`, and `.log` files directly in-browser using standard HTML5 `FileReader`.
- **Smart Semantic Chunker**: Splits text along paragraph/heading boundaries with configurable target sizes (~1,200 chars / ~250 words) and sliding window overlap (~150 chars) to prevent context fragmentation.
- **In-Memory BM25 Lexical Retrieval (`src/services/rag/ragEngine.ts`)**: Uses the industry-standard BM25 ranking function ($k_1=1.5, b=0.75$) with Inverse Document Frequency (IDF) computed purely in JavaScript memory without needing external embedding APIs.
- **Autonomous ReAct Tool (`search_reference_knowledge`)**: Registered across `ExcelAgent`, `WordAgent`, `PPTAgent`, and orchestrated by `SamCoordinator`.
- **Chat UI File Attachment**: A paperclip attachment button in `InputBar.tsx` allowing one-click file indexing, active document chip badges, and quick-action presets.

---

## 2. Architecture & Data Contracts

### 2.1 Types (`src/services/rag/types.ts`)

```typescript
export interface DocumentChunk {
  id: string;              // e.g. "doc-167890_chunk-0"
  documentId: string;      // Parent document ID
  fileName: string;        // e.g. "sop_pengadaan_2026.md"
  fileType: string;        // "txt" | "md" | "csv" | "json" | "html" | "log"
  chunkIndex: number;      // 0-indexed sequence
  text: string;            // Text content of the chunk
  charCount: number;       // Character length
  wordCount: number;       // Word count estimate
  keywords: string[];      // Extracted searchable tokens
}

export interface ChunkingOptions {
  targetChunkSize?: number; // Target characters per chunk (default: 1200)
  chunkOverlap?: number;    // Sliding window overlap (default: 150)
  preserveParagraphs?: boolean; // Keep paragraph boundaries (default: true)
}

export interface ProcessedDocument {
  id: string;
  name: string;
  size: number;
  type: string;
  uploadedAt: number;
  totalChunks: number;
  chunks: DocumentChunk[];
}

export interface SearchReferenceResult {
  chunk: DocumentChunk;
  score: number;
  snippet: string;
  sourceCitation: string;   // e.g. "[sop_pengadaan_2026.md (Bagian 1)]"
}

export interface RagEngineState {
  documents: ProcessedDocument[];
  allChunks: DocumentChunk[];
  lastIndexedAt: number;
}
```

---

## 3. Component Design & Implementation

### 3.1 Smart File Chunker (`src/services/rag/fileChunker.ts`)

- **`chunkText(text: string, fileName: string, options?: ChunkingOptions): DocumentChunk[]`**:
  1. Identifies file extension and normalizes line endings (`\r\n` -> `\n`).
  2. Splits into raw segments by paragraph boundaries (`\n\n`) or heading markers (`^#{1,6}\s`).
  3. Iteratively merges segments into chunks respecting `targetChunkSize` (1,200 chars).
  4. Appends a sliding window prefix of ~150 characters from the tail of the preceding chunk to preserve boundary continuity.
  5. Cleans and tokenizes keywords (removing common Indonesian & English stop words).
  6. Returns array of `DocumentChunk`.

### 3.2 In-Memory BM25 Retrieval Engine (`src/services/rag/ragEngine.ts`)

- **BM25 Algorithm Implementation**:
  - Calculates term frequencies ($TF$) per chunk.
  - Computes average chunk length ($avgdl$).
  - Calculates $IDF(q_i) = \ln\left(\frac{N - n(q_i) + 0.5}{n(q_i) + 0.5} + 1\right)$.
  - Scores chunks against query terms:
    $$\text{Score}(D, Q) = \sum_{q_i \in Q} \text{IDF}(q_i) \cdot \frac{f(q_i, D) \cdot (k_1 + 1)}{f(q_i, D) + k_1 \cdot \left(1 - b + b \cdot \frac{|D|}{\text{avgdl}}\right)}$$
- **State Management**:
  - `addDocument(name: string, content: string, options?: ChunkingOptions): ProcessedDocument`
  - `removeDocument(documentId: string): void`
  - `clearAllDocuments(): void`
  - `listDocuments(): ProcessedDocument[]`
  - `queryKnowledge(query: string, options?: { maxResults?: number; documentId?: string }): SearchReferenceResult[]`
    - Returns top matching chunks sorted by BM25 score descending (filtered to score > 0).

### 3.3 Specialist Agent ReAct Tool (`search_reference_knowledge`)

- Registered in `ExcelAgent`, `WordAgent`, and `PPTAgent`.
- **Tool Schema**:
  ```typescript
  {
    name: 'search_reference_knowledge',
    description: 'Mencari fakta, klausul, data angka, atau pedoman spesifik dari berkas referensi lokal yang dilampirkan pengguna (SOP, pedoman, data CSV, catatan laporan).',
    parameters: {
      type: 'object',
      properties: {
        query: {
          type: 'string',
          description: 'Pertanyaan atau kata kunci topik yang ingin dicari di dalam berkas referensi.',
        },
        maxResults: {
          type: 'number',
          description: 'Jumlah kutipan chunk relevan yang ingin diambil (default: 3, maksimal: 8).',
        },
      },
      required: ['query'],
    },
  }
  ```
- **Execution**:
  - Invokes `queryKnowledge(query, { maxResults })`.
  - Returns structured markdown containing relevant snippets, chunk IDs, and source file citations:
    ```markdown
    📖 **Hasil Pencarian Berkas Referensi:**
    1. **Kutipan dari [sop_pengadaan_2026.md (Bagian 1)]**:
       "..."
    ```
- **Coordinator Integration (`samCoordinator.ts`)**:
  - Adds Rule 9 instructing agents to query reference files before making factual assertions when documents are attached.

### 3.4 Chat UI Integration (`src/components/Chat/InputBar.tsx`)

- **Paperclip Attachment Button (📎)**:
  - Hidden file input accepting `.txt,.md,.csv,.json,.html,.log`.
  - On file selection: reads file using `FileReader.readAsText()`, calls `addDocument(file.name, content)`.
- **Active Document Chips**:
  - Renders above textarea showing file name, chunk count, and remove button `[✕]`.
- **Quick Action Preset**:
  - `id: 'universal_query_rag'`
  - Label: `📑 Analisis Berkas Referensi`
  - Prompt: `Analisis informasi dan poin-poin penting dari berkas referensi yang saya lampirkan, lalu simpulkan intisarinya ke dalam dokumen ini.`

---

## 4. Verification & Testing Strategy

1. **Unit Tests for File Chunker (`tests/fileChunker.test.ts`)**:
   - Verify splitting on paragraph boundaries and character thresholds.
   - Verify sliding window overlap preservation.
   - Verify handling of structured CSV, JSON, and short text.
2. **Unit Tests for BM25 Engine (`tests/ragEngine.test.ts`)**:
   - Verify BM25 ranking places relevant chunks first.
   - Verify source citation formatting.
   - Verify document addition, removal, and clearing.
3. **Integration Tests for Agent Tool (`tests/ragAgentTools.test.ts`)**:
   - Verify `search_reference_knowledge` is registered and executes across `ExcelAgent`, `WordAgent`, and `PPTAgent`.
4. **UI Tests (`src/components/__tests__/chatAndActions.test.tsx`)**:
   - Verify paperclip attachment button and active file badges render in `InputBar`.
5. **Full Regression Test Suite**:
   - Verify that all existing 328+ tests continue to pass with 0 regressions.
