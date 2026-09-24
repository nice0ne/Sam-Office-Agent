import {
  ChunkingOptions,
  DocumentChunk,
  ProcessedDocument,
  RagEngineState,
  SearchReferenceResult,
} from './types';
import { chunkText, extractFileType } from './fileChunker';

// In-memory document storage
let documents: ProcessedDocument[] = [];
let lastIndexedAt = 0;

/**
 * Tokenize a text string into lowercase words/terms.
 */
function tokenize(text: string): string[] {
  if (!text) return [];
  return text.toLowerCase().match(/[a-z0-9_\u00C0-\u017F]+/g) || [];
}

/**
 * Generates a unique document ID.
 */
function generateDocumentId(): string {
  return `doc_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
}

/**
 * Ingests and chunks a reference document into in-memory state.
 */
export function addDocument(
  name: string,
  content: string,
  options?: ChunkingOptions
): ProcessedDocument {
  const docId = generateDocumentId();
  const chunks = chunkText(content, name, options, docId);

  const docSize =
    typeof TextEncoder !== 'undefined'
      ? new TextEncoder().encode(content).length
      : content.length;

  const doc: ProcessedDocument = {
    id: docId,
    name,
    size: docSize,
    type: chunks.length > 0 ? chunks[0].fileType : extractFileType(name),
    uploadedAt: Date.now(),
    totalChunks: chunks.length,
    chunks,
  };

  documents.push(doc);
  lastIndexedAt = Date.now();
  return doc;
}

/**
 * Removes a document and all its chunks by document ID.
 */
export function removeDocument(documentId: string): void {
  documents = documents.filter((doc) => doc.id !== documentId);
  lastIndexedAt = Date.now();
}

/**
 * Clears all ingested documents and chunks.
 */
export function clearAllDocuments(): void {
  documents = [];
  lastIndexedAt = Date.now();
}

/**
 * Lists all active processed documents in memory.
 */
export function listDocuments(): ProcessedDocument[] {
  return [...documents];
}

/**
 * Returns current snapshot of the RAG engine state.
 */
export function getRagEngineState(): RagEngineState {
  return {
    documents: [...documents],
    allChunks: documents.flatMap((doc) => doc.chunks),
    lastIndexedAt,
  };
}

/**
 * Queries the knowledge base using the BM25 retrieval algorithm.
 *
 * BM25 Parameters:
 * - k1 = 1.5
 * - b = 0.75
 * - Standard Robertson-Spärck Jones IDF with Lucene smoothing (+1).
 */
export function queryKnowledge(
  query: string,
  options?: { maxResults?: number; documentId?: string }
): SearchReferenceResult[] {
  if (!query || !query.trim() || documents.length === 0) {
    return [];
  }

  // Determine target chunks
  let targetChunks: DocumentChunk[] = [];
  if (options?.documentId) {
    const targetDoc = documents.find((doc) => doc.id === options.documentId);
    if (!targetDoc) return [];
    targetChunks = targetDoc.chunks;
  } else {
    targetChunks = documents.flatMap((doc) => doc.chunks);
  }

  if (targetChunks.length === 0) {
    return [];
  }

  const queryTerms = tokenize(query);
  if (queryTerms.length === 0) {
    return [];
  }

  const uniqueQueryTerms = Array.from(new Set(queryTerms));

  // Pre-tokenize target chunks and calculate average document length
  const chunkTokenMap = new Map<string, string[]>();
  let totalLength = 0;

  for (const chunk of targetChunks) {
    const tokens = tokenize(chunk.text);
    chunkTokenMap.set(chunk.id, tokens);
    totalLength += tokens.length;
  }

  const N = targetChunks.length;
  const avgdl = totalLength / N || 1;

  // Compute Document Frequency n(q) for each unique query term
  const dfMap = new Map<string, number>();
  for (const term of uniqueQueryTerms) {
    let count = 0;
    for (const chunk of targetChunks) {
      const tokens = chunkTokenMap.get(chunk.id)!;
      if (tokens.includes(term)) {
        count++;
      }
    }
    dfMap.set(term, count);
  }

  const k1 = 1.5;
  const b = 0.75;
  const results: SearchReferenceResult[] = [];

  // Compute BM25 score per chunk
  for (const chunk of targetChunks) {
    const tokens = chunkTokenMap.get(chunk.id)!;
    const docLen = tokens.length;

    // Calculate term frequencies in current chunk
    const tfMap = new Map<string, number>();
    for (const token of tokens) {
      tfMap.set(token, (tfMap.get(token) || 0) + 1);
    }

    let score = 0;
    for (const term of uniqueQueryTerms) {
      const tf = tfMap.get(term) || 0;
      if (tf > 0) {
        const nq = dfMap.get(term) || 0;
        // Standard BM25 IDF formula with smoothing: ln(((N - nq + 0.5) / (nq + 0.5)) + 1)
        const idf = Math.log(((N - nq + 0.5) / (nq + 0.5)) + 1);
        const numerator = tf * (k1 + 1);
        const denominator = tf + k1 * (1 - b + b * (docLen / avgdl));
        score += idf * (numerator / denominator);
      }
    }

    if (score > 0) {
      const snippet =
        chunk.text.length > 250
          ? `${chunk.text.slice(0, 250).trim()}...`
          : chunk.text;
      const sourceCitation = `[${chunk.fileName} (Bagian ${chunk.chunkIndex + 1})]`;

      results.push({
        chunk,
        score,
        snippet,
        sourceCitation,
      });
    }
  }

  // Sort descending by relevance score
  results.sort((a, b) => b.score - a.score);

  const maxResults = options?.maxResults ?? 3;
  return results.slice(0, maxResults);
}
