export interface DocumentChunk {
  id: string;
  documentId: string;
  fileName: string;
  fileType: string;
  chunkIndex: number;
  text: string;
  charCount: number;
  wordCount: number;
  keywords: string[];
}

export interface ChunkingOptions {
  targetChunkSize?: number;
  chunkOverlap?: number;
  preserveParagraphs?: boolean;
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
  sourceCitation: string;
}

export interface RagEngineState {
  documents: ProcessedDocument[];
  allChunks: DocumentChunk[];
  lastIndexedAt: number;
}
