import { describe, it, expect, beforeEach } from 'vitest';
import {
  addDocument,
  removeDocument,
  clearAllDocuments,
  listDocuments,
  queryKnowledge,
  getRagEngineState,
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
    expect(financialResults[0].sourceCitation).toBe('[kebijakan_keuangan.md (Bagian 1)]');

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

  it('filters knowledge search by documentId when provided', () => {
    const doc1 = addDocument('doc1.txt', 'Pembayaran tagihan listrik kantor pusat.');
    const doc2 = addDocument('doc2.txt', 'Pembayaran tagihan air cabang surabaya.');

    const resultsDoc1 = queryKnowledge('pembayaran', { documentId: doc1.id });
    expect(resultsDoc1.length).toBe(1);
    expect(resultsDoc1[0].chunk.documentId).toBe(doc1.id);

    const resultsDoc2 = queryKnowledge('pembayaran', { documentId: doc2.id });
    expect(resultsDoc2.length).toBe(1);
    expect(resultsDoc2[0].chunk.documentId).toBe(doc2.id);
  });

  it('returns engine state correctly and handles empty queries', () => {
    expect(queryKnowledge('')).toEqual([]);
    expect(queryKnowledge('   ')).toEqual([]);

    const doc = addDocument('manual.txt', 'Petunjuk operasional mesin fotokopi.');
    const state = getRagEngineState();
    expect(state.documents.length).toBe(1);
    expect(state.allChunks.length).toBe(doc.totalChunks);
    expect(state.lastIndexedAt).toBeGreaterThan(0);
  });
});
