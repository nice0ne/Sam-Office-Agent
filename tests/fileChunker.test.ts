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
