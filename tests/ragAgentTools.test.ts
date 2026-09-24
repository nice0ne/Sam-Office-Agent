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
