import { describe, it, expect, beforeEach } from 'vitest';
import { ExcelAgent } from '../src/agents/excel/excelAgent';
import { WordAgent } from '../src/agents/word/wordAgent';
import { PPTAgent } from '../src/agents/powerpoint/pptAgent';
import { MockOfficeDriver, resetOfficeDriver } from '../src/services/office';
import {
  saveCrossAppSnapshot,
  getLatestCrossAppSnapshot,
  clearCrossAppSnapshots,
} from '../src/services/storage/crossAppBridge';

describe('Cross-App Universal Hub Agent Tools', () => {
  beforeEach(() => {
    clearCrossAppSnapshots();
    if (typeof localStorage !== 'undefined') {
      localStorage.clear();
    }
    resetOfficeDriver();
  });

  it('share_to_cross_app_hub in ExcelAgent saves table data and summary to Universal Hub', async () => {
    const mockDriver = new MockOfficeDriver();
    mockDriver.mockData = [
      ['Kategori', 'Q1', 'Q2'],
      ['Elektronik', 500, 700],
      ['Pakaian', 300, 450],
    ];

    const agent = new ExcelAgent(mockDriver);
    const tools = agent.getTools();
    expect(tools.some(t => t.name === 'share_to_cross_app_hub')).toBe(true);

    const res = await agent.executeTool(
      {
        id: 'ex-hub-1',
        name: 'share_to_cross_app_hub',
        arguments: {
          title: 'Penjualan Semester 1',
          summaryText: 'Pertumbuhan penjualan stabil di seluruh kategori.',
        },
        status: 'pending',
      },
      { host: 'Excel' }
    );

    expect(res.success).toBe(true);
    expect(res.result).toContain('Universal Hub');
    expect(res.result).toContain('Penjualan Semester 1');

    // Verify snapshot was saved in Universal Hub and accessible to Word/PowerPoint
    const saved = getLatestCrossAppSnapshot('Word');
    expect(saved).not.toBeNull();
    expect(saved?.title).toBe('Penjualan Semester 1');
    expect(saved?.sourceHost).toBe('Excel');
    expect(saved?.artifactType).toBe('table_data');
    expect(saved?.tableData?.headers).toEqual(['Kategori', 'Q1', 'Q2']);
    expect(saved?.tableData?.rows).toEqual([
      ['Elektronik', 500, 700],
      ['Pakaian', 300, 450],
    ]);
    expect(saved?.summaryText).toBe('Pertumbuhan penjualan stabil di seluruh kategori.');
  });

  it('import_from_cross_app_hub in WordAgent retrieves snapshot from Excel and inserts table & content into Word', async () => {
    saveCrossAppSnapshot({
      sourceHost: 'Excel',
      title: 'Data Operasional',
      artifactType: 'table_data',
      tableData: {
        headers: ['Unit', 'Skor'],
        rows: [['A', 90]],
        totalRows: 1,
      },
      summaryText: 'Unit A meraih performa optimal.',
    });

    const mockDriver = new MockOfficeDriver();
    const agent = new WordAgent(mockDriver);
    const tools = agent.getTools();
    expect(tools.some(t => t.name === 'import_from_cross_app_hub')).toBe(true);

    const res = await agent.executeTool(
      {
        id: 'w-hub-1',
        name: 'import_from_cross_app_hub',
        arguments: {},
        status: 'pending',
      },
      { host: 'Word' }
    );

    expect(res.success).toBe(true);
    expect(res.result).toContain('Data Operasional');

    const outline = await mockDriver.getWordOutline();
    expect(outline).toContain('Data Operasional');
    expect(outline).toContain('Unit A meraih performa optimal.');
    expect(outline).toContain('[Tabel 2x2]');
  });

  it('share_to_cross_app_hub in WordAgent saves document summary to Hub', async () => {
    const mockDriver = new MockOfficeDriver();
    await mockDriver.insertContent('end', 'Dokumen PRD Sistem Pembayaran Digital terpadu.', 'paragraph');

    const agent = new WordAgent(mockDriver);
    const tools = agent.getTools();
    expect(tools.some(t => t.name === 'share_to_cross_app_hub')).toBe(true);

    const res = await agent.executeTool(
      {
        id: 'w-share-1',
        name: 'share_to_cross_app_hub',
        arguments: {
          title: 'PRD Sistem Pembayaran',
          summaryText: 'Spesifikasi teknis integrasi gateway pembayaran QRIS dan Virtual Account.',
        },
        status: 'pending',
      },
      { host: 'Word' }
    );

    expect(res.success).toBe(true);
    expect(res.result).toContain('Universal Hub');
    expect(res.result).toContain('PRD Sistem Pembayaran');

    // Retrieve in PowerPoint
    const snapshot = getLatestCrossAppSnapshot('PowerPoint');
    expect(snapshot).not.toBeNull();
    expect(snapshot?.sourceHost).toBe('Word');
    expect(snapshot?.title).toBe('PRD Sistem Pembayaran');
    expect(snapshot?.summaryText).toContain('gateway pembayaran QRIS');
  });

  it('import_from_cross_app_hub in PPTAgent retrieves snapshot and invokes driver.transformDocToDeck to create slide deck', async () => {
    saveCrossAppSnapshot({
      sourceHost: 'Excel',
      title: 'Ringkasan Kinerja Tahunan',
      artifactType: 'table_data',
      tableData: {
        headers: ['Divisi', 'Capaian'],
        rows: [['Pemasaran', '110%']],
        totalRows: 1,
      },
      summaryText: 'Tantangan: Keterbatasan anggaran. Solusi: Optimalisasi saluran digital. Metrik: Capaian 110%. Rencana: Ekspansi pasar Q1.',
    });

    const mockDriver = new MockOfficeDriver();
    const agent = new PPTAgent(mockDriver);
    const tools = agent.getTools();
    expect(tools.some(t => t.name === 'import_from_cross_app_hub')).toBe(true);

    const res = await agent.executeTool(
      {
        id: 'ppt-hub-1',
        name: 'import_from_cross_app_hub',
        arguments: { theme: 'corporate_blue' },
        status: 'pending',
      },
      { host: 'PowerPoint' }
    );

    expect(res.success).toBe(true);
    expect(res.result).toContain('Ringkasan Kinerja Tahunan');
    expect(res.result).toContain('Naskah Presenter');
    expect(mockDriver.slides.length).toBeGreaterThanOrEqual(4);
  });
});
