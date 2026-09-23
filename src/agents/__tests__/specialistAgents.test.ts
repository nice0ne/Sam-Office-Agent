import { describe, it, expect, beforeEach } from 'vitest';
import { ExcelAgent } from '../excel/excelAgent';
import { WordAgent } from '../word/wordAgent';
import { PPTAgent } from '../powerpoint/pptAgent';
import { resetOfficeDriver, setOfficeDriver, MockOfficeDriver } from '../../services/office';

describe('Specialist Agents Tools & Prompts', () => {
  beforeEach(() => {
    resetOfficeDriver();
  });

  it('ExcelAgent exposes Excel specific tools', () => {
    const agent = new ExcelAgent();
    const tools = agent.getTools();
    const toolNames = tools.map(t => t.name);
    expect(toolNames).toContain('read_sheet');
    expect(toolNames).toContain('write_cells');
    expect(toolNames).toContain('format_range');
    expect(toolNames).toContain('create_chart');
    expect(toolNames).toContain('clean_data');
    expect(toolNames).toContain('apply_conditional_formatting');
    expect(toolNames).toContain('audit_sheet_data');
    expect(toolNames).toContain('generate_data_story');
    expect(agent.id).toBe('excel-specialist');
    expect(agent.hostType).toBe('Excel');
  });

  it('WordAgent exposes Word specific tools', () => {
    const agent = new WordAgent();
    const tools = agent.getTools();
    const toolNames = tools.map(t => t.name);
    expect(toolNames).toContain('insert_content');
    expect(toolNames).toContain('replace_selection');
    expect(toolNames).toContain('insert_table');
    expect(toolNames).toContain('generate_structured_doc');
    expect(toolNames).toContain('polish_document_text');
    expect(agent.id).toBe('word-specialist');
    expect(agent.hostType).toBe('Word');
  });

  it('PPTAgent exposes PowerPoint specific tools', () => {
    const agent = new PPTAgent();
    const tools = agent.getTools();
    const toolNames = tools.map(t => t.name);
    expect(toolNames).toContain('add_slide');
    expect(toolNames).toContain('insert_slide_content');
    expect(toolNames).toContain('set_speaker_notes');
    expect(toolNames).toContain('generate_themed_deck');
    expect(agent.id).toBe('ppt-specialist');
    expect(agent.hostType).toBe('PowerPoint');
  });

  describe('ExcelAgent execution', () => {
    const agent = new ExcelAgent();
    const context = { host: 'Excel' as const, activeCellOrRange: 'B2:C10' };

    it('generates system prompt including active range', () => {
      const prompt = agent.getSystemPrompt(context);
      expect(prompt).toContain('Microsoft Excel');
      expect(prompt).toContain('B2:C10');
    });

    it('writes cells with valid formula', async () => {
      const result = await agent.executeTool(
        {
          id: 'call-1',
          name: 'write_cells',
          arguments: { range: 'D2:D10', formula: '=SUM(B2:C2)' },
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(true);
      expect(result.result).toContain('Berhasil menulis ke sel D2:D10');
    });

    it('rejects write_cells with invalid formula', async () => {
      const result = await agent.executeTool(
        {
          id: 'call-2',
          name: 'write_cells',
          arguments: { range: 'D2:D10', formula: 'SUM(B2:C2)' },
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(false);
      expect(result.error).toContain('Formula Excel harus diawali dengan tanda sama dengan (=).');
    });

    it('writes cells with values array', async () => {
      const result = await agent.executeTool(
        {
          id: 'call-3',
          name: 'write_cells',
          arguments: { range: 'A1:B2', values: [[1, 2], [3, 4]] },
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(true);
      expect(result.result).toContain('Berhasil menulis ke sel A1:B2');
    });

    it('rejects write_cells when neither formula nor values is provided', async () => {
      const result = await agent.executeTool(
        {
          id: 'call-guard-1',
          name: 'write_cells',
          arguments: { range: 'A1:B2' },
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(false);
      expect(result.error).toBe('Harus menyertakan setidaknya formula atau values untuk range.');
    });

    it('reads sheet data and summary via read_sheet tool', async () => {
      const result = await agent.executeTool(
        {
          id: 'call-read-1',
          name: 'read_sheet',
          arguments: { range: 'A1:C5' },
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(true);
      expect(result.result).toBeDefined();
      expect(result.result.values.length).toBeGreaterThan(0);
      expect(result.result.sheetName).toBe('Sheet1');
    });

    it('rejects write_cells when values array is empty', async () => {
      const result = await agent.executeTool(
        {
          id: 'call-guard-2',
          name: 'write_cells',
          arguments: { range: 'A1:B2', values: [] },
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(false);
      expect(result.error).toBe('Harus menyertakan setidaknya formula atau values untuk range.');
    });

    it('formats range', async () => {
      const result = await agent.executeTool(
        {
          id: 'call-4',
          name: 'format_range',
          arguments: { range: 'A1:D1', bold: true, fillColor: '#00FF00' },
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(true);
      expect(result.result).toContain('Format berhasil diterapkan ke A1:D1');
    });

    it('creates chart', async () => {
      const result = await agent.executeTool(
        {
          id: 'call-5',
          name: 'create_chart',
          arguments: { chartType: 'ColumnClustered', dataRange: 'A1:B10', title: 'Sales Chart' },
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(true);
      expect(result.result).toContain('Grafik ColumnClustered berhasil dibuat!');
    });

    it('executes clean_data tool and removes duplicate rows', async () => {
      const mockDriver = new MockOfficeDriver();
      await mockDriver.writeCells('A1:B4', [
        ['Nama', 'Kota'],
        ['Andi', 'Jakarta '],
        ['Budi', 'Bandung'],
        ['Andi', 'Jakarta '],
      ]);
      setOfficeDriver(mockDriver);

      const res = await agent.executeTool(
        {
          id: 'c1',
          name: 'clean_data',
          arguments: { range: 'A1:B4', removeDuplicates: true, trimWhitespace: true },
          status: 'pending',
        },
        context
      );

      expect(res.success).toBe(true);
      expect(res.result).toBeDefined();
      expect(res.result).toContain('1 baris duplikat dihapus');
      expect(res.result).toContain('2 sel dirapikan');

      const updated = await mockDriver.readActiveRange();
      expect(updated.values.length).toBe(3);
      expect(updated.values[1][1]).toBe('Jakarta');
    });

    it('executes clean_data tool with default arguments and fills empty cells', async () => {
      const mockDriver = new MockOfficeDriver();
      await mockDriver.writeCells('A1:B3', [
        ['Nama', 'Skor'],
        ['Andi', ''],
        ['Budi', 90],
      ]);
      setOfficeDriver(mockDriver);

      const res = await agent.executeTool(
        {
          id: 'c2',
          name: 'clean_data',
          arguments: { range: 'A1:B3', fillEmptyValues: 0 },
          status: 'pending',
        },
        context
      );

      expect(res.success).toBe(true);
      expect(res.result).toContain('1 sel kosong diisi');

      const updated = await mockDriver.readActiveRange();
      expect(updated.values[1][1]).toBe(0);
    });

    it('executes apply_conditional_formatting tool', async () => {
      const res = await agent.executeTool(
        {
          id: 'cf1',
          name: 'apply_conditional_formatting',
          arguments: { range: 'A1:A10', type: 'color_scale' },
          status: 'pending',
        },
        context
      );

      expect(res.success).toBe(true);
      expect(res.result).toContain('color_scale');
    });

    it('fails apply_conditional_formatting when range or type is missing', async () => {
      const res = await agent.executeTool(
        {
          id: 'cf-err',
          name: 'apply_conditional_formatting',
          arguments: { range: 'A1:A10' },
          status: 'pending',
        },
        context
      );

      expect(res.success).toBe(false);
      expect(res.error).toContain('wajib diisi');
    });

    it('executes audit_sheet_data tool and returns markdown audit report', async () => {
      const mockDriver = new MockOfficeDriver();
      mockDriver.mockData = [
        ['Product', 'Qty', 'Price', 'Total'],
        ['A', 10, 5, 50],
        ['B', 2, 0, '#DIV/0!'],
      ];
      const testAgent = new ExcelAgent(mockDriver);
      const tools = testAgent.getTools();
      expect(tools.some(t => t.name === 'audit_sheet_data')).toBe(true);

      const res = await testAgent.executeTool(
        { id: 'c1', name: 'audit_sheet_data', arguments: {}, status: 'pending' },
        context
      );
      expect(res.success).toBe(true);
      expect(res.result).toContain('Total Error Formula');
      expect(res.result).toContain('#DIV/0!');
      expect(res.result).toContain('D3');
    });

    it('executes generate_data_story tool and returns executive data narrative', async () => {
      const mockDriver = new MockOfficeDriver();
      mockDriver.mockData = [
        ['Category', 'Revenue'],
        ['Laptops', 50000000],
        ['Phones', 75000000],
      ];
      const testAgent = new ExcelAgent(mockDriver);
      const tools = testAgent.getTools();
      expect(tools.some(t => t.name === 'generate_data_story')).toBe(true);

      const res = await testAgent.executeTool(
        { id: 'c2', name: 'generate_data_story', arguments: { focusMetric: 'Revenue' }, status: 'pending' },
        context
      );
      expect(res.success).toBe(true);
      expect(res.result).toContain('Headline');
      expect(res.result).toContain('Metrik Utama');
      expect(res.result).toContain('Laptops');
    });

    it('returns error for unknown tool', async () => {
      const result = await agent.executeTool(
        {
          id: 'call-unknown',
          name: 'unknown_tool',
          arguments: {},
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(false);
      expect(result.error).toContain('Tool unknown_tool tidak dikenali.');
    });

    it('handles driver error cleanly', async () => {
      const mockDriver = new MockOfficeDriver();
      mockDriver.writeCells = async () => {
        throw new Error('Write failed in driver');
      };
      setOfficeDriver(mockDriver);

      const result = await agent.executeTool(
        {
          id: 'call-err',
          name: 'write_cells',
          arguments: { range: 'A1', values: [['test']] },
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(false);
      expect(result.error).toBe('Write failed in driver');
    });
  });

  describe('WordAgent execution', () => {
    const agent = new WordAgent();
    const context = { host: 'Word' as const };

    it('generates system prompt', () => {
      const prompt = agent.getSystemPrompt(context);
      expect(prompt).toContain('Microsoft Word');
    });

    it('inserts content', async () => {
      const result = await agent.executeTool(
        {
          id: 'call-w1',
          name: 'insert_content',
          arguments: { position: 'end', text: 'Hello Word', type: 'paragraph' },
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(true);
      expect(result.result).toContain('Konten berhasil disisipkan.');
    });

    it('replaces selection', async () => {
      const result = await agent.executeTool(
        {
          id: 'call-w2',
          name: 'replace_selection',
          arguments: { newText: 'Improved text' },
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(true);
      expect(result.result).toContain('Seleksi berhasil diganti.');
    });

    it('inserts table', async () => {
      const result = await agent.executeTool(
        {
          id: 'call-w3',
          name: 'insert_table',
          arguments: { rows: 3, cols: 2, data: [['A', 'B'], ['C', 'D'], ['E', 'F']] },
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(true);
      expect(result.result).toContain('Tabel 3x2 berhasil dibuat.');
    });

    it('returns error for unknown tool', async () => {
      const result = await agent.executeTool(
        {
          id: 'call-w-unknown',
          name: 'unknown_tool',
          arguments: {},
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(false);
      expect(result.error).toContain('Tool tidak ditemukan.');
    });

    it('handles driver error cleanly', async () => {
      const mockDriver = new MockOfficeDriver();
      mockDriver.insertContent = async () => {
        throw new Error('Word insertion failed');
      };
      setOfficeDriver(mockDriver);

      const result = await agent.executeTool(
        {
          id: 'call-w-err',
          name: 'insert_content',
          arguments: { position: 'end', text: 'fail', type: 'paragraph' },
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(false);
      expect(result.error).toBe('Word insertion failed');
    });

    it('executes generate_structured_doc in WordAgent', async () => {
      const res = await agent.executeTool(
        {
          id: 'w1',
          name: 'generate_structured_doc',
          arguments: {
            templateType: 'SOP',
            title: 'SOP Pengajuan Cuti',
            sections: [
              { heading: 'Tujuan', content: 'Standarisasi pengajuan cuti tahunan karyawan.' },
              { heading: 'Prosedur', bullets: ['Isi form cuti', 'Minta persetujuan atasan'] },
            ],
          },
          status: 'pending',
        },
        context
      );

      expect(res.success).toBe(true);
      expect(res.result).toContain('SOP');
      expect(res.result).toContain('SOP Pengajuan Cuti');
    });

    it('fails generate_structured_doc when required arguments are missing', async () => {
      const res = await agent.executeTool(
        {
          id: 'w-err-missing',
          name: 'generate_structured_doc',
          arguments: {
            templateType: 'SOP',
          },
          status: 'pending',
        },
        context
      );

      expect(res.success).toBe(false);
      expect(res.error).toContain('wajib diisi');
    });

    it('executes polish_document_text in WordAgent', async () => {
      const res = await agent.executeTool(
        {
          id: 'w2',
          name: 'polish_document_text',
          arguments: {
            scope: 'selection',
            tone: 'formal_indonesia',
            customText: 'Yth. Bapak/Ibu Direksi, berikut kami sampaikan laporan kinerja.',
          },
          status: 'pending',
        },
        context
      );

      expect(res.success).toBe(true);
      expect(res.result).toContain('formal_indonesia');
    });

    it('provides review_compliance_clauses tool and executes it successfully', async () => {
      const mockDriver = new MockOfficeDriver();
      mockDriver.mockWordBody = 'Perjanjian Jasa Vendor. Pihak kedua menanggung seluruh ganti rugi tanpa batas untuk segala tuntutan.';
      const testAgent = new WordAgent(mockDriver);
      const tools = testAgent.getTools();
      expect(tools.some(t => t.name === 'review_compliance_clauses')).toBe(true);

      const res = await testAgent.executeTool(
        { id: 'w-comp', name: 'review_compliance_clauses', arguments: { contractType: 'vendor_service' }, status: 'pending' },
        context
      );
      expect(res.success).toBe(true);
      expect(res.result).toContain('Review Kepatuhan Kontrak');
      expect(res.result).toContain('TINGGI');
      expect(res.result).toContain('LIABILITY_INDEMNITY');
      expect(res.result).toContain('force_majeure');
    });

    it('provides apply_corporate_style tool and executes it successfully', async () => {
      const mockDriver = new MockOfficeDriver();
      mockDriver.mockWordBody = 'Judul Dokumen Resmi\nBab 1 Pendahuluan\nIsi naskah dokumen korporat.';
      const testAgent = new WordAgent(mockDriver);
      const tools = testAgent.getTools();
      expect(tools.some(t => t.name === 'apply_corporate_style')).toBe(true);

      const res = await testAgent.executeTool(
        { id: 'w-style', name: 'apply_corporate_style', arguments: { theme: 'corporate_navy' }, status: 'pending' },
        context
      );
      expect(res.success).toBe(true);
      expect(res.result).toContain('Format Brand Korporat Diterapkan');
      expect(res.result).toContain('corporate_navy');
    });
  });

  describe('PPTAgent execution', () => {
    const agent = new PPTAgent();
    const context = { host: 'PowerPoint' as const };

    it('generates system prompt', () => {
      const prompt = agent.getSystemPrompt(context);
      expect(prompt).toContain('Microsoft PowerPoint');
    });

    it('adds slide', async () => {
      const result = await agent.executeTool(
        {
          id: 'call-p1',
          name: 'add_slide',
          arguments: { layout: 'TitleAndContent' },
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(true);
      expect(result.result).toContain('Slide #1 berhasil ditambahkan.');
    });

    it('inserts slide content', async () => {
      const result = await agent.executeTool(
        {
          id: 'call-p2',
          name: 'insert_slide_content',
          arguments: { title: 'Agenda', bullets: ['Point 1', 'Point 2'] },
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(true);
      expect(result.result).toContain('Konten slide berhasil diisi.');
    });

    it('sets speaker notes', async () => {
      const result = await agent.executeTool(
        {
          id: 'call-p3',
          name: 'set_speaker_notes',
          arguments: { notes: 'Remember to smile' },
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(true);
      expect(result.result).toContain('Speaker notes berhasil disimpan.');
    });

    it('returns error for unknown tool', async () => {
      const result = await agent.executeTool(
        {
          id: 'call-p-unknown',
          name: 'unknown_tool',
          arguments: {},
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(false);
      expect(result.error).toContain('Tool tidak ditemukan.');
    });

    it('handles driver error cleanly', async () => {
      const mockDriver = new MockOfficeDriver();
      mockDriver.addSlide = async () => {
        throw new Error('PPT slide addition failed');
      };
      setOfficeDriver(mockDriver);

      const result = await agent.executeTool(
        {
          id: 'call-p-err',
          name: 'add_slide',
          arguments: { layout: 'TitleOnly' },
          status: 'pending',
        },
        context
      );
      expect(result.success).toBe(false);
      expect(result.error).toBe('PPT slide addition failed');
    });

    it('executes generate_themed_deck in PPTAgent', async () => {
      const res = await agent.executeTool(
        {
          id: 'p1',
          name: 'generate_themed_deck',
          arguments: {
            topic: 'Laporan Q3',
            theme: 'corporate_blue',
            slides: [
              { title: 'Cover', layout: 'title_cover', content: ['Q3 Performance Overview'] },
              { title: 'Metrik Kunci', layout: 'metric_highlights', metrics: [{ label: 'Pertumbuhan', value: '+25%' }] },
            ],
          },
          status: 'pending',
        },
        context
      );

      expect(res.success).toBe(true);
      expect(res.result).toContain('Laporan Q3');
      expect(res.result).toContain('corporate_blue');
    });

    it('fails generate_themed_deck when missing required arguments', async () => {
      const res = await agent.executeTool(
        {
          id: 'p-err-missing',
          name: 'generate_themed_deck',
          arguments: {
            topic: 'Laporan Q3',
            slides: [],
          },
          status: 'pending',
        },
        context
      );

      expect(res.success).toBe(false);
      expect(res.error).toContain('wajib diisi');
    });
  });
});
