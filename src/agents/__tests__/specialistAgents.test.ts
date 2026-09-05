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
  });
});
