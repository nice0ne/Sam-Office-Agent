import { describe, it, expect } from 'vitest';
import { WordAgent } from '../src/agents/word/wordAgent';
import { PPTAgent } from '../src/agents/powerpoint/pptAgent';
import { MockOfficeDriver } from '../src/services/office/mockDriver';

describe('Specialist Agents Suite', () => {
  it('provides insert_process_flowchart tool in WordAgent and executes successfully', async () => {
    const mockDriver = new MockOfficeDriver();
    const wordAgent = new WordAgent(mockDriver);

    const tools = wordAgent.getTools();
    expect(tools.some(t => t.name === 'insert_process_flowchart')).toBe(true);

    const res = await wordAgent.executeTool(
      {
        id: 'tool-flow-1',
        name: 'insert_process_flowchart',
        arguments: {
          textOrSteps: '1. Pengajuan berkas -> 2. Cek kelengkapan -> 3. Terbitkan surat',
          title: 'Alur Penerbitan Surat',
          theme: 'corporate_navy',
        },
        status: 'pending',
      },
      { host: 'Word' }
    );

    expect(res.success).toBe(true);
    expect(res.result).toContain('Diagram Alur Proses Berhasil Disisipkan');
    expect(res.result).toContain('Alur Penerbitan Surat');
    expect(mockDriver.diagrams.length).toBe(1);
  });

  it('provides insert_process_flowchart tool in PPTAgent and executes successfully', async () => {
    const mockDriver = new MockOfficeDriver();
    const pptAgent = new PPTAgent(mockDriver);

    const tools = pptAgent.getTools();
    expect(tools.some(t => t.name === 'insert_process_flowchart')).toBe(true);

    const res = await pptAgent.executeTool(
      {
        id: 'tool-flow-ppt-1',
        name: 'insert_process_flowchart',
        arguments: {
          textOrSteps: '1. Inisiasi Proyek -> 2. Analisis Kebutuhan -> 3. Eksekusi -> 4. Evaluasi Akhir',
          title: 'Siklus Proyek',
          theme: 'emerald_executive',
        },
        status: 'pending',
      },
      { host: 'PowerPoint' }
    );

    expect(res.success).toBe(true);
    expect(res.result).toContain('Diagram Alur Proses Berhasil Disisipkan');
    expect(mockDriver.diagrams.length).toBe(1);
  });
});
