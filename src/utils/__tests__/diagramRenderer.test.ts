import { describe, it, expect } from 'vitest';
import { synthesizeFlowchartFromText, renderFlowchartToPngBase64 } from '../diagramRenderer';
import { FlowchartDefinition } from '../../services/office/types';

describe('diagramRenderer', () => {
  it('synthesizes flowchart definition from structured SOP text', () => {
    const rawSOP = `
    1. Mulai: Karyawan mengajukan permohonan cuti tahunan melalui portal.
    2. Apakah sisa kuota cuti mencukupi? Jika tidak, tolak pengajuan.
    3. Manajer memvalidasi dan menyetujui jadwal cuti.
    4. HRD mencatat dokumen dan memotong kuota cuti.
    5. Selesai: Karyawan menerima notifikasi persetujuan cuti.
    `;

    const def = synthesizeFlowchartFromText(rawSOP, {
      title: 'Alur Pengajuan Cuti',
      theme: 'corporate_navy',
    });

    expect(def.title).toBe('Alur Pengajuan Cuti');
    expect(def.nodes.length).toBeGreaterThanOrEqual(4);
    expect(def.edges.length).toBeGreaterThanOrEqual(3);

    // Node 0 should be start
    expect(def.nodes[0].type).toBe('start');
    // Decision node should be detected
    const decisionNode = def.nodes.find(n => n.type === 'decision');
    expect(decisionNode).toBeDefined();
    // Document node should be detected (from "dokumen")
    const documentNode = def.nodes.find(n => n.type === 'document');
    expect(documentNode).toBeDefined();
    // End node should be detected
    const endNode = def.nodes.find(n => n.type === 'end');
    expect(endNode).toBeDefined();
  });

  it('synthesizes default flowchart when text is empty or minimal', () => {
    const def = synthesizeFlowchartFromText('');
    expect(def.nodes.length).toBeGreaterThan(0);
    expect(def.edges.length).toBeGreaterThan(0);
  });

  it('renders flowchart definition to clean Base64 PNG string', () => {
    const sampleDef: FlowchartDefinition = {
      title: 'Proses Approval Anggaran',
      direction: 'TD',
      theme: 'emerald_executive',
      nodes: [
        { id: '1', label: 'Input RAB', type: 'start' },
        { id: '2', label: 'Verifikasi Finance', type: 'decision' },
        { id: '3', label: 'Disetujui Direksi', type: 'end' },
      ],
      edges: [
        { from: '1', to: '2' },
        { from: '2', to: '3', label: 'Ya' },
      ],
    };

    const base64 = renderFlowchartToPngBase64(sampleDef);
    expect(typeof base64).toBe('string');
    expect(base64.length).toBeGreaterThan(20);
    expect(base64).not.toContain('data:image/png;base64,');
  });

  it('supports LR direction and modern_dark theme rendering', () => {
    const sampleDef: FlowchartDefinition = {
      title: 'Workflow Horisontal',
      direction: 'LR',
      theme: 'modern_dark',
      nodes: [
        { id: 'start', label: 'Start Step', type: 'start' },
        { id: 'doc', label: 'Berkas Formulir', type: 'document' },
        { id: 'proc', label: 'Review Tim', type: 'process' },
        { id: 'dec', label: 'Validasi?', type: 'decision' },
        { id: 'end', label: 'Selesai', type: 'end' },
      ],
      edges: [
        { from: 'start', to: 'doc' },
        { from: 'doc', to: 'proc' },
        { from: 'proc', to: 'dec' },
        { from: 'dec', to: 'end', label: 'OK' },
      ],
    };

    const base64 = renderFlowchartToPngBase64(sampleDef, { direction: 'LR', theme: 'modern_dark' });
    expect(typeof base64).toBe('string');
    expect(base64.length).toBeGreaterThan(20);
    expect(base64).not.toContain('data:image/png;base64,');
  });
});
