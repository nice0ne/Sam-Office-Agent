import { getOfficeDriver } from '../../services/office';
import { AgentContext, IAgent } from '../types';
import { ToolCall, ToolDefinition } from '../../types';

export class WordAgent implements IAgent {
  id = 'word-specialist';
  name = 'Word Specialist';
  hostType = 'Word' as const;

  getSystemPrompt(_context: AgentContext): string {
    return `Anda adalah Sam Office Agent spesialis Microsoft Word.
Anda ahli dalam penyusunan surat resmi, laporan profesional, perbaikan tata bahasa/proofreading, pemformatan heading dokumen, dan tabel terstruktur.`;
  }

  getTools(): ToolDefinition[] {
    return [
      {
        name: 'insert_content',
        description: 'Menyisipkan paragraf, heading (H1/H2), atau poin di dokumen.',
        parameters: {
          type: 'object',
          properties: {
            position: { type: 'string', enum: ['start', 'end', 'cursor'], description: 'Posisi penyisipan' },
            text: { type: 'string', description: 'Isi teks' },
            type: { type: 'string', enum: ['paragraph', 'h1', 'h2', 'bullet'], description: 'Tipe konten' },
          },
          required: ['position', 'text', 'type'],
        },
      },
      {
        name: 'replace_selection',
        description: 'Mengganti teks yang sedang diblok/diseleksi pengguna dengan teks baru yang telah disempurnakan.',
        parameters: {
          type: 'object',
          properties: {
            newText: { type: 'string', description: 'Teks baru pengganti' },
          },
          required: ['newText'],
        },
      },
      {
        name: 'insert_table',
        description: 'Menyisipkan tabel data di dokumen Word.',
        parameters: {
          type: 'object',
          properties: {
            rows: { type: 'number', description: 'Jumlah baris' },
            cols: { type: 'number', description: 'Jumlah kolom' },
            data: { type: 'array', description: 'Array 2D data teks' },
          },
          required: ['rows', 'cols'],
        },
      },
    ];
  }

  async executeTool(toolCall: ToolCall, _context: AgentContext): Promise<{ success: boolean; result?: any; error?: string }> {
    const driver = getOfficeDriver('Word');
    try {
      if (toolCall.name === 'insert_content') {
        const { position, text, type } = toolCall.arguments;
        await driver.insertContent(position, text, type);
        return { success: true, result: 'Konten berhasil disisipkan.' };
      }
      if (toolCall.name === 'replace_selection') {
        const { newText } = toolCall.arguments;
        await driver.replaceSelection(newText);
        return { success: true, result: 'Seleksi berhasil diganti.' };
      }
      if (toolCall.name === 'insert_table') {
        const { rows, cols, data } = toolCall.arguments;
        await driver.insertTable(rows, cols, data);
        return { success: true, result: `Tabel ${rows}x${cols} berhasil dibuat.` };
      }
      return { success: false, error: 'Tool tidak ditemukan.' };
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }
}
