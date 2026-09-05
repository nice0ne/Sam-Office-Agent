import { getOfficeDriver } from '../../services/office';
import { AgentContext, IAgent } from '../types';
import { ToolCall, ToolDefinition } from '../../types';

export class PPTAgent implements IAgent {
  id = 'ppt-specialist';
  name = 'PowerPoint Specialist';
  hostType = 'PowerPoint' as const;

  getSystemPrompt(_context: AgentContext): string {
    return `Anda adalah Sam Office Agent spesialis Microsoft PowerPoint.
Anda ahli dalam merancang struktur presentasi yang memikat, menyusun poin slide yang ringkas, dan membuat speaker notes untuk presentasi.`;
  }

  getTools(): ToolDefinition[] {
    return [
      {
        name: 'add_slide',
        description: 'Menambahkan slide baru ke presentasi.',
        parameters: {
          type: 'object',
          properties: {
            layout: { type: 'string', enum: ['TitleOnly', 'TitleAndContent', 'TwoContent'], description: 'Layout slide' },
          },
          required: ['layout'],
        },
      },
      {
        name: 'insert_slide_content',
        description: 'Mengisi judul slide dan poin-poin materi.',
        parameters: {
          type: 'object',
          properties: {
            title: { type: 'string', description: 'Judul slide' },
            bullets: { type: 'array', items: { type: 'string' }, description: 'Poin-poin presentasi' },
          },
          required: ['title', 'bullets'],
        },
      },
      {
        name: 'set_speaker_notes',
        description: 'Menambahkan catatan pemateri pada slide aktif.',
        parameters: {
          type: 'object',
          properties: {
            notes: { type: 'string', description: 'Catatan pemateri' },
          },
          required: ['notes'],
        },
      },
    ];
  }

  async executeTool(toolCall: ToolCall, _context: AgentContext): Promise<{ success: boolean; result?: any; error?: string }> {
    const driver = getOfficeDriver('PowerPoint');
    try {
      if (toolCall.name === 'add_slide') {
        const res = await driver.addSlide(toolCall.arguments.layout);
        return { success: true, result: `Slide #${res.slideNumber} berhasil ditambahkan.` };
      }
      if (toolCall.name === 'insert_slide_content') {
        const { title, bullets } = toolCall.arguments;
        await driver.insertSlideContent(title, bullets);
        return { success: true, result: 'Konten slide berhasil diisi.' };
      }
      if (toolCall.name === 'set_speaker_notes') {
        await driver.setSpeakerNotes(toolCall.arguments.notes);
        return { success: true, result: 'Speaker notes berhasil disimpan.' };
      }
      return { success: false, error: 'Tool tidak ditemukan.' };
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }
}
