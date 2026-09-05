import { HostType } from '../../types';
import { ExcelAgent } from '../excel/excelAgent';
import { PPTAgent } from '../powerpoint/pptAgent';
import { AgentContext, IAgent } from '../types';
import { WordAgent } from '../word/wordAgent';

export class SamCoordinator {
  private excelAgent = new ExcelAgent();
  private wordAgent = new WordAgent();
  private pptAgent = new PPTAgent();

  getSpecialist(host: HostType): IAgent {
    switch (host) {
      case 'Excel':
        return this.excelAgent;
      case 'Word':
        return this.wordAgent;
      case 'PowerPoint':
        return this.pptAgent;
      default:
        return this.excelAgent; // Default fallback for browser testing
    }
  }

  buildSystemPrompt(host: HostType, context: AgentContext): string {
    const specialist = this.getSpecialist(host);
    return `Anda adalah Sam, asisten AI produktivitas kantor yang ramah, profesional, dan cekatan.
${specialist.getSystemPrompt(context)}

Aturan Penting:
1. Jika pengguna meminta tindakan langsung pada dokumen (seperti mengisi rumus, mengedit teks, membuat slide), gunakan tool yang tersedia.
2. Jelaskan secara singkat dan ramah apa yang Anda lakukan sebelum atau setelah memanggil tool.
3. Selalu utamakan bahasa Indonesia yang baik dan profesional.`;
  }
}
