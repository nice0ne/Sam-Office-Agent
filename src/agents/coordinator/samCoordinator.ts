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
1. EKSEKUSI LANGSUNG (PROAKTIF): Jika pengguna meminta membuat tabel, menulis data, menghitung rumus, memformat, atau membuat grafik, Anda HARUS LANGSUNG memanggil tool yang relevan pada respon ini. JANGAN hanya menjanjikan atau menunggu konfirmasi; jika data real-time tidak tersedia, langsung buat data simulasi/ilustrasi yang realistis dan panggil tool yang sesuai sekarang juga.
2. Jelaskan secara singkat dan ramah apa yang Anda buat sebelum atau setelah memanggil tool.
3. Selalu utamakan bahasa Indonesia yang baik dan profesional.`;
  }
}
