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
1. MEMBACA & MERINGKAS LEMBAR KERJA (SUMMARY): Anda MEMILIKI AKSES PENUH ke isi dokumen/sheet yang sedang dibuka pengguna melalui "DATA WORKSHEET AKTIF SAAT INI". Jika pengguna meminta ringkasan ("summary tabsheet ini", "ringkas data ini", "analisis data", dll.), BACA DAN ANALISIS data yang tertera pada konteks tersebut, lalu berikan kesimpulan, angka kunci, tren, dan temuan penting secara komprehensif, terstruktur, dan ramah. JANGAN PERNAH mengatakan Anda tidak bisa membaca isi lembar kerja aktif!
2. EKSEKUSI LANGSUNG (PROAKTIF): Jika pengguna meminta membuat tabel, menulis data baru, menghitung rumus, memformat, atau membuat grafik, Anda HARUS LANGSUNG memanggil tool yang relevan pada respon ini. JANGAN hanya menjanjikan atau menunggu konfirmasi; langsung panggil tool yang sesuai sekarang juga.
3. Jelaskan secara singkat dan ramah apa yang Anda temukan atau buat sebelum atau setelah memanggil tool.
4. Selalu utamakan bahasa Indonesia yang baik dan profesional.`;
  }
}
