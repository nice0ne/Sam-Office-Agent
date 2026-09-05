import { getOfficeDriver } from '../../services/office';
import { validateExcelFormula } from '../../utils/formulaValidator';
import { AgentContext, IAgent } from '../types';
import { ToolCall, ToolDefinition } from '../../types';

export class ExcelAgent implements IAgent {
  id = 'excel-specialist';
  name = 'Excel Specialist';
  hostType = 'Excel' as const;

  getSystemPrompt(context: AgentContext): string {
    return `Anda adalah Sam Office Agent spesialis Microsoft Excel.
Anda ahli dalam analisis data, formula spreadsheet kompleks (SUM, AVERAGE, XLOOKUP, INDEX, MATCH), pemformatan sel, dan pembuatan grafik.
Jika pengguna meminta membuat tabel data: LANGSUNG panggil tool \`write_cells\` dengan data lengkap (gunakan parameter \`range\` seperti "A1:C11" dan parameter \`values\` berupa array 2D baris dan kolom). Jika data real-time tidak tersedia, buat data ilustrasi yang realistis dan langsung tulis ke sel.
Jika pengguna meminta menulis formula, pastikan sintaks formula valid dan diawali '='.
Gunakan \`format_range\` untuk memberi style pada header (bold, fillColor) dan format angka (numberFormat).
Gunakan \`create_chart\` jika diminta grafik visualisasi data.
Konteks saat ini: ${context.activeCellOrRange || 'Sheet aktif'}.`;
  }

  getTools(): ToolDefinition[] {
    return [
      {
        name: 'write_cells',
        description: 'Menulis nilai atau formula ke range sel Excel tertentu (misal: A1:B10).',
        parameters: {
          type: 'object',
          properties: {
            range: { type: 'string', description: 'Alamat range, misal: "E2:E15"' },
            formula: { type: 'string', description: 'Formula Excel (opsional), misal: "=SUM(B2:D2)"' },
            values: { type: 'array', description: 'Array 2D nilai data jika menulis data statis' },
          },
          required: ['range'],
        },
      },
      {
        name: 'format_range',
        description: 'Mengatur format sel (tebal, warna latar zebra, format mata uang/persentase).',
        parameters: {
          type: 'object',
          properties: {
            range: { type: 'string', description: 'Alamat range, misal: "E2:E15"' },
            bold: { type: 'boolean', description: 'Tebalkan teks' },
            fillColor: { type: 'string', description: 'Warna latar HEX, misal: "#F0F9FF"' },
            numberFormat: { type: 'string', description: 'Format angka, misal: "Rp#,##0"' },
          },
          required: ['range'],
        },
      },
      {
        name: 'create_chart',
        description: 'Membuat grafik otomatis berdasarkan range data.',
        parameters: {
          type: 'object',
          properties: {
            chartType: { type: 'string', enum: ['ColumnClustered', 'Line', 'Pie', 'BarClustered'], description: 'Jenis grafik' },
            dataRange: { type: 'string', description: 'Range data grafik, misal: "A1:E15"' },
            title: { type: 'string', description: 'Judul grafik' },
          },
          required: ['chartType', 'dataRange'],
        },
      },
    ];
  }

  async executeTool(toolCall: ToolCall, _context: AgentContext): Promise<{ success: boolean; result?: any; error?: string }> {
    const driver = getOfficeDriver('Excel');
    try {
      if (toolCall.name === 'write_cells') {
        const { range, formula, values } = toolCall.arguments;
        const hasValues = Array.isArray(values) ? values.length > 0 : Boolean(values);
        if (!formula && !hasValues) {
          return { success: false, error: 'Harus menyertakan setidaknya formula atau values untuk range.' };
        }
        if (formula) {
          const validation = validateExcelFormula(formula);
          if (!validation.isValid) {
            return { success: false, error: validation.error };
          }
          await driver.writeCells(range, undefined, [[formula]]);
        } else if (values) {
          await driver.writeCells(range, values);
        }
        return { success: true, result: `Berhasil menulis ke sel ${range}` };
      }

      if (toolCall.name === 'format_range') {
        const { range, ...styles } = toolCall.arguments;
        await driver.formatRange(range, styles);
        return { success: true, result: `Format berhasil diterapkan ke ${range}` };
      }

      if (toolCall.name === 'create_chart') {
        const { chartType, dataRange, title } = toolCall.arguments;
        await driver.createChart(chartType, dataRange, title);
        return { success: true, result: `Grafik ${chartType} berhasil dibuat!` };
      }

      return { success: false, error: `Tool ${toolCall.name} tidak dikenali.` };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  }
}
