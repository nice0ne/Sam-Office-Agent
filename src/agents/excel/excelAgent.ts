import { getOfficeDriver } from '../../services/office';
import { validateExcelFormula } from '../../utils/formulaValidator';
import { AgentContext, IAgent } from '../types';
import { ToolCall, ToolDefinition } from '../../types';
import { executeMetaOrCustomTool, getLearnedAndMetaTools } from '../metaTools';
import { saveCrossAppSnapshot } from '../../services/storage/crossAppBridge';
import { searchWeb } from '../../services/search';
import { addLearnedDirective } from '../../services/storage/soulStorage';
import { queryKnowledge } from '../../services/rag/ragEngine';

export class ExcelAgent implements IAgent {
  id = 'excel-specialist';
  name = 'Excel Specialist';
  hostType = 'Excel' as const;

  constructor(private driverOverride?: any) {}

  getSystemPrompt(context: AgentContext): string {
    return `Anda adalah Sam Office Agent spesialis Microsoft Excel.
Anda ahli dalam analisis data, formula spreadsheet kompleks (SUM, AVERAGE, XLOOKUP, INDEX, MATCH), pemformatan sel, dan pembuatan grafik.
Jika pengguna meminta membuat tabel data: LANGSUNG panggil tool \`write_cells\` dengan data lengkap (gunakan parameter \`range\` seperti "A1:C11" dan parameter \`values\` berupa array 2D baris dan kolom). Jika data real-time tidak tersedia, buat data ilustrasi yang realistis dan langsung tulis ke sel.
Jika pengguna meminta menulis formula, pastikan sintaks formula valid dan diawali '='.
Gunakan \`format_range\` untuk memberi style pada header (bold, fillColor) dan format angka (numberFormat).
Gunakan \`create_chart\` jika diminta grafik visualisasi data.
Konteks Sheet Saat Ini: ${context.activeCellOrRange || 'Sheet aktif'}.
${context.documentSummary ? `\n--- DATA WORKSHEET AKTIF SAAT INI ---\n${context.documentSummary}\n--------------------------------------\n` : ''}
PANDUAN MEMBUAT RINGKASAN (SUMMARY) & GRAFIK (CHART):
Ketika pengguna meminta membuat ringkasan/summary dan membuat chart (misal: "summary by kota, dan buat chartnya", "rekap per kategori dan buat grafiknya", dll.):
1. Data worksheet aktif sudah terlampir di atas. JANGAN HANYA MEMANGGIL \`read_sheet\` lalu berhenti!
2. Grafik Excel memerlukan tabel ringkasan data yang teragregasi agar dapat dirender dengan benar. Anda HARUS menghasilkan tool call sekaligus dalam satu respon:
   a. Tool \`write_cells\`: Hitung ringkasan agregasi per kategori (misal: jumlah kemunculan atau total nilai per Kota), lalu tulis tabel ringkasan tersebut ke range sel kosong di sebelah kanan tabel utama (misal jika data utama di kolom A-C dengan 140 baris, tempatkan tabel ringkasan di kolom E1:F5 atau sekitarnya lengkap dengan header dan barisnya).
   b. Tool \`format_range\`: Beri format header tebal (bold: true, fillColor: "#1E3A8A" atau "#F1F5F9") dan format angka yang sesuai (misal: "Rp#,##0" jika nilai uang).
   c. Tool \`auto_fit_columns\`: Sesuaikan lebar kolom agar tidak terpotong dan tidak memunculkan tanda '###'.
   d. Tool \`create_chart\`: Buat grafik dengan parameter \`dataRange\` mengarah tepat ke range tabel ringkasan yang baru Anda buat (misal: "E1:F5"), pilih \`chartType\` yang sesuai (Tren waktu -> "Line", Komparasi kategori -> "ColumnClustered" atau "BarClustered", Pangsa pasar -> "Pie"), dan berikan judul \`title\` (misal: "Ringkasan Penjualan per Kota").
3. Jika pengguna meminta "buatkan dashboard" atau "buatkan sheet baru untuk ringkasan", panggil \`add_worksheet\` terlebih dahulu (misal: sheetName: "Dashboard"), lalu tulis tabel dan buat grafik di sheet tersebut.
4. Jika pengguna meminta mengurutkan data (misal: "urutkan dari omset terbesar", "sort kota A-Z"), gunakan tool \`sort_and_filter\`.
5. JANGAN PERNAH berhenti di \`read_sheet\` ketika pengguna meminta membuat grafik/chart!

PENTING:
Jika pengguna meminta ringkasan tekstual saja ("summary sheet ini", "ringkas data ini", "analisis sheet ini", dll.), Anda DAPAT MEMBACA dan menganalisis data tabel di atas secara langsung! Jelaskan angka kunci, total, rata-rata, tren, dan kesimpulan secara komprehensif, cerdas, dan ramah dalam bahasa Indonesia.
PANDUAN PEMBELAJARAN MANDIRI (SELF-LEARNING SOUL.MD): Jika pengguna memberikan arahan gaya baru, preferensi format, koreksi istilah, atau direktif penulisan spesifik, proaktif panggil tool \`learn_corporate_directive\` untuk menyimpannya ke memori korporat.
PANDUAN REFERENSI BERKAS LOKAL (MINI-RAG): Jika pengguna menanyakan data, angka, atau pedoman dari berkas referensi yang dilampirkan, gunakan tool \`search_reference_knowledge\` untuk mencari kutipan relevan sebelum menulis data atau menyusun formula.`;
  }

  getTools(): ToolDefinition[] {
    return [
      {
        name: 'read_sheet',
        description: 'Membaca data tambahan dari worksheet jika range yang dibutuhkan belum tercakup pada data konteks di atas. JANGAN panggil jika data sudah tersedia di konteks aktif.',
        parameters: {
          type: 'object',
          properties: {
            range: { type: 'string', description: 'Range sel opsional, misal: "A1:D20". Jika kosong, membaca seluruh usedRange aktif.' },
          },
        },
      },
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
      {
        name: 'add_worksheet',
        description: 'Membuat tab lembar kerja baru (worksheet) di Excel (misal: "Dashboard", "Ringkasan", "Analisis") dan menjadikannya sheet aktif.',
        parameters: {
          type: 'object',
          properties: {
            sheetName: { type: 'string', description: 'Nama lembar kerja baru, misal: "Dashboard" atau "Ringkasan Penjualan"' },
          },
          required: ['sheetName'],
        },
      },
      {
        name: 'auto_fit_columns',
        description: 'Menyesuaikan lebar kolom secara otomatis (AutoFit) agar teks tidak terpotong dan angka tidak menjadi tanda pagar (###).',
        parameters: {
          type: 'object',
          properties: {
            range: { type: 'string', description: 'Range sel opsional, misal "A:G" atau "E1:F10". Jika kosong, merapikan seluruh usedRange.' },
          },
        },
      },
      {
        name: 'sort_and_filter',
        description: 'Mengurutkan (sort) baris data berdasarkan kolom tertentu (A-Z, Z-A, terkecil, terbesar) dan opsi mengaktifkan AutoFilter pada header.',
        parameters: {
          type: 'object',
          properties: {
            range: { type: 'string', description: 'Alamat range tabel data yang ingin diurutkan, misal: "A1:E100"' },
            columnIndex: { type: 'number', description: 'Indeks kolom pengurutan di dalam range (0 untuk kolom pertama range, 1 untuk kolom kedua, dst.)' },
            ascending: { type: 'boolean', description: 'True untuk urutan terkecil ke terbesar / A-Z, False untuk terbesar ke terkecil / Z-A. Default: true' },
            hasHeaders: { type: 'boolean', description: 'Apakah baris pertama merupakan judul kolom / header (default: true)' },
            enableAutoFilter: { type: 'boolean', description: 'Jika true, aktifkan tombol AutoFilter pada header tabel' },
          },
          required: ['range', 'columnIndex'],
        },
      },
      {
        name: 'clean_data',
        description: 'Membersihkan data pada worksheet: menghapus spasi ekstra (trim), menghapus baris duplikat, dan mengisi nilai sel kosong.',
        parameters: {
          type: 'object',
          properties: {
            range: { type: 'string', description: 'Range sel data yang ingin dibersihkan (opsional, default: seluruh usedRange aktif)' },
            removeDuplicates: { type: 'boolean', description: 'Jika true, hapus baris duplikat yang identik' },
            trimWhitespace: { type: 'boolean', description: 'Jika true, bersihkan spasi ekstra di awal dan akhir teks sel' },
            fillEmptyValues: { type: 'string', description: 'Nilai pengganti untuk sel yang kosong (misal: 0 atau "-")' },
          },
        },
      },
      {
        name: 'apply_conditional_formatting',
        description: 'Menerapkan pemformatan bersyarat (conditional formatting) pada rentang sel (skala warna, data bar, atau batas nilai/threshold).',
        parameters: {
          type: 'object',
          properties: {
            range: { type: 'string', description: 'Alamat range yang akan diformat, misal: "B2:B50"' },
            type: {
              type: 'string',
              enum: ['color_scale', 'data_bar', 'highlight_threshold'],
              description: 'Jenis aturan pemformatan bersyarat',
            },
            color: { type: 'string', description: 'Kode warna HEX opsional (misal: "#10B981" atau "#EF4444")' },
            thresholdValue: { type: 'number', description: 'Nilai ambang batas untuk highlight_threshold (misal: nilai > 1000)' },
          },
          required: ['range', 'type'],
        },
      },
      {
        name: 'audit_sheet_data',
        description: 'Mengaudit data lembar kerja secara mendalam untuk mendeteksi error formula (#REF!, #DIV/0!, dsb), rumus tidak konsisten, sel manual yang menimpa rumus, atau nilai anomali.',
        parameters: {
          type: 'object',
          properties: {
            range: { type: 'string', description: 'Range sel opsional, misal "A1:E50". Jika kosong, mengaudit seluruh usedRange lembar kerja aktif.' },
          },
        },
      },
      {
        name: 'generate_data_story',
        description: 'Menghasilkan ringkasan naratif eksekutif bisnis komprehensif, metrik kunci, tren pertumbuhan, dan rekomendasi berbasis data spreadsheet.',
        parameters: {
          type: 'object',
          properties: {
            range: { type: 'string', description: 'Range sel opsional, misal "A1:D20". Jika kosong, menganalisis seluruh data tabel pada worksheet aktif.' },
            focusMetric: { type: 'string', description: 'Nama kolom metrik atau angka yang ingin menjadi fokus utama analisis naratif (misal: "Revenue", "Laba Bersih", "Penjualan").' },
            includeRecommendations: { type: 'boolean', description: 'Apakah menyertakan poin rekomendasi bisnis strategis (default: true).' },
          },
        },
      },
      {
        name: 'share_to_cross_app_hub',
        description: 'Membagikan data tabel dan ringkasan eksekutif dari lembar kerja aktif ke Sam Universal Hub agar dapat langsung diimpor ke Word atau PowerPoint.',
        parameters: {
          type: 'object',
          properties: {
            title: { type: 'string', description: 'Judul snapshot data yang dibagikan (misal: "Rekap Penjualan Q3")' },
            range: { type: 'string', description: 'Range sel tabel opsional (misal: "A1:E20"). Jika kosong, membaca seluruh data aktif.' },
            summaryText: { type: 'string', description: 'Catatan ringkasan eksekutif opsional mengenai data ini.' },
          },
        },
      },
      {
        name: 'web_search',
        description: 'Mencari informasi, data terkini, statistik, kurs valuta asing, harga komoditas, atau fakta terbaru dari internet untuk disintesis dan dimasukkan ke dalam dokumen.',
        parameters: {
          type: 'object',
          properties: {
            query: { type: 'string', description: 'Kata kunci pencarian spesifik di internet' },
            maxResults: { type: 'number', description: 'Jumlah hasil maksimal (default: 5)' },
          },
          required: ['query'],
        },
      },
      {
        name: 'learn_corporate_directive',
        description: 'Mencatat dan mempelajari aturan gaya baru, preferensi format, koreksi istilah, atau direktif penulisan dari pengguna ke dalam memori persisten korporat (SOUL.md) agar selalu dipatuhi di masa mendatang.',
        parameters: {
          type: 'object',
          properties: {
            rule: { type: 'string', description: 'Aturan atau preferensi spesifik yang harus dipelajari' },
            category: {
              type: 'string',
              enum: ['tone', 'terminology', 'formatting', 'constraint', 'general'],
              description: 'Kategori direktif yang dipelajari.',
            },
            explanation: { type: 'string', description: 'Alasan atau konteks mengapa aturan ini dipelajari dari percakapan.' },
          },
          required: ['rule'],
        },
      },
      {
        name: 'search_reference_knowledge',
        description: 'Mencari fakta, klausul, data angka, atau pedoman spesifik dari berkas referensi lokal yang dilampirkan pengguna (SOP, pedoman, data CSV, catatan laporan).',
        parameters: {
          type: 'object',
          properties: {
            query: { type: 'string', description: 'Pertanyaan atau kata kunci topik yang ingin dicari di dalam berkas referensi' },
            maxResults: { type: 'number', description: 'Jumlah kutipan chunk relevan yang ingin diambil (default: 3)' },
          },
          required: ['query'],
        },
      },
      ...getLearnedAndMetaTools(this.hostType),
    ];
  }

  async executeTool(toolCall: ToolCall, _context: AgentContext): Promise<{ success: boolean; result?: any; error?: string }> {
    const metaCheck = await executeMetaOrCustomTool(this.hostType, toolCall, _context);
    if (metaCheck.handled) {
      return metaCheck.result!;
    }

    const driver = this.driverOverride || getOfficeDriver('Excel');
    try {
      if (toolCall.name === 'learn_corporate_directive') {
        const { rule, category, explanation } = toolCall.arguments || {};
        addLearnedDirective(rule, category || 'general', explanation, 'react_tool');
        return {
          success: true,
          result: `🧠 Berhasil mempelajari direktif korporat: "${rule}" (Kategori: ${category || 'general'}). Aturan ini telah dicatat ke dalam SOUL.md dan akan otomatis diterapkan pada seluruh dokumen mendatang.`,
        };
      }

      if (toolCall.name === 'read_sheet') {
        const { range } = toolCall.arguments || {};
        const data = (driver.readActiveSheetData ? await driver.readActiveSheetData(range) : await driver.readActiveRange());
        return {
          success: true,
          result: {
            sheetName: (data as any)?.sheetName || 'Sheet1',
            address: data.address,
            rowCount: (data as any)?.rowCount || data.values.length,
            columnCount: (data as any)?.columnCount || (data.values[0]?.length || 0),
            values: data.values,
          },
        };
      }

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

      if (toolCall.name === 'add_worksheet') {
        const { sheetName } = toolCall.arguments || {};
        if (!sheetName) return { success: false, error: 'Nama lembar kerja baru (sheetName) harus diisi.' };
        if (driver.addWorksheet) {
          const res = await driver.addWorksheet(sheetName);
          return { success: true, result: `Berhasil membuat lembar kerja baru "${res.sheetName}" dan mengaktifkannya.` };
        }
        return { success: false, error: 'Driver Excel tidak mendukung penambahan worksheet.' };
      }

      if (toolCall.name === 'auto_fit_columns') {
        const { range } = toolCall.arguments || {};
        if (driver.autoFitColumns) {
          await driver.autoFitColumns(range);
          return { success: true, result: `Lebar kolom berhasil disesuaikan secara otomatis (AutoFit).` };
        }
        return { success: false, error: 'Driver Excel tidak mendukung autoFitColumns.' };
      }

      if (toolCall.name === 'sort_and_filter') {
        const { range, columnIndex, ascending = true, hasHeaders = true, enableAutoFilter = false } = toolCall.arguments || {};
        if (!range || columnIndex === undefined) {
          return { success: false, error: 'Parameter range dan columnIndex wajib diisi.' };
        }
        if (driver.sortRange) {
          await driver.sortRange(range, Number(columnIndex), Boolean(ascending), Boolean(hasHeaders), Boolean(enableAutoFilter));
          return {
            success: true,
            result: `Data pada ${range} berhasil diurutkan berdasarkan kolom ke-${Number(columnIndex) + 1} (${ascending ? 'A-Z / Terkecil' : 'Z-A / Terbesar'}).`,
          };
        }
        return { success: false, error: 'Driver Excel tidak mendukung sortRange.' };
      }

      if (toolCall.name === 'clean_data') {
        const { range, removeDuplicates, trimWhitespace, fillEmptyValues } = toolCall.arguments || {};
        if (driver.cleanData) {
          const res = await driver.cleanData({ range, removeDuplicates, trimWhitespace, fillEmptyValues });
          return {
            success: true,
            result: `Pembersihan data selesai: ${res.cleanedRows} baris aktif, ${res.removedDuplicatesCount} baris duplikat dihapus, ${res.trimmedCellsCount} sel dirapikan, ${res.filledCellsCount} sel kosong diisi.`,
          };
        }
        return { success: false, error: 'Driver Excel tidak mendukung cleanData.' };
      }

      if (toolCall.name === 'apply_conditional_formatting') {
        const { range, type, color, thresholdValue } = toolCall.arguments || {};
        if (!range || !type) {
          return { success: false, error: 'Parameter range dan type wajib diisi untuk conditional formatting.' };
        }
        if (driver.applyConditionalFormatting) {
          const res = await driver.applyConditionalFormatting({ range, type, color, thresholdValue });
          return {
            success: true,
            result: `Pemformatan bersyarat (${res.rule}) berhasil diterapkan pada ${range}.`,
          };
        }
        return { success: false, error: 'Driver Excel tidak mendukung applyConditionalFormatting.' };
      }

      if (toolCall.name === 'audit_sheet_data') {
        const { range } = toolCall.arguments || {};
        if (driver.auditSheetData) {
          const audit = await driver.auditSheetData({ range });
          let report = `### 🔍 Laporan Audit Lembar Kerja: ${audit.sheetName}\n\n`;
          report += `**Ringkasan:** ${audit.summary}\n`;
          report += `- **Total Sel Diaudit:** ${audit.totalCellsAudited}\n`;
          report += `- **Total Error Formula:** ${audit.criticalIssues.length}\n`;
          report += `- **Total Peringatan/Inkonsistensi:** ${audit.warnings.length}\n\n`;

          if (audit.criticalIssues.length > 0) {
            report += `#### 🚨 Error Kritis Formula (${audit.criticalIssues.length})\n`;
            for (const issue of audit.criticalIssues) {
              const formulaInfo = issue.formula ? ` | Formula: \`${issue.formula}\`` : '';
              report += `- **Sel ${issue.address}**: Nilai \`${issue.currentValue}\`${formulaInfo}\n  - *Saran:* ${issue.suggestion}\n`;
            }
            report += '\n';
          }

          if (audit.warnings.length > 0) {
            report += `#### ⚠️ Peringatan & Anomali (${audit.warnings.length})\n`;
            for (const warn of audit.warnings) {
              const valInfo = warn.currentValue !== undefined ? ` (Nilai: \`${warn.currentValue}\`)` : '';
              report += `- **Sel ${warn.address}**${valInfo}: ${warn.suggestion}\n`;
            }
            report += '\n';
          }

          if (audit.totalErrorsFound === 0) {
            report += `✅ Lembar kerja berada dalam kondisi bersih tanpa error formula atau inkonsistensi yang terdeteksi.`;
          }

          return {
            success: true,
            result: report.trim(),
          };
        }
        return { success: false, error: 'Driver Excel tidak mendukung auditSheetData.' };
      }

      if (toolCall.name === 'generate_data_story') {
        const { range, focusMetric, includeRecommendations } = toolCall.arguments || {};
        if (driver.generateDataStory) {
          const story = await driver.generateDataStory({ range, focusMetric, includeRecommendations });
          let report = `### 📊 Narasi Eksekutif Data\n\n`;
          report += `**Headline:** ${story.headline}\n\n`;

          if (story.metrics && story.metrics.length > 0) {
            report += `#### 📈 Metrik Utama\n`;
            for (const m of story.metrics) {
              const trendIcon = m.trend === 'up' ? '▲' : m.trend === 'down' ? '▼' : '●';
              const change = m.changePercent !== undefined ? ` (${m.changePercent > 0 ? '+' : ''}${m.changePercent}%)` : '';
              report += `- **${m.label}**: ${m.value}${change} ${trendIcon}\n`;
            }
            report += '\n';
          }

          if (story.keyFindings && story.keyFindings.length > 0) {
            report += `#### 💡 Temuan Kunci\n`;
            for (const finding of story.keyFindings) {
              report += `- ${finding}\n`;
            }
            report += '\n';
          }

          if (story.risksOrAnomalies && story.risksOrAnomalies.length > 0) {
            report += `#### ⚠️ Risiko & Anomali Data\n`;
            for (const risk of story.risksOrAnomalies) {
              report += `- ${risk}\n`;
            }
            report += '\n';
          }

          if (story.recommendations && story.recommendations.length > 0) {
            report += `#### 🎯 Rekomendasi Strategis\n`;
            for (const rec of story.recommendations) {
              report += `- ${rec}\n`;
            }
          }

          return {
            success: true,
            result: report.trim(),
          };
        }
        return { success: false, error: 'Driver Excel tidak mendukung generateDataStory.' };
      }

      if (toolCall.name === 'share_to_cross_app_hub') {
        const { title, range, summaryText } = toolCall.arguments || {};
        const data = driver.readActiveSheetData ? await driver.readActiveSheetData(range) : await driver.readActiveRange();
        const values: any[][] = data?.values || [];

        if (!values || values.length === 0) {
          return { success: false, error: 'Tidak ada data pada lembar kerja untuk dibagikan ke Universal Hub.' };
        }

        const headers = values[0].map((h: any) => String(h ?? ''));
        const rows = values.length > 1 ? values.slice(1) : [];
        const resolvedTitle = title || data?.sheetName || 'Data Lembar Kerja Excel';

        let resolvedSummary = summaryText;
        if (!resolvedSummary) {
          resolvedSummary = `Tabel "${resolvedTitle}" memuat ${rows.length} baris data dan ${headers.length} kolom (${headers.slice(0, 5).join(', ')}${headers.length > 5 ? '...' : ''}).`;
        }

        const snapshot = saveCrossAppSnapshot({
          sourceHost: 'Excel',
          title: resolvedTitle,
          artifactType: 'table_data',
          tableData: {
            headers,
            rows,
            totalRows: rows.length,
          },
          summaryText: resolvedSummary,
        });

        return {
          success: true,
          result: `✅ Berhasil membagikan data "${resolvedTitle}" ke Sam Universal Hub (ID: ${snapshot.id}). Data siap diimpor di Microsoft Word atau PowerPoint.`,
        };
      }

      if (toolCall.name === 'web_search') {
        const { query, maxResults } = toolCall.arguments || {};
        if (!query) {
          return { success: false, error: 'Query pencarian web tidak boleh kosong.' };
        }
        const searchRes = await searchWeb(query, { maxResults });
        let text = `Hasil pencarian web untuk "${query}":\n\n`;
        if (searchRes.results.length === 0) {
          text += 'Tidak ditemukan hasil yang relevan di internet.';
        } else {
          searchRes.results.forEach((item, index) => {
            text += `${index + 1}. **${item.title}**\n${item.snippet}\nTautan: ${item.url}\n\n`;
          });
          if (searchRes.citationsFormatted) {
            text += `\n${searchRes.citationsFormatted}`;
          }
        }
        return { success: true, result: text.trim() };
      }

      if (toolCall.name === 'search_reference_knowledge') {
        const { query, maxResults = 3 } = toolCall.arguments || {};
        if (!query) {
          return { success: false, error: 'Query pencarian referensi wajib diisi.' };
        }
        const results = queryKnowledge(query, { maxResults: Number(maxResults) || 3 });
        if (results.length === 0) {
          return {
            success: true,
            result: `Tidak ditemukan kutipan yang relevan dari berkas referensi untuk kata kunci: "${query}".`,
          };
        }
        let formattedText = `📖 **Hasil Temuan Berkas Referensi:**\n\n`;
        results.forEach((r, idx) => {
          formattedText += `[${idx + 1}] **Sumber ${r.sourceCitation}** (Skor relevansi: ${r.score.toFixed(2)}):\n> "${r.snippet}"\n\n`;
        });
        return { success: true, result: formattedText.trim() };
      }

      return { success: false, error: `Tool ${toolCall.name} tidak dikenali.` };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  }
}
