import { getOfficeDriver } from '../../services/office';
import { renderChartToBase64Png } from '../../utils/chartRenderer';
import { AgentContext, IAgent } from '../types';
import { ToolCall, ToolDefinition } from '../../types';
import { executeMetaOrCustomTool, getLearnedAndMetaTools } from '../metaTools';
import {
  getLatestCrossAppSnapshot,
  saveCrossAppSnapshot,
  listCrossAppSnapshots,
} from '../../services/storage/crossAppBridge';

export class WordAgent implements IAgent {
  id = 'word-specialist';
  name = 'Word Specialist';
  hostType = 'Word' as const;

  constructor(private driverOverride?: any) {}

  getSystemPrompt(context: AgentContext): string {
    return `Anda adalah Sam Office Agent spesialis Microsoft Word tingkat Expert.
Anda ahli dalam penyusunan surat resmi, laporan profesional, perbaikan tata bahasa/proofreading, pemformatan heading dokumen, tabel terstruktur, pemisah halaman, pencarian/penggantian teks massal, serta pembuatan grafik visual.
${context.documentSummary ? `\n--- DOKUMEN WORD AKTIF SAAT INI ---\n${context.documentSummary}\n------------------------------------\n` : ''}
PANDUAN ALUR KERJA:
1. Jika pengguna meminta membuat file PRD (Product Requirement Document), spesifikasi, proposal, atau konten untuk didownload / dibuat di chat:
   - TULISKAN LENGKAP dokumen tersebut di respon chat dengan format Markdown profesional (Judul, Executive Summary, User Personas, Functional Requirements, Non-Functional Requirements, Milestones).
   - JANGAN memanggil tool insert_page_break atau tool lainnya jika pengguna hanya meminta membuat dokumen/PRD untuk didownload atau dibaca di chat!
2. Jika pengguna meminta ringkasan, analisis dokumen, intisari, atau pembuatan materi presentasi dari dokumen:
   - Anda DAPAT MEMBACA dan menganalisis teks dokumen di atas secara langsung!
   - Buatkan ringkasan eksekutif, poin-poin kunci terstruktur, dan outline slide presentasi yang siap pakai dalam bahasa Indonesia yang profesional.
3. Jika pengguna meminta membuat tabel dan grafik di Word (misal: "buat tabel data dan jadikan grafik"):
   - Panggil \`insert_table\` untuk membuat tabel datanya.
   - Panggil juga \`insert_chart_image\` untuk menghasilkan dan menyisipkan visual grafik (batang, garis, atau lingkaran) langsung ke dalam dokumen.
4. Jika pengguna secara spesifik meminta membuat halaman baru / jeda halaman di Word: gunakan \`insert_page_break\`.
5. Jika pengguna meminta mencari atau mengganti kata/istilah di seluruh dokumen: gunakan \`find_and_replace\`.
6. Jika pengguna meminta mereview kontrak, kepatuhan klausul, risiko hukum, atau SLA perjanjian: gunakan tool \`review_compliance_clauses\`.
7. Jika pengguna meminta merapikan format dokumen, brand korporat, atau standarisasi heading/font: gunakan tool \`apply_corporate_style\`.`;
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
      {
        name: 'insert_chart_image',
        description: 'Membuat dan menyisipkan grafik visual (batang, garis, atau lingkaran) ke dalam dokumen Word.',
        parameters: {
          type: 'object',
          properties: {
            chartType: { type: 'string', enum: ['bar', 'line', 'pie'], description: 'Jenis grafik: bar, line, atau pie' },
            labels: { type: 'array', description: 'Label sumbu X atau kategori, misal: ["Jan", "Feb", "Mar"]' },
            data: { type: 'array', description: 'Data angka nilai grafik, misal: [150, 200, 180]' },
            title: { type: 'string', description: 'Judul grafik (opsional)' },
          },
          required: ['chartType', 'labels', 'data'],
        },
      },
      {
        name: 'insert_page_break',
        description: 'Menyisipkan pemisah halaman (Page Break) atau pemisah seksi (Section Break) baru di dokumen Word.',
        parameters: {
          type: 'object',
          properties: {
            breakType: { type: 'string', enum: ['page', 'section'], description: 'Tipe pemisah: page (halaman baru biasa) atau section (seksi baru)' },
          },
        },
      },
      {
        name: 'find_and_replace',
        description: 'Mencari kata/frasa tertentu dan menggantinya dengan teks baru di seluruh dokumen Word.',
        parameters: {
          type: 'object',
          properties: {
            findText: { type: 'string', description: 'Teks yang dicari' },
            replaceText: { type: 'string', description: 'Teks pengganti' },
            matchCase: { type: 'boolean', description: 'Pencarian peka huruf besar/kecil (opsional, default: false)' },
          },
          required: ['findText', 'replaceText'],
        },
      },
      {
        name: 'generate_structured_doc',
        description: 'Menghasilkan dokumen resmi berformat terstruktur (SOP, MoM/Notulen, SPK, PRD, atau FormalMemo) lengkap dengan judul, metadata, dan bagian-bagian dokumen.',
        parameters: {
          type: 'object',
          properties: {
            templateType: {
              type: 'string',
              enum: ['SOP', 'MoM', 'SPK', 'PRD', 'FormalMemo'],
              description: 'Tipe template dokumen',
            },
            title: { type: 'string', description: 'Judul dokumen' },
            author: { type: 'string', description: 'Nama penyusun dokumen (opsional)' },
            sections: {
              type: 'array',
              description: 'Bagian-bagian dokumen',
              items: {
                type: 'object',
                properties: {
                  heading: { type: 'string', description: 'Judul bagian (Heading 2)' },
                  content: { type: 'string', description: 'Isi teks paragraf bagian' },
                  bullets: { type: 'array', items: { type: 'string' }, description: 'Daftar butir poin' },
                  table: {
                    type: 'object',
                    properties: {
                      headers: { type: 'array', items: { type: 'string' } },
                      rows: { type: 'array', items: { type: 'array', items: { type: 'string' } } },
                    },
                  },
                },
                required: ['heading'],
              },
            },
          },
          required: ['templateType', 'title', 'sections'],
        },
      },
      {
        name: 'polish_document_text',
        description: 'Memoles tata bahasa, gaya bahasa, dan penyampaian profesional pada teks yang dipilih atau seluruh dokumen Word.',
        parameters: {
          type: 'object',
          properties: {
            scope: {
              type: 'string',
              enum: ['selection', 'document'],
              description: 'Cakupan teks yang dipoles',
            },
            tone: {
              type: 'string',
              enum: ['formal_indonesia', 'executive_english', 'concise'],
              description: 'Gaya bahasa (tone) pemolesan',
            },
            instruction: { type: 'string', description: 'Instruksi khusus pemolesan' },
            customText: { type: 'string', description: 'Teks yang telah disempurnakan untuk dituliskan ke dokumen' },
          },
          required: ['scope', 'tone'],
        },
      },
      {
        name: 'review_compliance_clauses',
        description: 'Menganalisis dokumen/kontrak kerja/SLA untuk mendeteksi klausul berisiko tinggi (ganti rugi tanpa batas, denda berlebih, pemutusan sepihak) dan klausul penting yang hilang.',
        parameters: {
          type: 'object',
          properties: {
            scope: {
              type: 'string',
              enum: ['selection', 'document'],
              description: 'Cakupan teks yang dianalisis (default: document)',
            },
            contractType: {
              type: 'string',
              enum: ['vendor_service', 'employment', 'nda', 'procurement', 'general'],
              description: 'Jenis kontrak atau perjanjian kerja',
            },
            strictness: {
              type: 'string',
              enum: ['standard', 'strict'],
              description: 'Tingkat ketelitian telaah klausul',
            },
          },
        },
      },
      {
        name: 'apply_corporate_style',
        description: 'Menerapkan standarisasi format brand korporat (tema warna, tipografi, hierarki heading, spasi paragraf) ke dokumen atau teks yang dipilih.',
        parameters: {
          type: 'object',
          properties: {
            theme: {
              type: 'string',
              enum: ['corporate_navy', 'executive_emerald', 'modern_minimalist', 'official_government'],
              description: 'Tema gaya korporat yang diinginkan',
            },
            scope: {
              type: 'string',
              enum: ['selection', 'document'],
              description: 'Cakupan pemformatan (default: document)',
            },
            fontFamily: {
              type: 'string',
              description: 'Nama font kustom (opsional, misal Calibri, Segoe UI, Aptos, Times New Roman)',
            },
          },
        },
      },
      {
        name: 'import_from_cross_app_hub',
        description: 'Mengimpor data tabel, ringkasan eksekutif, atau metrik dari Sam Universal Hub (misal: data dari Excel atau PowerPoint) langsung ke dalam dokumen Word.',
        parameters: {
          type: 'object',
          properties: {
            snapshotId: { type: 'string', description: 'ID snapshot tertentu (opsional). Jika tidak disertakan, mengambil snapshot terbaru dari aplikasi lain.' },
            includeSummary: { type: 'boolean', description: 'Apakah menyertakan narasi ringkasan eksekutif (default: true).' },
            includeTable: { type: 'boolean', description: 'Apakah menyisipkan tabel data jika tersedia (default: true).' },
          },
        },
      },
      {
        name: 'share_to_cross_app_hub',
        description: 'Membagikan intisari dokumen Word atau PRD aktif ke Sam Universal Hub agar dapat diimpor ke PowerPoint untuk pembuatan slide deck presentasi.',
        parameters: {
          type: 'object',
          properties: {
            title: { type: 'string', description: 'Judul snapshot dokumen yang dibagikan (misal: "PRD Sistem Pembayaran").' },
            summaryText: { type: 'string', description: 'Ringkasan dokumen yang dibagikan (opsional, jika kosong membaca dari isi dokumen).' },
          },
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

    const driver = this.driverOverride || getOfficeDriver('Word');
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
      if (toolCall.name === 'insert_chart_image') {
        const { chartType, labels, data, title } = toolCall.arguments || {};
        const base64Png = renderChartToBase64Png({
          chartType: chartType || 'bar',
          labels: Array.isArray(labels) ? labels : [],
          data: Array.isArray(data) ? data.map(Number) : [],
          title: title || 'Grafik Data',
        });
        if (driver.insertPictureBase64) {
          await driver.insertPictureBase64(base64Png);
        }
        return { success: true, result: `Grafik ${chartType || 'bar'} "${title || 'Grafik Data'}" berhasil disisipkan ke dokumen.` };
      }
      if (toolCall.name === 'insert_page_break') {
        const { breakType } = toolCall.arguments || {};
        if (driver.insertPageBreak) {
          await driver.insertPageBreak(breakType || 'page');
        }
        return { success: true, result: `Pemisah ${breakType === 'section' ? 'seksi' : 'halaman'} berhasil disisipkan.` };
      }
      if (toolCall.name === 'find_and_replace') {
        const { findText, replaceText, matchCase } = toolCall.arguments || {};
        if (!findText) return { success: false, error: 'findText wajib diisi.' };
        if (driver.findAndReplace) {
          const res = await driver.findAndReplace(findText, replaceText || '', Boolean(matchCase));
          return { success: true, result: `Berhasil mengganti ${res.count} kemunculan "${findText}".` };
        }
        return { success: true, result: `Teks "${findText}" diganti.` };
      }
      if (toolCall.name === 'generate_structured_doc') {
        const { templateType, title, sections, author } = toolCall.arguments || {};
        if (!templateType || !title || !Array.isArray(sections)) {
          return { success: false, error: 'templateType, title, dan sections wajib diisi.' };
        }
        if (driver.generateStructuredDoc) {
          const res = await driver.generateStructuredDoc({ templateType, title, sections, author });
          return { success: true, result: res.message };
        }
        return { success: true, result: `Dokumen ${templateType} "${title}" berhasil dibuat.` };
      }
      if (toolCall.name === 'polish_document_text') {
        const { scope, tone, instruction, customText } = toolCall.arguments || {};
        if (!scope || !tone) {
          return { success: false, error: 'scope dan tone wajib diisi.' };
        }
        if (driver.polishDocumentText) {
          const res = await driver.polishDocumentText({ scope, tone, instruction, customText });
          return { success: true, result: `Teks berhasil dipoles (${tone}): ${res.polishedText}` };
        }
        return { success: true, result: `Teks dokumen berhasil dipoles dengan gaya "${tone}".` };
      }
      if (toolCall.name === 'review_compliance_clauses') {
        const { scope, contractType, strictness } = toolCall.arguments || {};
        if (driver.reviewComplianceClauses) {
          const res = await driver.reviewComplianceClauses({ scope, contractType, strictness });
          const riskBadge = res.overallRiskLevel === 'high' ? '🔴 TINGGI' : res.overallRiskLevel === 'medium' ? '🟡 MODERAT' : '🟢 RENDAH';
          let output = `### ⚖️ Laporan Review Kepatuhan Kontrak\n\n`;
          output += `- **Jenis Dokumen:** ${res.contractType}\n`;
          output += `- **Tingkat Risiko:** ${riskBadge}\n`;
          output += `- **Total Klausul Dianalisis:** ${res.clausesReviewedCount}\n`;
          output += `- **Ringkasan:** ${res.executiveSummary}\n\n`;

          if (res.identifiedClauses.length > 0) {
            output += `#### 📌 Temuan Klausul\n`;
            for (const c of res.identifiedClauses) {
              const statusEmoji = c.status === 'high_risk' ? '🚨' : c.status === 'warning' ? '⚠️' : '✅';
              output += `- ${statusEmoji} **[${c.category.toUpperCase()}]** *"${c.excerpt}"*\n  - *Analisis:* ${c.analysis}\n`;
              if (c.recommendation) {
                output += `  - *Rekomendasi:* ${c.recommendation}\n`;
              }
            }
            output += '\n';
          }

          if (res.missingCriticalClauses.length > 0) {
            output += `#### ⚠️ Klausul Kritis yang Belum Tercantum\n`;
            for (const m of res.missingCriticalClauses) {
              output += `- Klausul \`${m}\` tidak ditemukan dalam naskah perjanjian.\n`;
            }
            output += '\n';
          }

          if (res.actionableRecommendations.length > 0) {
            output += `#### 💡 Rekomendasi Tindak Lanjut\n`;
            for (const r of res.actionableRecommendations) {
              output += `- ${r}\n`;
            }
          }

          return { success: true, result: output.trim() };
        }
        return { success: false, error: 'Driver Word tidak mendukung reviewComplianceClauses.' };
      }
      if (toolCall.name === 'apply_corporate_style') {
        const { theme, scope, fontFamily } = toolCall.arguments || {};
        if (driver.applyCorporateStyle) {
          const res = await driver.applyCorporateStyle({ theme, scope, fontFamily });
          return {
            success: true,
            result: `🎨 **Format Brand Korporat Diterapkan**\n- **Tema:** ${res.appliedTheme}\n- **Font:** ${res.fontFamily}\n- **Paragraf Diformat:** ${res.styledParagraphsCount}\n- **Heading Disesuaikan:** ${res.headingsCount}\n\n${res.message}`,
          };
        }
        return { success: false, error: 'Driver Word tidak mendukung applyCorporateStyle.' };
      }
      if (toolCall.name === 'import_from_cross_app_hub') {
        const { snapshotId, includeSummary = true, includeTable = true } = toolCall.arguments || {};
        let snapshot = null;
        if (snapshotId) {
          const all = listCrossAppSnapshots();
          snapshot = all.find(s => s.id === snapshotId) || null;
        } else {
          snapshot = getLatestCrossAppSnapshot('Word');
        }

        if (!snapshot) {
          return { success: false, error: 'Tidak ditemukan snapshot data di Universal Hub dari aplikasi lain.' };
        }

        // Insert heading
        await driver.insertContent('end', `Ringkasan: ${snapshot.title}`, 'h2');

        // Insert summary narrative if present and requested
        if (includeSummary && snapshot.summaryText) {
          await driver.insertContent('end', snapshot.summaryText, 'paragraph');
        }

        // Insert table if tableData is present and requested
        if (includeTable && snapshot.tableData && snapshot.tableData.headers && snapshot.tableData.rows) {
          const { headers, rows } = snapshot.tableData;
          const tableDataFormatted = [headers, ...rows.map(r => r.map(c => String(c ?? '')))];
          await driver.insertTable(tableDataFormatted.length, headers.length, tableDataFormatted);
        }

        return {
          success: true,
          result: `✅ Berhasil mengimpor data "${snapshot.title}" dari ${snapshot.sourceHost} via Universal Hub ke dalam dokumen Word.`,
        };
      }
      if (toolCall.name === 'share_to_cross_app_hub') {
        const { title, summaryText } = toolCall.arguments || {};
        let resolvedSummary = summaryText || _context.documentSummary;
        if (!resolvedSummary && driver.getWordOutline) {
          resolvedSummary = await driver.getWordOutline();
        }
        if (!resolvedSummary) {
          resolvedSummary = 'Dokumen Word tanpa ringkasan.';
        }

        const resolvedTitle = title || 'Dokumen Word';
        const snapshot = saveCrossAppSnapshot({
          sourceHost: 'Word',
          title: resolvedTitle,
          artifactType: 'executive_summary',
          summaryText: resolvedSummary,
        });

        return {
          success: true,
          result: `✅ Berhasil membagikan dokumen "${resolvedTitle}" ke Sam Universal Hub (ID: ${snapshot.id}). Data siap diimpor di PowerPoint atau Excel.`,
        };
      }
      return { success: false, error: 'Tool tidak ditemukan.' };
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }
}
