import { getOfficeDriver } from '../../services/office';
import { AgentContext, IAgent } from '../types';
import { ToolCall, ToolDefinition } from '../../types';
import { executeMetaOrCustomTool, getLearnedAndMetaTools } from '../metaTools';
import {
  getLatestCrossAppSnapshot,
  listCrossAppSnapshots,
} from '../../services/storage/crossAppBridge';
import { searchWeb } from '../../services/search';
import { addLearnedDirective } from '../../services/storage/soulStorage';
import { queryKnowledge } from '../../services/rag/ragEngine';

export class PPTAgent implements IAgent {
  id = 'ppt-specialist';
  name = 'PowerPoint Specialist';
  hostType = 'PowerPoint' as const;

  constructor(private driverOverride?: any) {}

  getSystemPrompt(context: AgentContext): string {
    const summaryHeader = context.documentSummary
      ? '\n--- SLIDE / MATERI SAAT INI ---\n' + context.documentSummary + '\n----------------------------\n'
      : '';

    return 'Anda adalah Sam Office Agent spesialis Microsoft PowerPoint tingkat Expert.\n' +
      'Anda ahli dalam merancang struktur presentasi yang memikat, menyusun poin slide yang ringkas dan padat, membuat slide baru secara instan, dan membuat speaker notes untuk pemateri.\n' +
      summaryHeader +
      '\nPANDUAN ALUR KERJA:\n' +
      '1. MEMBACA & MERINGKAS SLIDE PRESENTASI:\n' +
      '   - Anda MEMILIKI AKSES PENUH ke teks materi pada slide aktif dan seluruh slide dalam presentasi yang tertera di bagian "--- SLIDE / MATERI SAAT INI ---" di atas.\n' +
      '   - Selain itu, Anda memiliki tool "read_slides" untuk membaca isi slide aktif maupun seluruh slide secara langsung dari PowerPoint.\n' +
      '   - Jika pengguna meminta "baca slide ini buat summary", "ringkas slide aktif", atau menganalisis materi slide:\n' +
      '     * Jika konteks slide di atas sudah memiliki isi materi yang jelas, Anda bisa langsung menyajikan ringkasannya secara komprehensif, padat, dan profesional di chat.\n' +
      '     * Jika teks materi di konteks belum terbaca lengkap atau pengguna meminta membaca ulang/slide lain, PANGGIL tool "read_slides" untuk mengambil isi slide terbaru dari PowerPoint!\n' +
      '     * JANGAN PERNAH berasumsi slide kosong tanpa memanggil "read_slides" jika pengguna menyatakan ada materi di slidenya!\n' +
      '2. MEMBUAT SLIDE PRESENTASI:\n' +
      '   - Jika pengguna meminta membuat slide presentasi (misal: "buatkan slide tentang...", "buat 3 slide presentasi", "jadikan materi presentasi dari chat"):\n' +
      '     * Gunakan tool create_presentation_deck jika pengguna meminta membuat beberapa slide sekaligus.\n' +
      '     * Gunakan tool create_slide jika pengguna meminta membuat 1 slide baru.\n' +
      '     * Pastikan setiap slide memiliki judul yang kuat (title) dan 3-5 butir materi (bullets) yang ringkas, profesional, dan to-the-point. Sertakan notes (catatan pemateri) yang bermanfaat.\n' +
      '3. PANDUAN SPEAKER NOTES (CATATAN PEMATERI):\n' +
      '   - BATASAN TEKNIS OFFICE.JS: Microsoft PowerPoint JavaScript API (Office.js) secara resmi TIDAK menyediakan API untuk menulis langsung ke panel Notes bawaan PowerPoint (properti slide.notesPage tidak ada/tidak diekspos oleh Microsoft).\n' +
      '   - JANGAN PERNAH membuat dynamic office script yang memanggil "slide.notesPage" karena akan error runtime.\n' +
      '   - Jika pengguna meminta speaker notes untuk satu atau seluruh slide: Sajikan naskah speaker notes yang terstruktur, elegan, dan siap pakai langsung di pesan chat dengan format per slide (Intonasi, Naskah Penyampaian, Poin Nilai Juri, Transisi) agar pemateri dapat langsung menyalinnya ke panel Notes PowerPoint atau menjadikannya lembar panduan saat presentasi.\n' +
      '4. TRANSFORMASI DOKUMEN KE SLIDE PRESENTASI (DOC-TO-DECK):\n' +
      '   - Gunakan tool "transform_doc_to_deck" untuk mengubah teks dokumen, proposal, memo, atau laporan menjadi rangkaian slide presentasi eksekutif terstruktur (Executive Storyline Arc).\n' +
      '   - Tool ini secara komprehensif menghasilkan naskah presenter (hook, poin elaborasi, transisi) untuk setiap slide.\n' +
      '   - Jika documentText dikosongkan, tool akan otomatis mengambil teks dari dokumen aktif saat ini.\n' +
      '5. VISUALISASI DIAGRAM ALUR PROSES (FLOWCHART):\n' +
      '   - Ketika pengguna meminta dibuatkan diagram alur, flowchart, SOP, atau visualisasi alur kerja di PowerPoint, gunakan tool "insert_process_flowchart" untuk menghasilkan diagram visual profesional dan langsung menyisipkannya ke slide presentasi.\n' +
      '6. Berikan penjelasan ringkas, terstruktur, dan ramah di chat mengenai topik slide yang telah Anda ringkas atau buatkan ke dalam presentasi.\n' +
      '7. PANDUAN PEMBELAJARAN MANDIRI (SELF-LEARNING SOUL.MD): Jika pengguna memberikan arahan gaya baru, tema warna, preferensi tata letak/format slide, atau direktif presentasi spesifik, proaktif panggil tool "learn_corporate_directive" untuk menyimpannya ke memori korporat.\n' +
      '8. PANDUAN REFERENSI BERKAS LOKAL (MINI-RAG): Jika pengguna menanyakan materi, fakta, atau pedoman dari berkas referensi yang dilampirkan, gunakan tool "search_reference_knowledge" untuk mencari kutipan relevan sebelum menyusun slide presentasi.';
  }

  getTools(): ToolDefinition[] {
    return [
      {
        name: 'read_slides',
        description: 'Membaca teks, judul, poin-poin materi, dan isi slide presentasi yang sedang dibuka di PowerPoint.',
        parameters: {
          type: 'object',
          properties: {
            slideNumber: {
              type: 'number',
              description: 'Nomor slide spesifik yang ingin dibaca (opsional, contoh: 1). Jika diabaikan, akan membaca seluruh slide atau slide aktif.',
            },
            allSlides: {
              type: 'boolean',
              description: 'Set true untuk membaca seluruh slide yang ada di presentasi.',
            },
          },
        },
      },
      {
        name: 'create_slide',
        description: 'Membuat slide baru lengkap dengan judul, poin-poin materi, dan catatan pemateri (opsional) di PowerPoint.',
        parameters: {
          type: 'object',
          properties: {
            title: { type: 'string', description: 'Judul slide' },
            bullets: { type: 'array', items: { type: 'string' }, description: 'Daftar butir poin materi slide' },
            notes: { type: 'string', description: 'Catatan pemateri / speaker notes (opsional)' },
            layout: { type: 'string', enum: ['TitleOnly', 'TitleAndContent', 'TwoContent'], description: 'Layout slide (opsional)' },
          },
          required: ['title', 'bullets'],
        },
      },
      {
        name: 'create_presentation_deck',
        description: 'Membuat beberapa slide presentasi sekaligus secara berurutan ke dalam PowerPoint.',
        parameters: {
          type: 'object',
          properties: {
            slides: {
              type: 'array',
              items: {
                type: 'object',
                properties: {
                  title: { type: 'string', description: 'Judul slide' },
                  bullets: { type: 'array', items: { type: 'string' }, description: 'Poin-poin materi' },
                  notes: { type: 'string', description: 'Catatan pemateri (opsional)' },
                  layout: { type: 'string', enum: ['TitleOnly', 'TitleAndContent', 'TwoContent'], description: 'Layout slide (opsional)' },
                },
                required: ['title', 'bullets'],
              },
              description: 'Daftar slide yang akan dibuat',
            },
          },
          required: ['slides'],
        },
      },
      {
        name: 'add_slide',
        description: 'Menambahkan slide baru kosong ke presentasi.',
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
        description: 'Mengisi judul slide dan poin-poin materi pada slide aktif.',
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
      {
        name: 'generate_themed_deck',
        description: 'Membuat presentasi lengkap dengan tema desain visual profesional (corporate_blue, emerald_executive, modern_dark, minimalist_clean) dan variasi layout slide.',
        parameters: {
          type: 'object',
          properties: {
            topic: { type: 'string', description: 'Topik utama presentasi' },
            theme: {
              type: 'string',
              enum: ['corporate_blue', 'emerald_executive', 'modern_dark', 'minimalist_clean'],
              description: 'Tema warna visual presentasi',
            },
            slides: {
              type: 'array',
              description: 'Daftar slide berformat tema dan layout terstruktur',
              items: {
                type: 'object',
                properties: {
                  title: { type: 'string', description: 'Judul slide' },
                  layout: {
                    type: 'string',
                    enum: ['title_cover', 'split_comparison', 'bullet_points', 'metric_highlights'],
                    description: 'Layout slide',
                  },
                  content: { type: 'array', items: { type: 'string' }, description: 'Poin-poin isi materi slide' },
                  metrics: {
                    type: 'array',
                    items: {
                      type: 'object',
                      properties: {
                        label: { type: 'string' },
                        value: { type: 'string' },
                      },
                      required: ['label', 'value'],
                    },
                    description: 'Data metrik untuk layout metric_highlights',
                  },
                  notes: { type: 'string', description: 'Catatan pemateri' },
                },
                required: ['title'],
              },
            },
          },
          required: ['topic', 'slides'],
        },
      },
      {
        name: 'transform_doc_to_deck',
        description: 'Mengubah teks dokumen, ringkasan laporan, atau data bisnis menjadi rangkaian slide presentasi eksekutif terstruktur (Executive Storyline Arc) lengkap dengan naskah pemateri (speaker notes) per slide.',
        parameters: {
          type: 'object',
          properties: {
            documentText: { type: 'string', description: 'Teks dokumen mentah yang ingin diubah menjadi slide presentasi (opsional, jika kosong membaca konteks aktif).' },
            targetSlideCount: { type: 'number', description: 'Target jumlah slide yang dihasilkan (default: 5).' },
            theme: {
              type: 'string',
              enum: ['corporate_blue', 'emerald_executive', 'modern_dark', 'minimalist_clean'],
              description: 'Tema desain visual presentasi korporat.',
            },
            presentationTitle: { type: 'string', description: 'Judul presentasi kustom (opsional).' },
            targetAudience: { type: 'string', enum: ['executive', 'technical', 'team_all_hands'], description: 'Target audiens presentasi (opsional).' },
          },
        },
      },
      {
        name: 'import_from_cross_app_hub',
        description: 'Mengimpor data snapshot dari Sam Universal Hub (tabel Excel, ringkasan Word, PRD) dan secara otomatis mentransformasikannya menjadi rangkaian slide presentasi eksekutif lengkap dengan speaker notes.',
        parameters: {
          type: 'object',
          properties: {
            snapshotId: { type: 'string', description: 'ID snapshot tertentu (opsional). Jika tidak disertakan, mengambil snapshot terbaru dari aplikasi lain.' },
            theme: {
              type: 'string',
              enum: ['corporate_blue', 'emerald_executive', 'modern_dark', 'minimalist_clean'],
              description: 'Tema desain visual presentasi.',
            },
            targetSlideCount: { type: 'number', description: 'Target jumlah slide yang dihasilkan (default: 5).' },
            targetAudience: { type: 'string', enum: ['executive', 'technical', 'team_all_hands'], description: 'Target audiens presentasi.' },
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
        name: 'insert_process_flowchart',
        description: 'Membuat diagram alur proses bisnis/SOP (flowchart visual profesional) dari teks alur kerja dan langsung menyisipkannya sebagai slide visual khusus di PowerPoint.',
        parameters: {
          type: 'object',
          properties: {
            textOrSteps: {
              type: 'string',
              description: 'Deskripsi alur kerja atau tahapan SOP.',
            },
            title: {
              type: 'string',
              description: 'Judul slide diagram alur.',
            },
            theme: {
              type: 'string',
              enum: ['corporate_navy', 'emerald_executive', 'modern_dark', 'amber_warm'],
              description: 'Tema warna korporat presentasi.',
            },
            direction: {
              type: 'string',
              enum: ['TD', 'LR'],
              description: 'Arah diagram alur.',
            },
          },
          required: ['textOrSteps'],
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

  async executeTool(toolCall: ToolCall, context: AgentContext): Promise<{ success: boolean; result?: any; error?: string }> {
    const metaCheck = await executeMetaOrCustomTool(this.hostType, toolCall, context);
    if (metaCheck.handled) {
      return metaCheck.result!;
    }

    const driver = this.driverOverride || getOfficeDriver('PowerPoint');
    try {
      if (toolCall.name === 'learn_corporate_directive') {
        const { rule, category, explanation } = toolCall.arguments || {};
        addLearnedDirective(rule, category || 'general', explanation, 'react_tool');
        return {
          success: true,
          result: `🧠 Berhasil mempelajari direktif korporat: "${rule}" (Kategori: ${category || 'general'}). Aturan ini telah dicatat ke dalam SOUL.md dan akan otomatis diterapkan pada seluruh dokumen mendatang.`,
        };
      }

      if (toolCall.name === 'read_slides') {
        const { slideNumber, allSlides } = toolCall.arguments || {};
        if (driver.readSlideData) {
          const res = await driver.readSlideData(slideNumber, allSlides ?? (slideNumber === undefined));
          const slideLines = res.slides
            .map((s: any) => {
              let line = `[Slide #${s.slideIndex}]: "${s.title}"\n${s.textContent}`;
              if (s.notes) line += `\nSpeaker Notes: ${s.notes}`;
              return line;
            })
            .join('\n\n');

          return {
            success: true,
            result: `📋 Hasil Pembacaan Slide PowerPoint:\nTotal Slide: ${res.totalSlides} | Slide Aktif: #${res.activeSlideIndex}\n\n${slideLines}`,
          };
        }

        const ctx = await driver.getSlideContext();
        return {
          success: true,
          result: `📋 Hasil Pembacaan Slide #${ctx.slideNumber}:\nJudul: "${ctx.title}"\nIsi:\n${ctx.textContent}`,
        };
      }

      if (toolCall.name === 'create_slide') {
        const { title, bullets, notes, layout } = toolCall.arguments || {};
        if (!title) return { success: false, error: 'Judul slide wajib diisi.' };
        if (driver.createSlide) {
          const res = await driver.createSlide(title, Array.isArray(bullets) ? bullets : [], notes, layout);
          return { success: true, result: `Slide #${res.slideNumber} ("${title}") berhasil dibuat.` };
        }
        await driver.addSlide(layout || 'TitleAndContent');
        await driver.insertSlideContent(title, Array.isArray(bullets) ? bullets : []);
        if (notes) await driver.setSpeakerNotes(notes);
        return { success: true, result: `Slide ("${title}") berhasil dibuat.` };
      }

      if (toolCall.name === 'create_presentation_deck') {
        const { slides } = toolCall.arguments || {};
        if (!Array.isArray(slides) || slides.length === 0) {
          return { success: false, error: 'Daftar slides tidak boleh kosong.' };
        }
        if (driver.createPresentationDeck) {
          const res = await driver.createPresentationDeck(slides);
          return { success: true, result: `${res.createdCount} slide presentasi berhasil dibuat ke PowerPoint.` };
        }
        for (const s of slides) {
          await driver.addSlide(s.layout || 'TitleAndContent');
          await driver.insertSlideContent(s.title, Array.isArray(s.bullets) ? s.bullets : []);
          if (s.notes) await driver.setSpeakerNotes(s.notes);
        }
        return { success: true, result: `${slides.length} slide presentasi berhasil dibuat.` };
      }

      if (toolCall.name === 'generate_themed_deck') {
        const { topic, theme, slides } = toolCall.arguments || {};
        if (!topic || !Array.isArray(slides) || slides.length === 0) {
          return { success: false, error: 'topic dan slides wajib diisi.' };
        }
        if (driver.generateThemedDeck) {
          const res = await driver.generateThemedDeck({ topic, theme, slides });
          return { success: true, result: `Deck bertema "${theme || 'corporate_blue'}" untuk topik "${topic}" berhasil dibuat (${res.createdCount} slide).` };
        }
        for (const s of slides) {
          const bullets = s.content || (s.metrics ? s.metrics.map((m: any) => `${m.label}: ${m.value}`) : []);
          await driver.addSlide(s.layout || 'TitleAndContent');
          await driver.insertSlideContent(s.title, bullets);
          if (s.notes) await driver.setSpeakerNotes(s.notes);
        }
        return { success: true, result: `Deck bertema "${theme || 'corporate_blue'}" berhasil dibuat (${slides.length} slide).` };
      }

      if (toolCall.name === 'transform_doc_to_deck') {
        const { documentText, targetSlideCount, theme, presentationTitle, targetAudience } = toolCall.arguments || {};
        const textToTransform = documentText || context.documentSummary || '';
        if (driver.transformDocToDeck) {
          const res = await driver.transformDocToDeck({
            documentText: textToTransform,
            targetSlideCount,
            theme,
            presentationTitle,
            targetAudience,
          });

          let output = `### 📊 Hasil Transformasi Dokumen ke Slide Presentasi: ${res.deckTitle}\n`;
          output += `Tema: ${res.appliedTheme} | Total Slide: ${res.totalSlidesCreated}\n\n`;

          const slideCards = res.slides.map((slide: any, idx: number) => {
            const bulletsStr = slide.bullets.map((b: string) => `  - ${b}`).join('\n');
            const talkingPoints = slide.speakerScript.keyTalkingPoints.map((tp: string) => `    * ${tp}`).join('\n');
            const metricsStr = slide.metrics && slide.metrics.length > 0
              ? `\n  - **Metrik:**\n` + slide.metrics.map((m: any) => `    * ${m.label}: ${m.value}${m.trend ? ` (${m.trend})` : ''}`).join('\n')
              : '';

            return `#### Slide ${idx + 1}: ${slide.title} [${slide.category.toUpperCase()}]\n` +
              `*Poin Slide:*\n${bulletsStr}${metricsStr}\n\n` +
              `🎙️ **Naskah Presenter (Speaker Notes)**:\n` +
              `- **Pembuka (Hook):** "${slide.speakerScript.hook}"\n` +
              `- **Poin Elaborasi:**\n${talkingPoints}\n` +
              `- **Transisi ke Slide Berikutnya:** "${slide.speakerScript.transition}"`;
          });

          output += slideCards.join('\n\n---\n\n');
          return { success: true, result: output.trim() };
        }
        return { success: false, error: 'Driver PowerPoint tidak mendukung transformDocToDeck.' };
      }

      if (toolCall.name === 'import_from_cross_app_hub') {
        const { snapshotId, theme, targetSlideCount, targetAudience } = toolCall.arguments || {};
        let snapshot = null;
        if (snapshotId) {
          const all = listCrossAppSnapshots();
          snapshot = all.find(s => s.id === snapshotId) || null;
        } else {
          snapshot = getLatestCrossAppSnapshot('PowerPoint');
        }

        if (!snapshot) {
          return { success: false, error: 'Tidak ditemukan snapshot data di Universal Hub dari aplikasi lain.' };
        }

        // Format snapshot into document text for doc-to-deck synthesis
        let docText = `Judul: ${snapshot.title}\nSumber: ${snapshot.sourceHost}\n`;
        if (snapshot.summaryText) {
          docText += `Ringkasan Eksekutif:\n${snapshot.summaryText}\n\n`;
        }
        if (snapshot.tableData && snapshot.tableData.headers) {
          docText += `Data Tabel:\n`;
          docText += snapshot.tableData.headers.join(' | ') + '\n';
          for (const row of snapshot.tableData.rows.slice(0, 10)) {
            docText += row.join(' | ') + '\n';
          }
          docText += '\n';
        }
        if (snapshot.metrics && snapshot.metrics.length > 0) {
          docText += `Metrik Kunci:\n`;
          for (const m of snapshot.metrics) {
            docText += `- ${m.label}: ${m.value}${m.trend ? ` (${m.trend})` : ''}\n`;
          }
        }

        if (driver.transformDocToDeck) {
          const res = await driver.transformDocToDeck({
            documentText: docText,
            targetSlideCount: targetSlideCount || 5,
            theme: theme || 'corporate_blue',
            presentationTitle: snapshot.title,
            targetAudience: targetAudience || 'executive',
          });

          let output = `### 📊 Hasil Transformasi Universal Hub ke Slide Presentasi: ${res.deckTitle}\n`;
          output += `Sumber: ${snapshot.sourceHost} Hub | Tema: ${res.appliedTheme} | Total Slide: ${res.totalSlidesCreated}\n\n`;

          const slideCards = res.slides.map((slide: any, idx: number) => {
            const bulletsStr = slide.bullets.map((b: string) => `  - ${b}`).join('\n');
            const talkingPoints = slide.speakerScript?.keyTalkingPoints
              ? slide.speakerScript.keyTalkingPoints.map((tp: string) => `    * ${tp}`).join('\n')
              : '';
            const metricsStr = slide.metrics && slide.metrics.length > 0
              ? `\n  - **Metrik:**\n` + slide.metrics.map((m: any) => `    * ${m.label}: ${m.value}${m.trend ? ` (${m.trend})` : ''}`).join('\n')
              : '';

            return `#### Slide ${idx + 1}: ${slide.title} [${slide.category.toUpperCase()}]\n` +
              `*Poin Slide:*\n${bulletsStr}${metricsStr}\n\n` +
              `🎙️ **Naskah Presenter (Speaker Notes)**:\n` +
              `- **Pembuka (Hook):** "${slide.speakerScript?.hook || ''}"\n` +
              `- **Poin Elaborasi:**\n${talkingPoints}\n` +
              `- **Transisi ke Slide Berikutnya:** "${slide.speakerScript?.transition || ''}"`;
          });

          output += slideCards.join('\n\n---\n\n');
          return { success: true, result: output.trim() };
        }

        return { success: false, error: 'Driver PowerPoint tidak mendukung transformDocToDeck.' };
      }

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

      if (toolCall.name === 'insert_process_flowchart') {
        const { textOrSteps, title, theme, direction } = toolCall.arguments || {};
        const steps = textOrSteps || context.documentSummary || '1. Mulai -> 2. Proses -> 3. Selesai';
        if (driver.insertProcessFlowchart) {
          const res = await driver.insertProcessFlowchart({
            textOrSteps: steps,
            title,
            theme,
            direction,
          });
          const diagramTitle = res.title || title || 'Diagram Alur Proses';
          return {
            success: true,
            result: `### 🔄 Diagram Alur Proses Berhasil Disisipkan ke Slide: ${diagramTitle}\n\n` +
              `- **Tema Presentasi:** ${res.appliedTheme}\n` +
              `- **Total Tahapan Alur:** ${res.nodeCount} node\n` +
              `- **Total Relasi:** ${res.edgeCount} koneksi\n\n` +
              `🎙️ **Naskah Pemateri (Speaker Script):**\n` +
              `*"Bapak/Ibu sekalian, slide visual ini memetakan alur proses '${diagramTitle}'. Terdapat ${res.nodeCount} tahapan kunci yang tersusun secara sistematis dari awal hingga evaluasi akhir guna menjamin operasional berjalan optimal."*`,
          };
        }
        return { success: false, error: 'Driver PowerPoint tidak mendukung insertProcessFlowchart.' };
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

      return { success: false, error: 'Tool tidak ditemukan.' };
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }
}
